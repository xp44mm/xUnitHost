/// STA 线程。
///
/// 这是**进程级单例**，故意做成模块而不是类：
///   模块的静态构造由 CLR 保证只跑一次，这份「唯一」是编译器给的，不靠人记得。
/// 要“显式”的只有关停：`shutdown`。
[<RequireQualifiedAccess>]
module OsHost.STA

open System
open System.Collections.Concurrent
open System.Runtime.ExceptionServices
open System.Threading
open System.Threading.Tasks

///一个待办队列
let private jobs = new BlockingCollection<unit -> unit>()

/// 起线程走「锁 + 字段」，不用 `Lazy`：`Lazy` 会把首次异常缓存住，
/// 一旦第一次 Start 失败（资源紧张 / 被拦），后面每次来都抬同一个旧异常，永远好不了。
/// 失败就保持 null，下次还有机会。
let private startGate = obj()

/// STA 线程：固定一条线程，open / exec / close 全在它上面跑，谁调用都不会串公寓。
let mutable private worker: Thread = null

let private pump () =
    for job in jobs.GetConsumingEnumerable() do
        job ()

let private startedThread () : Thread =
    lock startGate (fun () ->
        if isNull worker then
            let thread = Thread((fun () -> pump ()), IsBackground = true)

            thread.SetApartmentState ApartmentState.STA

            try
                thread.Start()
                worker <- thread
            with ex ->
                // 原样带栈往上抛，别包一层
                ExceptionDispatchInfo.Capture(ex).Throw()

        worker)

/// 0 = 开着，1 = 已关停（Interlocked 换，只关一次）
let mutable private closed = 0

let private refuse () : unit =
    raise (
        ObjectDisposedException(
            "OsHost.STA",
            "STA 线程已经 shutdown，不再接活；旧 RCW 随线程一起作废，要重来请整进程重起"
        )
    )

/// 排活：关停了就当场拒，别让 BlockingCollection 抛一句没头没尾的 InvalidOperationException
let private enqueue (job: unit -> unit) : unit =
    if Volatile.Read(&closed) = 1 then
        refuse ()

    try
        jobs.Add job
    with :? InvalidOperationException ->
        // 和 shutdown 撞上了：队列已经 CompleteAdding
        refuse ()

/// 排到固定线程上，把它的 Task 交回去 —— 逻辑只有这一份：
/// `runAsync` 原样给出去，`run` 在它外面加一道自等检查再阻塞等。
/// 强类型 TCS 直接捕获进闭包（不转手 `obj`：省一次装箱，也免掉一次运行时 unbox）。
/// RunContinuationsAsynchronously：await 的续体不许同步跑在那条 STA 线程上，
/// 否则续体里再调 run/runAsync 就会自己等自己。
let private schedule (work: unit -> 'T) : Task<'T> =
    if Volatile.Read(&closed) = 1 then
        refuse ()

    startedThread () |> ignore

    let pending =
        TaskCompletionSource<'T>(TaskCreationOptions.RunContinuationsAsynchronously)

    enqueue (fun () ->
        try
            pending.SetResult(work ())
        with ex ->
            pending.SetException ex)

    pending.Task

/// 已经在固定 STA 线程上还往队列里排 = 排队等自己，永久卡住。
/// 那个位置只该就地做：同步活儿直接调（或 `Sta.here`），要 await 的用 `Sta.runAsync` 把 Task 交回去。
let private refuseNesting () : unit =
    // 还没起过就是 null，自然「不在上面」
    let current = lock startGate (fun () -> worker)

    if not (isNull current) && Object.ReferenceEquals(Thread.CurrentThread, current) then
        raise (
            InvalidOperationException(
                "已经在固定 STA 线程上，改用 Sta.here 或 Sta.runAsync"
            )
        )

/// 长活：所有 COM 调用排在同一个 STA 线程上做。
/// host 是长活的：RCW 一旦跨公寓使用就要封送，而封送要求创建它的 STA 线程在泵消息 ——
/// 干脆固定一条线程，open / exec / close 全在它上面跑，谁调用都不会串公寓。
///
/// 就是 `runAsync` + 在调用方这里等。底下的「异步」是**线程亲和性**（活儿根本不能在
/// 调用线程上做），不是并发 I/O，所以阻塞是它的本体，不属于「sync-over-async 白占线程」。
/// **别在这条线程上的活儿里再调它**（自等）—— `refuseNesting` 当场报错。
let run (work: unit -> 'T) : 'T =
    refuseNesting ()
    (schedule work).GetAwaiter().GetResult()

/// 异步版：活儿仍排在**同一条固定 STA 线程**上，但立刻把 Task 交回调用方，
/// 不占调用线程、也不占一条线程池线程干等。
///
/// 这一条**允许从 STA 线程里调**（`... |> ignore` 发了就走）：它只入队、不等，不会自等 ——
/// EventPair 的退订就靠这个。但它也就成了「下一个活儿」，不是就地跑。
/// 注意：COM 对象依旧只在 STA 线程上创建 / 使用，await 之后要用它，仍得 run / runAsync。
let runAsync (work: unit -> 'T) : Task<'T> = schedule work

/// 短活：当前线程不是 STA 时，把活儿丢给一条临时 STA 线程跑完再回来
let here (work: unit -> 'T) : 'T =
    if Thread.CurrentThread.GetApartmentState() = ApartmentState.STA then
        work ()
    else
        let mutable result = Unchecked.defaultof<'T>
        let mutable failure: exn = null

        let body =
            ThreadStart(fun () ->
                try
                    result <- work ()
                with ex ->
                    failure <- ex
            )

        let worker = Thread(body)
        worker.SetApartmentState ApartmentState.STA
        worker.Start()
        worker.Join()

        if not (isNull failure) then
            raise failure

        result

/// 关停：把已经排上的活做完 → 固定线程退出。之后 `run` / `runAsync` 一律 ObjectDisposedException。
/// **不重启**：这条线程上创建过的 RCW 随线程一起作废，另起一条新的也接不回来 ——
/// 所以是一次性的，只给「这个进程归自己管」的场合收尾用。
/// 测试里别随手调，会毒掉同进程后面所有 Fact（要试就单开一个进程，见 `XUNITHOST_STA_SHUTDOWN`）。
let shutdown () : unit =
    let existing = lock startGate (fun () -> worker)

    if not (isNull existing) && Object.ReferenceEquals(Thread.CurrentThread, existing) then
        raise (InvalidOperationException("shutdown 不能在这条固定 STA 线程上等自己退出；请在排活的那一侧调"))

    if Interlocked.Exchange(&closed, 1) = 0 then
        jobs.CompleteAdding()

        if not (isNull existing) then
            existing.Join()
