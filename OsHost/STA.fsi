/// STA 线程的对外样子（实现见 STA.fs）。
///
/// **这是进程级单例，故意做成模块而不是类**：
/// 状态只有「一条后台线程 + 一个待办队列」，物上就只该有一份 —— 再 new 一条 STA 线程
/// 不会多出什么好处，只会让 RCW 跨公寓（症状 RPC_E_WRONG_THREAD 或静默封送）。
/// 模块的静态构造由 CLR 保证只跑一次，这份「唯一」是编译器给的，不靠人记得。
///
/// 两种用法，别混：
///   · `run`  —— 长活：固定一条后台线程排队干活（host 开着的那段时间都用它）；
///   · `here` —— 短活：当前线程已经是 STA 就地跑，不是就临时起一条，跑完就收。
[<RequireQualifiedAccess>]
module OsHost.STA

open System.Threading.Tasks

/// 长活：所有 COM 调用排在同一个 STA 线程上做。
/// host 是长活的：RCW 一旦跨公寓使用就要封送，而封送要求创建它的 STA 线程在泵消息 ——
/// 干脆固定一条线程，open / exec / close 全在它上面跑，谁调用都不会串公寓。
///
/// 就是 `runAsync` + 在调用方这里等。底下的「异步」是线程亲和性（活儿根本不能在
/// 调用线程上做），不是并发 I/O，所以阻塞是它的本体，不属于「sync-over-async 白占线程」。
///
/// **别在这条线程上的活儿里再调它**（排队等自己，永久卡住）—— 当场抛 `InvalidOperationException`。
/// 那种场合同步就地调用，或者用 `here`；要 await 的活儿用 `runAsync` 把 Task 交回去。
val run: work: (unit -> 'T) -> 'T

/// 异步版：活儿仍排在**同一条固定 STA 线程**上，但立刻把 Task 交回调用方，
/// 不占调用线程、也不占一条线程池线程干等。异常由 TaskCompletionSource 置成 faulted，await 时原样抛出。
///
/// 这一条**允许从 STA 线程里调**（`... |> ignore` 发了就走）：它只入队、不等，不会自等 ——
/// EventPair 的退订就靠这个。但它也就成了「下一个活儿」，不是就地跑。
/// 注意：COM 对象依旧只在 STA 线程上创建 / 使用，await 之后要用它，仍得 run / runAsync。
val runAsync: work: (unit -> 'T) -> Task<'T>

/// 短活：当前线程不是 STA 时，把活儿丢给一条临时 STA 线程跑完再回来（每次一条，不复用）。
/// 已经是 STA 就地跑 —— 所以在这里面调 `run` 也一样是自等。
val here: work: (unit -> 'T) -> 'T

/// 关停：把已经排上的活做完 → 固定线程退出。之后 `run` / `runAsync` 一律 `ObjectDisposedException`。
///
/// **不重启**：这条线程上创建过的 RCW 随线程一起作废，另起一条新的也接不回来 ——
/// 所以是一次性的，只给「这个进程归自己管」的场合收尾用。幂等。
/// 测试里别随手调，会毒掉同进程后面所有 Fact（要试就单开一个进程，见 `XUNITHOST_STA_SHUTDOWN`）。
val shutdown: unit -> unit
