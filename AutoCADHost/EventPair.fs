/// 一对「开始 / 结束」事件 → 一个 Task。
///
/// 出处：Notebook/ConcurrencyCookbook/08 Interop.md
///   · §8.2 Async Wrappers for "Begin/End" Methods —— 官方路子是 TaskFactory.FromAsync；
///   · §8.3 Async Wrappers for Anything —— FromAsync 只吃 APM（Begin 返回 IAsyncResult），
///     而 AutoCAD 的 BeginCommand / EndCommand 是**两条彼此独立的事件**，没有 IAsyncResult，
///     喂不进 FromAsync。这时就按 §8.3：TaskCompletionSource 自己当状态机。
///
/// 所以这里的形状是：
///   订阅 End → 订阅 Begin → 启动动作 → 等 End（或超时 / 取消）→ 退订 → Task 完成
///
/// 四条必须守住的规矩（缺一条就出鬼，都是这套代码里踩过的）：
///   1. **先订阅，后启动**：动作一发出就可能马上结束，订阅晚一步就漏掉那条 End，Task 永远挂住。
///   2. **委托一直被强引用**：COM 连接点只记弱引用，处理器被 GC 收走 = 订阅悄悄失效、事件不再来。
///      做法与 AcadBeginEnd.fs 一致 —— 返回的订阅对象自己握着委托。
///   3. **只完成一次**：End 可能来两条（并发命令），超时 / 取消也可能先到；谁先到谁算数。
///   4. **一定退订**：成功、超时、取消、启动抛异常，四条路都在 finish 里退订，不在用户实例上留连接点。
///
/// 这一层不认任何 COM 类型（AcadApplication 之类），所以 Word / Excel / 任意 .NET 事件都能套。
namespace AutoCADHost

open System
open System.Threading
open System.Threading.Tasks

/// 一次订阅。Dispose = 退订；同时它就是「委托的强引用持有者」。
/// 退订时必须交回**同一个委托实例**，所以这个对象不能随手丢掉。
type ISubscription =
    inherit IDisposable

/// 一对「开始 / 结束」事件源：挂上一个处理器、摘掉一个处理器，只有这两件事。
/// 谁来提供都行 —— COM 事件（AcadBeginEnd.fs）、.NET 的 IEvent（ofEvents）、或测试用的假源。
/// 两个泛型参数分别是 Begin / End 的参数类型；无参事件用 unit。
type EventPairSource<'Begin, 'End> =
    {
        /// 订阅「开始了」。返回的订阅对象要一直握着，Dispose 时退订。
        SubscribeBegin: ('Begin -> unit) -> IDisposable
        /// 订阅「结束了」。同上；它的参数就是 Task 的结果。
        SubscribeEnd: ('End -> unit) -> IDisposable
    }

[<RequireQualifiedAccess>]
module EventPair =

    /// 包装策略。默认值是对着「一条 AutoCAD 命令」这个最常见场景调的。
    type Options =
        {
            /// 必须先见到 Begin 才认 End。
            /// 防串台：上一条命令的 End、或别的动作触发的 End 混进来冒充本次结束。
            RequireBegin: bool
            /// 启动到收到 End 的时限；None = 不设时限（不推荐：调用方可能被永久挂住）。
            Timeout: TimeSpan option
            /// 调用方取消。取消后 Task 是 Canceled（不是 Faulted），与 TAP 约定一致。
            Cancellation: CancellationToken
            /// 退订怎么执行。COM 连接点必须回到挂它的那条公寓线程（STA），
            /// 所以留一个注入点；默认就地 Dispose。
            /// 注意别写成「阻塞等 STA」：End 可能是在 STA 线程上被重入送达的，
            /// 那时 init/finish 已经在 STA 上，再同步排一条 STA 活就是自己等自己 → 死锁。
            /// 要在 STA 上退订且可能从 STA 进来，用 `Sta.runAsync(...) |> ignore`（发了就走）。
            Unsubscribe: IDisposable -> unit
        }

    /// 默认策略：认 Begin、30 秒时限、不取消、就地退订。
    let defaults: Options =
        {
            RequireBegin = true
            Timeout = Some(TimeSpan.FromSeconds 30.0)
            Cancellation = CancellationToken.None
            Unsubscribe = fun subscription -> subscription.Dispose()
        }

    /// 把 .NET 的两条 IEvent 当成一对事件源（GUI 事件、自定义事件都能直接套）。
    let ofEvents
        (beginEvent: IEvent<'Begin>)
        (endEvent: IEvent<'End>)
        : EventPairSource<'Begin, 'End>
        =
        {
            SubscribeBegin =
                fun behavior ->
                    // handler 被下面的对象表达式捕获 → 与连接点同寿命
                    let handler = Handler<'Begin>(fun _ args -> behavior args)
                    beginEvent.AddHandler handler
                    { new IDisposable with
                        member _.Dispose() = beginEvent.RemoveHandler handler
                    }
            SubscribeEnd =
                fun behavior ->
                    let handler = Handler<'End>(fun _ args -> behavior args)
                    endEvent.AddHandler handler
                    { new IDisposable with
                        member _.Dispose() = endEvent.RemoveHandler handler
                    }
        }

    /// 核心：一对 Begin/End 事件 + 一个启动动作 → 一个 Task。
    /// 返回的 Task 只有在「End 真的到了」之后才完成；start 返回不等于事情做完。
    ///
    /// 线程：start 在**调用方线程**上跑（COM 场景下必须是挂着 RCW 的那条 STA）。
    /// End 在**事件送达的线程**上完成 Task（RunContinuationsAsynchronously，续体不会被塞进那条线程）。
    let runTaskWith
        (options: Options)
        (source: EventPairSource<'Begin, 'End>)
        (start: unit -> unit)
        : Task<'End>
        =

        let tcs =
            TaskCompletionSource<'End>(TaskCreationOptions.RunContinuationsAsynchronously)

        // 可变态全在这把锁下面：完成没有、Begin 到过没有、当前挂着哪几条订阅
        let gate = obj()
        let mutable completed = false
        let mutable began = false
        let mutable subscriptions: IDisposable list = []
        // 超时计时器 / 取消登记。这两样只在收尾时释放，且一律丢线程池去 Dispose：
        // 在 CancellationTokenRegistration 自己的回调里 Dispose 自己，会阻塞等回调返回 → 死锁。
        let cleanups = ResizeArray<IDisposable>()

        let disposeEventually (items: IDisposable list) =
            if not items.IsEmpty then
                ThreadPool.QueueUserWorkItem(fun _ ->
                    for item in items do
                        try
                            item.Dispose()
                        with _ ->
                            ()
                )
                |> ignore

        /// 收尾才登记：若这期间 Task 已经完成了，就直接排队释放，别攒着不处理
        let addCleanup (item: IDisposable) =
            let late =
                lock
                    gate
                    (fun () ->
                        if completed then
                            true
                        else
                            cleanups.Add item
                            false
                    )

            if late then
                disposeEventually [ item ]

        let unsubscribeAll () =
            let dropped =
                lock
                    gate
                    (fun () ->
                        let current = subscriptions
                        subscriptions <- []
                        current
                    )

            for subscription in dropped do
                // 退订失败不掩盖动作结果：连接点只在这次 execute 上有意义，摘不掉也只是脏一条
                try
                    options.Unsubscribe subscription
                with _ ->
                    ()

        /// 谁先到谁算数：只有第一个到达者能完成 Task，而且四条路都在这里统一退订
        let finish (complete: unit -> unit) =
            let first =
                lock
                    gate
                    (fun () ->
                        if completed then
                            false
                        else
                            completed <- true
                            true
                    )

            if first then
                let pending =
                    lock
                        gate
                        (fun () ->
                            let current = cleanups |> List.ofSeq
                            cleanups.Clear()
                            current
                        )

                disposeEventually pending
                unsubscribeAll()
                complete()

        let onBegin (_: 'Begin) =
            // Begin 的参数（命令名 / 文件名）目前只用来「证明这次真的开始过」；
            // 要拿它核对「是不是我这条命令」，在调用方套一层过滤更清楚。
            lock
                gate
                (fun () ->
                    if not completed then
                        began <- true
                )
            |> ignore

        let onEnd (args: 'End) =
            let accept =
                lock
                    gate
                    (fun () ->
                        if completed then
                            false
                        elif options.RequireBegin && not began then
                            false
                        else
                            true
                    )

            // 没见到 Begin 的 End：丢弃（按约定它不属于本次动作），继续等真的那条
            if accept then
                finish(fun () -> tcs.TrySetResult args |> ignore)

        // ── 1. 先订阅（End 先挂：连「Begin 之前就结束」这种极端顺序也不漏）
        let subscribed =
            try
                let endSubscription = source.SubscribeEnd onEnd
                lock gate (fun () -> subscriptions <- endSubscription :: subscriptions)
                let beginSubscription = source.SubscribeBegin onBegin
                lock gate (fun () -> subscriptions <- beginSubscription :: subscriptions)
                true
            with ex ->
                // 连接点可能拒接（AutoCAD 忙），一条都挂不上就别假装能等到通知
                finish(fun () -> tcs.TrySetException ex |> ignore)
                false

        if subscribed then
            // ── 2. 时限与取消（从「启动」开始算，不是从订阅开始算）
            match options.Timeout with
            | Some limit when limit > TimeSpan.Zero ->
                let timer =
                    new Timer(
                        TimerCallback(fun _ ->
                            finish(fun () ->
                                let seen = lock gate (fun () -> began)
                                let shown = if seen then "已到" else "未到"

                                tcs.TrySetException(
                                    TimeoutException(
                                        sprintf
                                            "等「结束」事件超时 %O：Begin %s（RequireBegin=%b）"
                                            limit
                                            shown
                                            options.RequireBegin
                                    )
                                )
                                |> ignore
                            )
                        ),
                        null,
                        limit,
                        Timeout.InfiniteTimeSpan
                    )

                addCleanup timer
            | _ -> ()

            if options.Cancellation.CanBeCanceled then
                let registration =
                    options.Cancellation.Register(fun () ->
                        finish(fun () -> tcs.TrySetCanceled() |> ignore)
                    )

                addCleanup(registration :> IDisposable)

            // ── 3. 启动动作：抛异常就是 faulted Task，并立刻退订
            if not(lock gate (fun () -> completed)) then
                try
                    start()
                with ex ->
                    finish(fun () -> tcs.TrySetException ex |> ignore)

        tcs.Task

    /// 默认策略版。
    let runTask
        (source: EventPairSource<'Begin, 'End>)
        (start: unit -> unit)
        : Task<'End>
        =
        runTaskWith defaults source start
