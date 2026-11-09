/// AutoCAD 的「就绪订阅」：在把动作送进 AutoCAD 的那个入口挂上 COM 事件订阅，
/// 由 AutoCAD 自己通知「这一步做完了 / 我空闲了」，才让动作 Task 完成、才允许发下一条。
///
/// 起因：Documents.Add() / Open() 这类 COM 调用返回时，AutoCAD 往往还在切窗口、重绘，
/// 紧接着的下一条 COM 调用会撞 0x80010001 RPC_E_CALL_REJECTED。与其在调用侧反复重试，
/// 不如在入口订阅 AutoCAD 的完成通知，等它说「完了、我空了」再继续。
///
/// 订阅源（都在 COM 侧，跨进程可订阅，不需要改 AutoCAD 里的插件）：
///   · _DAcadApplicationEvents_Event 的 EndOpen / NewDrawing / EndCommand / EndModal ——
///     AutoCAD 主动推的「这一步结束了」通知。「打开一个 dwg 后 task 才 completed」对应的
///     就是 EndOpen。
///   · AcadState.IsQuiescent —— Autodesk 官方 ready 判据（idle and ready to accept
///     automation calls from out-of-process applications）。
/// 判据沿用 gotchas.md 的冷启动经验：连续 5 次 IsQuiescent=true，且距最近一次「动静」
/// （一条结束通知，或一次 false）静置 ≥ 1.5 秒，才算 ready。事件只负责把计时归零，
/// 最终放不放行仍由 AutoCAD 自己的 quiescent 状态说了算 —— 这样就算某个动作不触发事件
/// （读个属性之类），Task 也不会永远挂住。
namespace AutoCADHost

open System
open System.Diagnostics
open System.Threading
open System.Threading.Tasks
open Autodesk.AutoCAD.Interop
open OsHost

[<RequireQualifiedAccess>]
module Ready =

    /// 单调时钟（毫秒）。netstandard2.0 没有 Environment.TickCount64，用 Stopwatch 自己算。
    let private nowMs () : int64 =
        int64 (float (Stopwatch.GetTimestamp()) / (float Stopwatch.Frequency / 1000.0))

    /// 默认判据：连中 5 次 + 静置 1.5 秒（references/gotchas.md 的冷启动经验）。
    let StableSamples = 5
    let QuietMs = 1500L
    let PollMs = 200
    /// 等「一条完成通知」的默认时限；超时是 faulted Task，不无限挂住调用方
    let SignalTimeoutMs = 30000

    /// 纯判定器：不碰 COM，只把「通知 / 采样」喂进来，告诉我们此刻算不算 ready。
    /// 特意抽成公开类型并允许注入时钟，好在不连 AutoCAD 的情况下单测「Task 何时完成」这条规则。
    type Rule(stableSamples: int, quietMs: int64, clock: unit -> int64) =
        let gate = obj ()
        let mutable streak = 0
        let mutable lastChangeMs = clock ()
        let mutable notices: (string * string) list = []

        /// AutoCAD 来了一条通知：记下它，并把静置计时归零（连中次数也要重新攒）
        member _.Signal(label: string, detail: string) =
            lock gate (fun () ->
                lastChangeMs <- clock ()
                streak <- 0
                notices <- (label, detail) :: notices)

        /// 采一次 IsQuiescent
        member _.Sample(quiescent: bool) =
            lock gate (fun () ->
                if quiescent then
                    streak <- streak + 1
                else
                    streak <- 0
                    lastChangeMs <- clock ())

        /// 这一刻能不能放行：连中够次数，且离最近一次动静够久
        member _.IsReady() =
            lock gate (fun () -> streak >= stableSamples && clock () - lastChangeMs >= quietMs)

        member _.Streak = lock gate (fun () -> streak)
        member _.IdleMs = lock gate (fun () -> clock () - lastChangeMs)
        member _.Notices = lock gate (fun () -> List.rev notices)

    let defaultRule () = Rule(StableSamples, QuietMs, nowMs)

    /// 一次订阅：把 AutoCAD 的完成通知接成 Task 的完成信号。
    /// 生命周期跟着一次 executeAsync：use 完自动退订，不在用户实例上留东西。
    type Gate internal (app: AcadApplication) =
        // AcadApplication 本身就继承 _DAcadApplicationEvents_Event，是向上转换（不是 QI 猜类型）
        let events: _DAcadApplicationEvents_Event = upcast app
        let rule = defaultRule ()
        let gate = obj ()
        let mutable disposed = false
        /// 一次性等待者：WaitSignalSinceAsync 把「下一条通知」接成一个 Task。
        /// 委托由本对象握着，连接点只记弱引用，不会被 GC 悄悄断掉。
        let mutable waiters: (string * string -> unit) list = []

        let signal (label: string) (detail: string) =
            rule.Signal(label, detail)
            let targets = lock gate (fun () -> let current = waiters in waiters <- []; current)
            targets |> List.iter(fun deliver -> try deliver (label, detail) with _ -> ())

        // 委托必须一直握在手里：连接点只记弱引用，被 GC 掉订阅就悄悄断了。
        let onEndOpen =
            _DAcadApplicationEvents_EndOpenEventHandler(fun (fileName: string) -> signal "EndOpen" fileName)

        let onNewDrawing =
            _DAcadApplicationEvents_NewDrawingEventHandler(fun () -> signal "NewDrawing" "")

        let onEndCommand =
            _DAcadApplicationEvents_EndCommandEventHandler(fun (command: string) -> signal "EndCommand" command)

        let onEndModal =
            _DAcadApplicationEvents_EndModalEventHandler(fun () -> signal "EndModal" "")

        /// 尽力订阅：AutoCAD 正忙时连接点也可能拒接，挂不上就退化成纯 IsQuiescent 观察；
        /// 把实际挂上的名字报出来，别假装四条都成。
        let subscriptions =
            let tryAdd (label: string) (add: unit -> unit) =
                try
                    add ()
                    Some label
                with _ ->
                    None

            [ tryAdd "EndOpen" (fun () -> events.add_EndOpen onEndOpen)
              tryAdd "NewDrawing" (fun () -> events.add_NewDrawing onNewDrawing)
              tryAdd "EndCommand" (fun () -> events.add_EndCommand onEndCommand)
              tryAdd "EndModal" (fun () -> events.add_EndModal onEndModal) ]
            |> List.choose id

        /// 采一次 AutoCAD 自己的就绪判据（排回固定 STA 线程读 RCW）
        let sampleQuiescent () : Task<bool> =
            Sta.runAsync(fun () ->
                try
                    (app.GetAcadState()).IsQuiescent
                with _ ->
                    false)

        /// 等到 AutoCAD「通知过、也确实空下来了」为止。超时是 faulted Task，不无限挂住调用方。
        member _.WhenReadyAsync(?timeoutMs: int) : Task<unit> =
            let limit = defaultArg timeoutMs 60000
            let clock = Stopwatch.StartNew()

            let rec wait () : Task<unit> =
                task {
                    let! quiescent = sampleQuiescent ()
                    rule.Sample(quiescent)

                    if rule.IsReady() then
                        return ()
                    elif clock.ElapsedMilliseconds > int64 limit then
                        return!
                            Task.FromException<unit>(
                                TimeoutException(
                                    sprintf
                                        "等 AutoCAD ready 超时 %dms：连中 %d 次 / 静置 %dms / 通知 %A；订阅 %A"
                                        limit
                                        rule.Streak
                                        rule.IdleMs
                                        rule.Notices
                                        subscriptions
                                )
                            )
                    else
                        do! Task.Delay(PollMs)
                        return! wait ()
                }

            wait ()

        /// 等到「基线之后」AutoCAD 主动推来的**下一条完成通知**。
        /// 与 WhenReadyAsync 的区别：这里必须真的收到过通知才算完成，光是 IsQuiescent 静置放行不算；
        /// 这正是「动作发出去 ≠ 活干完」的判据。基线之后动作执行期间就送达的通知也算，
        /// 不漏这一条。超时给 faulted Task（await 抛 TimeoutException），不无限挂住调用方。
        /// 返回 (通知名, 细节)：EndOpen 给文件名、EndCommand 给命令名，其余为空串。
        member _.WaitSignalSinceAsync(baseline: int, ?timeoutMs: int) : Task<string * string> =
            let limit = defaultArg timeoutMs SignalTimeoutMs

            let pending =
                TaskCompletionSource<string * string>(TaskCreationOptions.RunContinuationsAsynchronously)

            let deliver (hit: string * string) = pending.TrySetResult(hit) |> ignore

            let immediate =
                lock gate (fun () ->
                    let seen = rule.Notices

                    if seen.Length > baseline then
                        Some seen.[baseline]
                    else
                        waiters <- deliver :: waiters
                        None)

            match immediate with
            | Some hit -> Task.FromResult<string * string>(hit)
            | None ->
                let cts = new CancellationTokenSource()

                let registration =
                    cts.Token.Register(fun () ->
                        pending.TrySetException(
                            TimeoutException(
                                sprintf
                                    "等 AutoCAD 完成通知超时 %dms：订阅 %A；已收到 %A"
                                    limit
                                    subscriptions
                                    rule.Notices
                            )
                            :> exn
                        )
                        |> ignore)

                cts.CancelAfter limit

                pending.Task.ContinueWith(
                    fun (_: Task<string * string>) ->
                        registration.Dispose()
                        cts.Dispose()
                        lock gate (fun () ->
                            waiters <- waiters |> List.filter(fun f -> not (obj.ReferenceEquals(f, deliver))))
                )
                |> ignore

                pending.Task

        /// 提交一个动作，返回一个 Task。
        /// 关键：动作在 STA 上跑完**不算完** —— 还要等 AutoCAD 的完成通知 + quiescent 稳定，
        /// Task 才完成；调用方 await 到它，才允许发下一条命令。
        /// 这就把「Documents.Add() 返回了但 AutoCAD 还没接住」这段空窗挡在 Task 里面。
        member this.SubmitAsync(action: AcadApplication -> 'T) : Task<'T> =
            task {
                do! this.WhenReadyAsync()
                let! value = Sta.runAsync(fun () -> action app)
                do! this.WhenReadyAsync()
                return value
            }

        /// 实际挂上的事件（报告用）
        member _.Subscriptions = subscriptions

        /// 收到过的通知（报告用，能看到 EndOpen / EndCommand 之类真的来过）
        member _.Notices = rule.Notices

        interface IDisposable with
            member _.Dispose() =
                let go =
                    lock gate (fun () ->
                        if disposed then
                            false
                        else
                            disposed <- true
                            true)

                if go then
                    // 退订也排回当初挂连接点的那条 STA；失败不掩盖动作结果。
                    try
                        Sta.run(fun () ->
                            let drop (work: unit -> unit) =
                                try
                                    work ()
                                with _ ->
                                    ()

                            if List.contains "EndOpen" subscriptions then
                                drop (fun () -> events.remove_EndOpen onEndOpen)

                            if List.contains "NewDrawing" subscriptions then
                                drop (fun () -> events.remove_NewDrawing onNewDrawing)

                            if List.contains "EndCommand" subscriptions then
                                drop (fun () -> events.remove_EndCommand onEndCommand)

                            if List.contains "EndModal" subscriptions then
                                drop (fun () -> events.remove_EndModal onEndModal))
                    with _ ->
                        ()

    /// 在固定 STA 线程上挂订阅：连接点必须跟 RCW 同一个公寓。
    let subscribe (app: AcadApplication) : Gate = Sta.run(fun () -> new Gate(app))
