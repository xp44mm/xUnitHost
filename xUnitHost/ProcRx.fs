/// System.Reactive 版的两条命令：把「起 / 关进程」编排成一条事件流。
///
/// 进程生命周期本来就是流：Started → Closed → Exited → OnCompleted。
/// 这里只把 OsHost 那条同步命令（Result<Launched> / Result<Closed>）包成 IObservable，
/// 用 SelectMany 把「拿到地址 → 派发 close」接起来，Timeout 防挂死，Disposable 收句柄。
/// 不认任何 Office 类型：Excel 的 COM 策略由调用方注入（CommandTest.fs 的 Dispatch）。
///
/// 与示范代码的两点不同（项目铁律）：
///   · 不用 Process.Start "excel.exe" —— 那会复用用户手上已开的实例；
///   · close 只按地址认自己起的那只，绝不「给所有 excel 进程发 CloseMainWindow」。
///
/// 顺序上的一个讲究：退出检测放在 **close 那一侧**，且必须在 close 命令执行前就订阅
/// （Process.Exited 事件要先注册好，等会儿才收得到），回信 Closed 先发、Exited 后补。
/// 放在 start 一侧会出两个毛病：轮询误判进程已退，Exited 反而抢在 Closed 前面到。
namespace XUnitHost

open System
open System.Diagnostics
open System.Reactive.Disposables
open System.Reactive.Linq
open System.Threading
open OsHost

/// 事件流里的一条「回信」：就是命令带回来的真实值
type ProcessEvent =
    /// start 命令带回来的现场（地址 / pid / 启动时间 / 起前现场）
    | Started of Launched
    /// close 命令带回来的结果
    | Closed of OsHost.Closed
    /// 进程真的退出了：Process.Exited 事件，或轮询兜底（via 记下是哪条通道报的）
    | Exited of pid: int * exitCode: int option * via: string
    /// 命令被闸门拒接 / 启动没成功
    | Refused of reason: string

[<RequireQualifiedAccess>]
module ProcRx =

    let private alive (pid: int) =
        try
            (Process.GetProcessById pid).Id = pid
        with _ ->
            false

    /// Process.Exited（C# 事件）→ IObservable。挂不上（权限 / 进程已经走了）就静默，
    /// 由轮询兜底；退订时只摘事件、放句柄，**绝不主动 kill 进程**。
    let private exitedEvents (pid: int) : IObservable<ProcessEvent> =
        Observable.Create<ProcessEvent>(
            Func<IObserver<ProcessEvent>, IDisposable>(fun observer ->
                let proc =
                    try
                        let p = Process.GetProcessById pid
                        try
                            p.EnableRaisingEvents <- true
                            Some p
                        with _ ->
                            p.Dispose()
                            None
                    with _ ->
                        None

                let handler =
                    match proc with
                    | Some p ->
                        let h =
                            EventHandler(fun _ _ ->
                                let code = try Some p.ExitCode with _ -> None
                                try
                                    observer.OnNext(Exited(pid, code, "Process.Exited"))
                                    observer.OnCompleted()
                                with _ ->
                                    ())

                        p.Exited.AddHandler h
                        Some h
                    | None -> None

                Disposable.Create(
                    Action(fun () ->
                        match proc, handler with
                        | Some p, Some h ->
                            (try p.Exited.RemoveHandler h with _ -> ())
                            (try p.Dispose() with _ -> ())
                        | Some p, None -> (try p.Dispose() with _ -> ())
                        | _ -> ())
                )))

    /// 第一条命令（Rx 版）：订阅时派发同步 start，送完 Started 就 Complete。
    /// 进程退出不在这里等 —— 那是 close 那一段的事。
    let start (command: unit -> Result<Launched, string>) : IObservable<ProcessEvent> =
        Observable.Create<ProcessEvent>(
            Func<IObserver<ProcessEvent>, IDisposable>(fun observer ->
                match command () with
                | Error message -> observer.OnNext(Refused message)
                | Ok launched -> observer.OnNext(Started launched)
                observer.OnCompleted()
                Disposable.Empty))

    /// 第二条命令（Rx 版）：订阅时**先**注册退出检测，再派发同步 close；
    /// Closed 先送出去，等进程真的退了再补 Exited，最后 OnCompleted。
    /// 退出检测两条通道 Merge：Process.Exited 事件优先，Interval 轮询兜底，Take(1) 谁先到算谁。
    let close
        (address: ProcAddress)
        (command: string -> Result<OsHost.Closed, string>)
        (pollMs: int)
        : IObservable<ProcessEvent> =
        Observable.Create<ProcessEvent>(
            Func<IObserver<ProcessEvent>, IDisposable>(fun observer ->
                let pid = address.Pid
                let arrived = ref (None: ProcessEvent option)
                use signal = new ManualResetEventSlim(false)

                let poll =
                    Observable.Interval(TimeSpan.FromMilliseconds(float pollMs))
                        .Where(fun (_: int64) -> not (alive pid))
                        .Select(fun (_: int64) -> Exited(pid, None, "轮询兜底"))

                let watch = Observable.Merge(exitedEvents pid, poll).Take(1)

                // 关键：close 命令执行前就订阅，Process.Exited 的 handler 此刻已经挂上
                use watchSub =
                    watch.Subscribe(
                        Action<ProcessEvent>(fun ev ->
                            lock arrived (fun () -> arrived.Value <- Some ev)
                            signal.Set())
                    )

                match command (Address.format address) with
                | Ok c -> observer.OnNext(Closed c)
                | Error message -> observer.OnNext(Refused message)

                // 回信先出去；退出事件可能已经在命令执行期间到了，这里补发（顺序严格）
                signal.Wait(TimeSpan.FromSeconds 5.) |> ignore

                match arrived.Value with
                | Some ev -> observer.OnNext ev
                | None -> observer.OnNext(Exited(pid, None, "5 秒内没等到退出事件"))

                observer.OnCompleted()
                Disposable.Empty))

    /// 顺序编排：Started 一到，隔 settle 才派发 close（示范里「启动完等一会儿再下一步」）。
    /// close 子流自己负责把 Closed → Exited 送完；start 流送完 Started 就 Complete，
    /// SelectMany 会等子流完成才整体 OnCompleted。Timeout 卡住就 OnError，不会把测试挂死。
    let flow
        (timeout: TimeSpan)
        (settle: TimeSpan)
        (startObs: IObservable<ProcessEvent>)
        (closeFor: ProcAddress -> IObservable<ProcessEvent>)
        : IObservable<ProcessEvent> =
        startObs
            .SelectMany(fun ev ->
                match ev with
                | Started launched ->
                    let delayedClose =
                        Observable.Timer(settle)
                            .SelectMany(fun (_: int64) -> closeFor launched.Address)

                    Observable.Merge(Observable.Return ev, delayedClose)
                | other -> Observable.Return other)
            .Timeout(timeout)
