/// PowerPoint 主机库：拿到一个 PowerPoint Application，并且能干净地起一个、用一次、关掉。
/// 系统里已有 PowerPoint 实例就用现成的（运行对象表里注册的那个），没有就新建一个。
///
/// 与 WordHost / ExcelHost 同形，只做 PowerPoint 自己的事：取/建实例、不弹框、收尾。
/// 线程、进程账本、COM 底层这些通用的活儿都在 `OsHost` 里。
namespace PowerPointHost

open System
open System.Runtime.InteropServices
open Microsoft.Office.Interop.PowerPoint
open OsHost

[<RequireQualifiedAccess>]
module Host =

    /// PowerPoint 的 ProgID 不带版本后缀：2016/2019/365 都注册成同一个名字（类型库都是 16）
    [<Literal>]
    let private ProgId = "PowerPoint.Application"

    let private progIds = [ ProgId ]

    /// 有实例就返回那个实例，没有就新建一个并返回新建的。不自动退出。
    /// 这是**对外签约**用的：可以真的起一个进程。按地址回来别用它 —— 用 TryGetRunningApplication。
    let GetOrCreateApplication () : Application =
        match Com.running progIds with
        | Some running -> running :?> Application
        | None -> Com.create progIds :?> Application

    /// 只接现有的：ROT 里有就交出来，没有就是 None —— **绝不新建**。
    let TryGetRunningApplication () : Application option =
        Com.running progIds |> Option.map(fun value -> value :?> Application)

    // ---------------- 黑板模式：按 info 挂回现成的那一只 ----------------

    /// 这只 PowerPoint 的进程号（Application.HWND 是主窗口句柄）
    let pidOf (app: Application) : int option =
        try
            Win32.pidOfWindow(nativeint app.HWND)
        with _ ->
            None

    /// 这家 host 自己的身份参数：原来 AppKind.image() / kindName() 里 PowerPoint 那一支搬到专用 Host。
    /// 调用方填 GeneralAppInfo 的 Identity 就用它（进程名 / 显示名 / 种类标记三件套）。
    let SelfIdentity : GeneralIdentity =
        GeneralIdentity.create ("POWERPNT", "PowerPoint", "ppt")

    /// 按插件抄下来的 GeneralAppInfo 精确挂回那一只 PowerPoint 实例。
    /// 只附着、不新建、不 Quit、不改应用状态 —— 黑板的生命周期不归这儿管。
    /// 返回的 Application 请在 OsHost.Sta.run 那条固定 STA 线程上用（与附着同公寓）。
    /// 文档级交接检查（换了演示文稿 / 采集后存过盘等）按用户要求推迟，暂不做；
    /// info 里的 DocName / DocPath 只是收下的数据，不参与判断。
    let GetApplication (info: GeneralAppInfo) : Result<Application, string> =
        GeneralAppInfo.attach info SelfIdentity.Tag progIds (fun raw ->
            try
                let app = raw :?> Application

                match pidOf app with
                | Some number when number <> info.Pid ->
                    Error(sprintf "ROT 名字指向的是 PID %d，不是黑板上的 %d" number info.Pid)
                | _ -> Ok app
            with ex ->
                Error(Errors.explain ex))

    // ---------------- 收尾 ----------------

    /// 逐份关掉演示文稿（Close() 无参数，不弹「是否保存」），返回关掉几份。
    /// 有关不掉就停手，别死循环。
    let CloseAllPresentations (app: Application) : int =
        let mutable closed = 0
        let mutable keepGoing = true
        let mutable lastCount = Int32.MaxValue

        while keepGoing && closed < 32 do
            let remaining =
                try
                    app.Presentations.Count
                with _ ->
                    0

            if remaining <= 0 || remaining >= lastCount then
                keepGoing <- false
            else
                lastCount <- remaining

                let ok =
                    try
                        app.Presentations.Item(remaining).Close()
                        true
                    with _ ->
                        false

                if ok then
                    closed <- closed + 1
                else
                    keepGoing <- false

        closed

    /// 先关掉所有演示文稿再 Quit —— 不弹保存对话框，调用方不会被模态框卡住。
    /// 只对「本程序自己起的实例」用；用户手上正开着的实例别碰。
    let QuitWithoutSaving (app: Application) : unit =
        CloseAllPresentations app |> ignore
        app.Quit()

    // ---------------- 方式一：一次命令（exec / execQuery） ----------------

    /// 一切确认框都不弹。
    /// 坑：PowerPoint **不允许**把 Application.Visible 设成 msoFalse 来「藏起来」；
    /// 不弹窗口靠 Presentations.Open / Add 的 WithWindow=msoFalse，不靠 Visible。
    let private quieten (app: Application) : unit =
        (try
            app.DisplayAlerts <- PpAlertLevel.ppAlertsNone
         with _ ->
             ())

    /// 放掉 RCW：进程外服务器要等客户端松开引用才肯真的退出。
    let private release (target: obj) : unit =
        (try
            Marshal.FinalReleaseComObject target |> ignore
         with _ ->
             ())

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    /// 起一个新的 PowerPoint → 跑 work → 关掉退出。已有实例在跑就返回 Error（绝不碰用户的实例）。
    let private runOnce (work: Application -> 'T) : Result<'T, string> =
        Sta.here (fun () ->
            match Com.running progIds with
            | Some _ -> Error "本机已有 PowerPoint 在运行：一次性命令要求它先退干净，否则会碰你手上的演示文稿"
            | None ->
                let before = Proc.ids "POWERPNT"

                match
                    (try
                        Ok(Com.create progIds :?> Application)
                     with ex ->
                         Error(Errors.explain ex))
                with
                | Error message -> Error message
                | Ok app ->
                    let pids = Set.difference (Proc.ids "POWERPNT") before
                    quieten app

                    let outcome =
                        try
                            Ok(work app)
                        with ex ->
                            Error(Errors.explain ex)

                    (try
                        QuitWithoutSaving app
                     with _ ->
                         ())

                    release app
                    Proc.kill (Proc.waitGone "POWERPNT" pids 10000)
                    outcome)

    /// 一次性执行，只关心成败（异常原因以 Error 返回，不向上抛）。
    let exec (action: Application -> unit) : Result<unit, string> = runOnce action

    /// 一次性执行并取回一个值。
    let execQuery (query: Application -> 'T) : Result<'T, string> = runOnce query
