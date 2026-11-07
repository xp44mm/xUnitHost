/// Excel 主机库：拿到一个 Excel Application，并且能干净地起一个、用一次、关掉。
/// 系统里已有 Excel 实例就用现成的（运行对象表里注册的那个），没有就新建一个。
///
/// 只负责 Excel 自己的事：取/建实例、认自己是谁、干净地关掉自己起的实例。
/// 线程、进程账本、窗口 → 进程号、COM 底层这些三家通用的活儿都在 `OsHost` 里。
namespace ExcelHost

open System
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open OsHost

[<RequireQualifiedAccess>]
module Host =

    /// Excel 的 ProgID 不带版本后缀：2016/2019/365 都注册成同一个名字（类型库都是 16）
    [<Literal>]
    let private ProgId = "Excel.Application"

    let private progIds = [ ProgId ]

    /// 有实例就返回那个实例，没有就新建一个并返回新建的。
    /// 新建的实例会补一张空白工作簿：一个工作簿都没有的自动化实例会自己退出。
    ///
    /// 这是**对外签约**用的：可以真的起一个进程。按地址回来（harness 内部交接）别用它 —— 用 `TryGetRunningApplication`。
    let GetOrCreateApplication () : Application =
        match Com.running progIds with
        | Some running -> running :?> Application
        | None ->
            let app = Com.create progIds :?> Application

            (try
                if app.Workbooks.Count = 0 then
                    app.Workbooks.Add() |> ignore
             with _ ->
                 ())

            app

    /// 只接现有的：ROT 里有就交出来，没有就是 None —— **绝不新建**（也就不补空白工作簿）。
    ///
    /// 「凭地址回到那台 host」走这条：地址说的人不在，就该老实报错，
    /// 不能因为它不在就顺手变一个出来（那等于把别人的会话顶掉，还认错了人）。
    let TryGetRunningApplication () : Application option =
        Com.running progIds |> Option.map(fun value -> value :?> Application)

    // ---------------- 黑板模式：按 GeneralAppInfo 挂回现成的那一只 ----------------

    /// 这家 host 自己的身份参数：原来 AppKind.image() / kindName() 里 Excel 那一支搬到专用 Host。
    /// 调用方填 GeneralAppInfo 的 Identity 就用它（进程名 / 显示名 / 种类标记三件套）。
    /// 旧路（AppInfo / AppKind）与并存的 Host.General 子模块都已撤掉，这里就是唯一的 GetApplication。
    let SelfIdentity : GeneralIdentity =
        GeneralIdentity.create ("EXCEL", "Excel", "excel")

    /// 这只 Excel 的进程号（Excel 的 Application 自己有 Hwnd，一问就到）
    let pidOf (app: Application) : int option =
        try
            Win32.pidOfWindow(nativeint app.Hwnd)
        with _ ->
            None

    /// 按插件抄下来的 GeneralAppInfo 精确挂回那一只 Excel 实例。
    /// 只附着、不新建、不 Quit、不改应用状态 —— 黑板的生命周期不归这儿管。
    /// 返回的 Application 请在 OsHost.Sta.run 那条固定 STA 线程上用（与附着同公寓）。
    /// 种类标记不符 / 信息过时 / ROT 里没有这个名字 / 认错进程，一律 Error。
    /// 文档级交接检查（换了工作簿 / 采集后存过盘等）按用户要求推迟，暂不做；
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

    // ---------------- 认自己是谁 ----------------
    // 地址里的指纹就是从这儿来的：哪个进程、锚定哪本工作簿、那本工作簿最后什么时候存的。

    /// 现场：进程号 + 锚定的工作簿（活动工作簿）+ 那个文件在磁盘上的最后写入时间 + 有没有未保存改动。
    /// 新工作簿（还没存过盘）路径是 None、保存时间是 0。
    let identity (app: Application) : Identity =
        let pid = pidOf app |> Option.defaultValue 0

        let path, dirty =
            try
                let book = app.ActiveWorkbook
                let saved = try book.Saved with _ -> true
                Paths.canonical book.FullName, not saved
            with _ ->
                // 一本工作簿都没有（或还没就绪）：没有锚点，也就不谈保存时间
                None, false

        { Pid = pid
          StartedMs = Proc.startedMs pid |> Option.defaultValue 0L
          DocPath = path
          SavedMs = Paths.savedMs path
          Dirty = dirty }

    // ---------------- 收尾 ----------------
    // Excel 跟 Word 不一样：Word 有 Quit(不保存) 这种参数，Excel 的 Quit() 没有。
    // 只要还有「改过没存」的工作簿，Quit 就会弹「是否保存 Book1？」，
    // 而且 DisplayAlerts=false 也不保证压得住 —— 自动化场景没人点按钮，调用方就卡在那儿。
    // 所以先逐本 Close(SaveChanges=false) 把改动丢掉，再 Quit。

    /// 关掉所有工作簿，改动统统丢弃（Close(SaveChanges=false)），返回关掉几本。
    /// 有本关不掉就停手；关完又被自动补一本的情况（有的配置会这样）也当停手，别死循环。
    let CloseAllWorkbooksWithoutSaving (app: Application) : int =
        let mutable closed = 0
        let mutable keepGoing = true
        let mutable lastCount = Int32.MaxValue

        while keepGoing && closed < 32 do
            let remaining =
                try
                    app.Workbooks.Count
                with _ ->
                    0

            if remaining <= 0 || remaining >= lastCount then
                keepGoing <- false
            else
                lastCount <- remaining

                let ok =
                    try
                        app.Workbooks.Item(remaining).Close(box false)
                        true
                    with _ ->
                        false

                if ok then
                    closed <- closed + 1
                else
                    keepGoing <- false

        closed

    /// 先关掉所有工作簿（丢弃改动）再 Quit —— 不弹保存对话框，调用方不会被模态框卡住。
    /// 只对「本程序自己起的实例」用；用户手上正开着表的实例别碰。
    let QuitWithoutSaving (app: Application) : unit =
        CloseAllWorkbooksWithoutSaving app |> ignore
        app.Quit()

    // ---------------- 方式一：一次命令（exec / execQuery） ----------------
    // 语义照参考项目 ExcelHost.Host.exec：**一次进程、一个实例、一段活**。
    // 改了一处：碰到「已有 Excel 在跑」不弹框求人退出，直接返回 Error。

    /// 后台跑，不给用户弹窗口；一切确认框都不弹
    let private quieten (app: Application) : unit =
        (try
            app.Visible <- false
         with _ ->
             ())

        (try
            app.DisplayAlerts <- false
         with _ ->
             ())

    /// 放掉 RCW：Excel 尤其明显 —— 客户端还捏着引用时，Quit 之后进程就是不肯退
    let private release (target: obj) : unit =
        (try
            Marshal.FinalReleaseComObject target |> ignore
         with _ ->
             ())

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    /// 起一个新的 Excel → 跑 work → 关掉退出。已有实例在跑就返回 Error（绝不碰用户的实例）。
    let private runOnce (work: Application -> 'T) : Result<'T, string> =
        Sta.here (fun () ->
            match Com.running progIds with
            | Some _ -> Error "本机已有 Excel 在运行：一次性命令要求它先退干净，否则会碰你手上的表"
            | None ->
                let before = Proc.ids "EXCEL"

                match
                    (try
                        Ok(Com.create progIds :?> Application)
                     with ex ->
                         Error(Errors.explain ex))
                with
                | Error message -> Error message
                | Ok app ->
                    let pids = Set.difference (Proc.ids "EXCEL") before
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
                    Proc.kill (Proc.waitGone "EXCEL" pids 10000)
                    outcome)

    /// 一次性执行，只关心成败（异常原因以 Error 返回，不向上抛）。
    let exec (action: Application -> unit) : Result<unit, string> = runOnce action

    /// 一次性执行并取回一个值。
    let execQuery (query: Application -> 'T) : Result<'T, string> = runOnce query
