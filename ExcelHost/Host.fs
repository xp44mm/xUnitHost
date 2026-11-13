/// Excel 主机库：拿到一个 Excel Application，并且能干净地起一个、用一次、关掉。
/// 系统里已有 Excel 实例就用现成的（运行对象表里注册的那个），没有就新建一个。
///
/// 只负责 Excel 自己的事：取/建实例、认自己是谁、干净地关掉自己起的实例。
/// 线程、进程账本、窗口 → 进程号、COM 底层这些三家通用的活儿都在 `OsHost` 里。
namespace ExcelHost

open System
open System.Runtime.InteropServices
open Microsoft.Office.Core
open Microsoft.Office.Interop.Excel
open OsHost

[<RequireQualifiedAccess>]
module Host =

    /// Excel 的 ProgID 不带版本后缀：2016/2019/365 都注册成同一个名字（类型库都是 16）
    [<Literal>]
    let private ProgId = "Excel.Application"

    /// ROT 直取 / 一次性命令 / 起或取（`ExcelTask.getOrCreateOnSta`）共用这一份 ProgID
    let progIds = [ ProgId ]

    /// 只接现有的：ROT 里有就交出来，没有就是 None —— **绝不新建**（也就不补空白工作簿）。
    ///
    /// 「凭地址回到那台 host」走这条：地址说的人不在，就该老实报错，
    /// 不能因为它不在就顺手变一个出来（那等于把别人的会话顶掉，还认错了人）。
    let TryGetRunningApplication () : Application option =
        Com.running progIds |> Option.map(fun value -> value :?> Application)

    // ---------------- 黑板模式：按进程号挂回现成的那一只 ----------------

    /// 这家 host 的进程名：Proc.ids 的口径，不带 .exe（checkPid 用它核实「确实是 EXCEL」）。
    /// 黑板只把「进程号」交过来；进程名以这里为准（checkPid 用它核实）。
    let SelfImage : string = "EXCEL"

    /// 这只 Excel 的进程号（Excel 的 Application 自己有 Hwnd，一问就到）
    let pidOf (app: Application) : int option =
        try
            Win32.pidOfWindow(nativeint app.Hwnd)
        with _ ->
            None

    // ---------------- 按进程号挂回（判定链） ----------------
    // 能唯一确定一只实例的只有 Pid，所以这是唯一的挂回入口。
    // 与 AutoCADHost.Application 同一口径：
    //   · 只认现成对象 —— 绝不 CoCreateInstance、绝不 GetOrCreateApplication、绝不 Quit；
    //   · 认人靠 `Application.Hwnd` → PID 硬核对，对不上就是 None，不猜；
    //   · PID 闸门是**同步**做的：不合格立刻交回已完成且 faulted 的 Task，连 STA 队列都不排。
    // 隐藏跑的自动化实例 Hwnd 可能是 0，认不出就不认 —— 这种连 ROT 名也帮不上，只报错不新建。
    //
    // 成品 Task 入口在 ExcelTask.fs（GetApplicationAsync）；这里的 checkPid / findOnSta / unwrap
    // 就是它走下去的那条链，所以公开 —— 除它之外没有别的调用方。

    /// obj → Application，并用 Hwnd → PID 硬核对。不是 Excel / pid 对不上一律 None。
    let private matchApp (pid: int) (raw: obj) : Application option =
        try
            let app = raw :?> Application

            if obj.ReferenceEquals(app, null) then
                None
            elif pidOf app = Some pid then
                Some app
            else
                None
        with _ ->
            None

    /// 只认现成对象：ProgID 直取没对上就扫一遍 ROT，都没有就是 Error。
    /// 必须在 Sta.run / Sta.runAsync 那条固定 STA 线程上调用（由 ExcelTask.GetApplicationAsync 包好）。
    let findOnSta (pid: int) : Result<Application, string> =
        match Com.running progIds |> Option.bind(matchApp pid) with
        | Some app -> Ok app
        | None ->
            let entries = Com.runningEntries ()

            match entries |> List.tryPick(fun (_, raw) -> matchApp pid raw) with
            | Some app -> Ok app
            | None ->
                Error(
                    sprintf
                        "ROT 里没有 PID %d 的 Excel 自动化对象（ProgID 直取没对上，另扫了 %d 条登记项）：可能是启动中 / 隐藏实例没在主窗口上登记 / 尚未注册；绝不新建"
                        pid
                        entries.Length
                )

    /// 前置闸门：PID 有效，且确实是正在运行的 EXCEL 进程。纯查进程表，不碰 COM —— 自检能离线跑到这一层。
    let checkPid (pid: int) : Result<unit, string> =
        if pid <= 0 then
            Error "PID 必须大于 0：不猜、不新建"
        elif not (Process.ids SelfImage |> Set.contains pid) then
            Error(sprintf "PID %d 不是正在运行的 %s 进程：拒绝挂回，绝不新建" pid SelfImage)
        else
            Ok()

    /// 挂不回就是 faulted Task，await 时抛 InvalidOperationException（外层不再包 Result）
    let unwrap (result: Result<'T, string>) : 'T =
        match result with
        | Ok value -> value
        | Error message -> raise (InvalidOperationException(message) :> exn)

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
          StartedMs = Process.startedMs pid |> Option.defaultValue 0L
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

    /// 转换专用：不启用宏。带 Workbook_Open 的老 .xls 会弹宏提示，把 COM 调用挂住；
    /// ForceDisable 让宏一律不跑，判宏仍只看 HasVBProject。
    /// 返回原值，收尾 restore 回去 —— 即便挂回用户实例也不留痕。
    module Security =
        let apply (app: Application) : MsoAutomationSecurity option =
            let previous =
                try
                    Some app.AutomationSecurity
                with _ ->
                    None

            (try
                app.AutomationSecurity <- MsoAutomationSecurity.msoAutomationSecurityForceDisable
             with _ ->
                 ())

            previous

        let restore (app: Application) (previous: MsoAutomationSecurity option) : unit =
            match previous with
            | Some value ->
                (try
                    app.AutomationSecurity <- value
                 with _ ->
                     ())
            | None -> ()

    /// 后台跑，不给用户弹窗口；一切确认框都不弹。
    /// 返回 AutomationSecurity 原值，runOnce 在活干完后恢复。
    let private quieten (app: Application) : MsoAutomationSecurity option =
        (try
            app.Visible <- false
         with _ ->
             ())

        (try
            app.DisplayAlerts <- false
         with _ ->
             ())

        (try
            app.AskToUpdateLinks <- false
         with _ ->
             ())

        Security.apply app

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
        STA.here (fun () ->
            match Com.running progIds with
            | Some _ -> Error "本机已有 Excel 在运行：一次性命令要求它先退干净，否则会碰你手上的表"
            | None ->
                let before = Process.ids "EXCEL"

                match
                    (try
                        Ok(Com.create progIds :?> Application)
                     with ex ->
                         Error(Errors.explain ex))
                with
                | Error message -> Error message
                | Ok app ->
                    let pids = Set.difference (Process.ids "EXCEL") before
                    let previousSecurity = quieten app

                    let outcome =
                        try
                            Ok(work app)
                        with ex ->
                            Error(Errors.explain ex)

                    // 另存完把宏安全级别恢复成起实例时的值，不留痕
                    Security.restore app previousSecurity

                    (try
                        QuitWithoutSaving app
                     with _ ->
                         ())

                    release app
                    Process.kill (Process.waitGone "EXCEL" pids 10000)
                    outcome)

    /// 一次性执行，只关心成败（异常原因以 Error 返回，不向上抛）。
    let exec (action: Application -> unit) : Result<unit, string> = runOnce action

    /// 一次性执行并取回一个值。
    let execQuery (query: Application -> 'T) : Result<'T, string> = runOnce query
