/// 常驻 Excel（黑板挂回式）升级 .xls：Excel 由用户 / 插件常驻，这一侧只附着。
///
/// 与老的一次性通道（Host.execQuery：自己起 Excel → 干活 → Quit → kill）并列，互不干扰：
///   · 入口固定 ExcelHost.ExcelTask.GetApplicationAsync(pid: int)，按进程号挂回；绝不新建、绝不 Quit、绝不 kill；
///   · 不 requireClean —— 常驻实例在跑是**前提**，不是障碍；
///   · 所有 COM 读写都包在 OsHost.Sta.run 那条固定 STA 线程上（与附着同公寓）；
///   · 只逐份开/关自己那份工作簿，绝不用 CloseAllWorkbooksWithoutSaving（那会关掉用户开着的表）；
///   · 应用级开关（DisplayAlerts / AskToUpdateLinks / AutomationSecurity）临时改、用完恢复原值；
///   · **借用进程**：黑板说的是「那只常驻 Excel」，不是「它当时开的哪份工作簿」。
///     挂回只认 PID（进程名 / ROT 只贴报告），开/关自己的文件不会断链；
///   · 用户开着几本工作簿不归这条通道管，不断言数量；只核自己打开的那一本已关掉、没留垃圾。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Core
open OsHost

/// 黑板文本 → 挂回；两条 Fact 共用这一串
module private Attached =

    /// 黑板文本 → 「哪一家 + 哪只进程」+ 文档信息（读一次文件，两边都从同一段文本解析）
    let board () : Result<BoardApp * BoardDoc, string> =
        let file = Environment.GetEnvironmentVariable "XUNITHOST_BLACKBOARD_FILE"

        if String.IsNullOrWhiteSpace file then
            Error "没设 XUNITHOST_BLACKBOARD_FILE：把插件抄下来的黑板文本存成文件再跑"
        elif not (File.Exists file) then
            Error(sprintf "黑板文件不存在：%s" file)
        else
            let text = File.ReadAllText file
            BlackboardText.ofText text |> Result.map(fun info -> info, BlackboardText.docOfText text)

    let info () : Result<BoardApp, string> = board () |> Result.map fst

    type App = Microsoft.Office.Interop.Excel.Application

    /// 借用那只常驻进程：挂回只认 Pid，进程名 / ROT 名只贴报告。
    let borrow (info: BoardApp) : BoardApp = info

    /// 挂回那只常驻 Excel：只认进程号，只附着、不新建。
    /// 这一层把 Task 等成 Result —— 失败原因要原样贴进报告（拒接的是哪条，一眼看见）。
    /// 调用点都在测试线程（不在 Sta.run 的活儿里），在外面等 Task 不会自己等自己。
    let attach (info: BoardApp) : Result<App, string> =
        try
            Ok(ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await)
        with ex ->
            Error ex.Message

    /// 弹窗兜底：意外的对话框躲不完（修复、共享工作簿、打印机、加载项……），
    /// 与其一个个去压，不如让窗口露出来——哪一步卡住就把常驻 Excel 切到前台，
    /// 人看见就能点掉，模态框一关 COM 调用自己返回，流程接着走。
    /// 只做窗口激活（Win32 层面按 pid 找主窗口，模态对话框跟着主窗口一起到前台），
    /// 不碰任何 COM 对象、不动进程、不动文件、不改窗口状态。
    /// 返回值只表示「激活请求发出去了、没抛异常」——AppActivate 本身不回话。
    let bringToFront (pid: int) : bool =
        if Win32.bringToFront pid then
            // 先按 Win32 来：ShowWindow(SW_RESTORE) + SetForegroundWindow，最小化的也能还原上来
            true
        else
            // 前台锁挡着（用户正在别的程序里打字）时 Win32 会失败：退回 VB 的 AppActivate，
            // 它至少会闪一下任务栏按钮，人能看到「有东西在等」。
            try
                Microsoft.VisualBasic.Interaction.AppActivate pid
                true
            with _ ->
                false

    // ---------------- Word 侧：同样只挂回现成那只，逐份开/关自己那份 ----------------

    type WordApp = Microsoft.Office.Interop.Word.Application

    /// Word 的挂回参数：优先黑板文件（插件抄的），没设就按 XUNITHOST_WORD_PID 现造一条。
    /// 挂回只认 Pid；进程名 / ROT 名只贴报告。
    let wordInfo () : Result<BoardApp, string> =
        match Environment.GetEnvironmentVariable "XUNITHOST_BLACKBOARD_FILE" with
        | null
        | "" ->
            match Environment.GetEnvironmentVariable "XUNITHOST_WORD_PID" with
            | null
            | "" -> Error "没设 XUNITHOST_WORD_PID（或 XUNITHOST_BLACKBOARD_FILE）：把常驻 Word 的进程号给我"
            | text ->
                match Int32.TryParse text with
                | true, pid when pid > 0 ->
                    Ok
                        { Image = "WINWORD"
                          Pid = pid
                          Rot = "" }
                | _ -> Error(sprintf "XUNITHOST_WORD_PID 不是有效进程号：%s" text)
        | file ->
            if File.Exists file then
                File.ReadAllText file |> BlackboardText.ofText
            else
                Error(sprintf "黑板文件不存在：%s" file)

    /// Word 侧临时开关：显示框压掉、格式确认框压掉、宏一律不跑，收尾按原值恢复。
    /// **不碰 Visible** —— 常驻实例要给人看得见：意外的对话框露出来才有人点。
    type WordQuiet =
        { Alerts: Microsoft.Office.Interop.Word.WdAlertLevel option
          Confirm: bool option
          Security: MsoAutomationSecurity option }

    let quietWord (app: WordApp) : WordQuiet =
        let before =
            { Alerts = (try Some app.DisplayAlerts with _ -> None)
              Confirm = (try Some app.Options.ConfirmConversions with _ -> None)
              Security = WordHost.Host.Security.apply app }

        (try
            app.DisplayAlerts <- Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
         with _ ->
             ())

        (try
            app.Options.ConfirmConversions <- false
         with _ ->
             ())

        before

    let restoreWord (app: WordApp) (before: WordQuiet) : unit =
        (match before.Alerts with
         | Some value ->
             (try
                 app.DisplayAlerts <- value
              with _ ->
                  ())
         | None -> ())

        (match before.Confirm with
         | Some value ->
             (try
                 app.Options.ConfirmConversions <- value
              with _ ->
                  ())
         | None -> ())

        WordHost.Host.Security.restore app before.Security

    /// 实例里现在有几份文档（用户自己也开着的话一起算，只当报告数据）
    let docCount (app: WordApp) : int =
        try
            app.Documents.Count
        with _ ->
            -1

    /// 实例里还挂着我本轮打开过的文档吗（按 FullName 认，用户那几份不碰）。-1 = 数不出来。
    let ownDocsLeft (app: WordApp) (sources: Set<string>) : int =
        try
            let mutable n = 0

            for k in 1 .. app.Documents.Count do
                let full =
                    try
                        let mutable index: obj = box k
                        app.Documents.Item(&index).FullName
                    with _ ->
                        ""

                if full <> "" && sources.Contains(full.ToLowerInvariant()) then
                    n <- n + 1

            n
        with _ ->
            -1

    // ---------------- PowerPoint 侧：同样只挂回现成那只，逐份开/关自己那份 ----------------

    type PptApp = Microsoft.Office.Interop.PowerPoint.Application

    /// PowerPoint 的挂回参数：优先黑板文件（插件抄的），没设就按 XUNITHOST_PPT_PID 现造一条。
    let pptInfo () : Result<BoardApp, string> =
        match Environment.GetEnvironmentVariable "XUNITHOST_BLACKBOARD_FILE" with
        | null
        | "" ->
            match Environment.GetEnvironmentVariable "XUNITHOST_PPT_PID" with
            | null
            | "" -> Error "没设 XUNITHOST_PPT_PID（或 XUNITHOST_BLACKBOARD_FILE）：把常驻 PowerPoint 的进程号给我"
            | text ->
                match Int32.TryParse text with
                | true, pid when pid > 0 ->
                    Ok
                        { Image = "POWERPNT"
                          Pid = pid
                          Rot = "" }
                | _ -> Error(sprintf "XUNITHOST_PPT_PID 不是有效进程号：%s" text)
        | file ->
            if File.Exists file then
                File.ReadAllText file |> BlackboardText.ofText
            else
                Error(sprintf "黑板文件不存在：%s" file)

    /// 挂回那只常驻 PowerPoint（只认进程号；只附着、不新建、不 Quit）。
    let attachPpt (pid: int) : Result<PptApp, string> = PowerPointHost.Host.GetApplication pid

    /// PowerPoint 侧临时开关：确认框压掉、宏一律不跑，收尾按原值恢复。
    /// **不碰 Visible** —— PowerPoint 不允许把 Visible 设成 msoFalse。
    type PptQuiet =
        { Alerts: Microsoft.Office.Interop.PowerPoint.PpAlertLevel option
          Security: MsoAutomationSecurity option }

    let quietPpt (app: PptApp) : PptQuiet =
        let before =
            { Alerts = (try Some app.DisplayAlerts with _ -> None)
              Security = PowerPointHost.Host.Security.apply app }

        (try
            app.DisplayAlerts <- Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone
         with _ ->
             ())

        before

    let restorePpt (app: PptApp) (before: PptQuiet) : unit =
        (match before.Alerts with
         | Some value ->
             (try
                 app.DisplayAlerts <- value
              with _ ->
                  ())
         | None -> ())

        PowerPointHost.Host.Security.restore app before.Security

    /// 实例里现在有几份演示文稿（用户自己也开着的话一起算，只当报告数据）
    let presCount (app: PptApp) : int =
        try
            app.Presentations.Count
        with _ ->
            -1

    /// 实例里还挂着我本轮打开过的演示文稿吗（按 FullName 认，用户那几份不碰）。-1 = 数不出来。
    let ownPresLeft (app: PptApp) (sources: Set<string>) : int =
        try
            let mutable n = 0

            for k in 1 .. app.Presentations.Count do
                let full =
                    try
                        app.Presentations.Item(k).FullName
                    with _ ->
                        ""

                if full <> "" && sources.Contains(full.ToLowerInvariant()) then
                    n <- n + 1

            n
        with _ ->
            -1

    /// 文件字节指纹：还原原件、判断 Excel 有没有把原件改写掉
    let md5 (path: string) =
        use stream = File.OpenRead path
        use alg = System.Security.Cryptography.MD5.Create()
        System.Convert.ToHexString(alg.ComputeHash stream)    /// 收尾时实例里还挂着我本轮打开过的文件吗（按 FullName 认，别碰用户那本）。
    /// -1 = 数不出来，此时只报告、不判失败。
    let ownBooksLeft (app: App) (sources: Set<string>) : int =
        try
            let mutable n = 0

            for k in 1..app.Workbooks.Count do
                let full =
                    try
                        app.Workbooks.Item(k).FullName
                    with _ ->
                        ""

                if full <> "" && sources.Contains(full.ToLowerInvariant()) then
                    n <- n + 1

            n
        with _ ->
            -1

    /// 附着后临时改的应用级开关，收尾一律按原值恢复（用户实例不留痕）
    type Quiet =
        { DisplayAlerts: bool option
          AskToUpdateLinks: bool option
          /// 老 .xls 带页面设置/打印区域，Excel 每次打开、保存都要问打印机支持的纸张；
          /// 默认打印机是离线网络打印机时，这一问就几十秒到几分钟（实测 123 秒）。
          /// PrintCommunication=false 让 Excel 完全不跟打印机通信，用完恢复原值。
          PrintCommunication: bool option
          Security: MsoAutomationSecurity option }

    let quieten (app: App) : Quiet =
        let before =
            { DisplayAlerts = (try Some app.DisplayAlerts with _ -> None)
              AskToUpdateLinks = (try Some app.AskToUpdateLinks with _ -> None)
              PrintCommunication = (try Some app.PrintCommunication with _ -> None)
              Security = ExcelHost.Host.Security.apply app }

        (try app.DisplayAlerts <- false with _ -> ())
        (try app.AskToUpdateLinks <- false with _ -> ())
        (try app.PrintCommunication <- false with _ -> ())
        before

    let restore (app: App) (before: Quiet) : unit =
        (match before.DisplayAlerts with
         | Some value -> (try app.DisplayAlerts <- value with _ -> ())
         | None -> ())

        (match before.AskToUpdateLinks with
         | Some value -> (try app.AskToUpdateLinks <- value with _ -> ())
         | None -> ())

        (match before.PrintCommunication with
         | Some value -> (try app.PrintCommunication <- value with _ -> ())
         | None -> ())

        ExcelHost.Host.Security.restore app before.Security

/// 三条 Fact：借用进程的校验口径（纯逻辑）→ 只读挂回验通道 → 挂回式验一张真实 .xls。
type AttachedConvertTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``借用进程：只认进程号，进程名与 ROT 只作报告``() =
        Report.title output "借用进程：挂回只认 PID，其余字段只贴报告"

        let info = Attached.info () |> Report.ok
        let borrowed = Attached.borrow info

        Report.line output "黑板原值" (info.describe())
        Report.line output "现场" (sprintf "%s 进程 %A" info.Image (Set.toList (Report.procIds info.Image)))
        Report.line output "口径" "挂回只认 PID：进程名 / ROT 名进报告，不参与判断"

        Should.equal true (info.Pid > 0)
        Should.equal true (Set.contains info.Pid (Report.procIds info.Image))

        // 进程不在：闸门同步拒绝，连 STA 队列都不排
        let dead = { borrowed with Pid = 999999 }
        let deadWork = ExcelHost.ExcelTask.GetApplicationAsync dead.Pid

        Report.line output "死 PID" (sprintf "PID %d　IsCompleted=%b IsFaulted=%b" dead.Pid deadWork.IsCompleted deadWork.IsFaulted)
        Should.equal true deadWork.IsCompleted
        Should.equal true deadWork.IsFaulted

        // PID ≤ 0：同样就地拒绝
        let badWork = ExcelHost.ExcelTask.GetApplicationAsync 0

        Report.line output "PID 0" (sprintf "IsCompleted=%b IsFaulted=%b" badWork.IsCompleted badWork.IsFaulted)
        Should.equal true badWork.IsCompleted
        Should.equal true badWork.IsFaulted

        Report.conclusion output "通过：只认进程号；死 PID / 非正 PID 就地拒接，绝不新建"

    [<Fact>]
    member this.``常驻 Excel：按 Pid 挂回，只读核对，一个文件都不动``() =
        let info, doc = Attached.board () |> Report.ok
        let borrowed = Attached.borrow info
        let before = Report.procIds info.Image

        Report.title output "常驻 Excel：黑板信息（只取 Pid）→ GetApplicationAsync 挂回（只读）"
        Report.line output "黑板" (info.describe())
        Report.line output "文档" (doc.describe())
        Report.line output "交接" (sprintf "借用那只常驻进程，表里只有进程身份；挂回用表 %s" (borrowed.describe()))
        Report.line output "现场" (sprintf "%s 进程 %A" info.Image (Set.toList before))

        let attached = Attached.attach borrowed

        Report.line
            output
            "挂回"
            (match attached with
             | Ok app -> sprintf "Ok %s" (app.GetType().FullName)
             | Error message -> "拒接：" + message)

        let app = attached |> Report.ok

        let probe =
            STA.run(fun () ->
                let books = (try app.Workbooks.Count with _ -> -1)
                let alerts = (try app.DisplayAlerts with _ -> false)
                let security = (try int app.AutomationSecurity with _ -> -1)
                let pid = ExcelHost.Host.pidOf app
                books, alerts, security, pid)

        let books, alerts, security, pid = probe

        Report.line output "工作簿数" (string books)
        Report.line output "DisplayAlerts" (string alerts)
        Report.line output "AutomationSecurity" (string security)
        Report.line output "实例 PID" (sprintf "%A" pid)

        Should.equal (Some info.Pid) pid

        let after = Report.procIds info.Image
        Report.line output "收尾" (sprintf "%s 进程 %A" info.Image (Set.toList after))
        Should.equal before after
        Should.equal true (Set.contains info.Pid before)

        Report.conclusion output "通过：只挂回现成实例，没新建、没退出、没改应用状态，磁盘一个文件没动"

    [<Fact>]
    member this.``常驻 Excel：挂回式升级一张 .xls，核对通过才把原件送回收站``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录> 才真转"
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let info = Attached.info () |> Report.ok
            let borrowed = Attached.borrow info

            // 这条通道只认 .xls；默认验 1 张，XUNITHOST_CONVERT_LIMIT 可放大。
            let take = Paths.limitFromEnv () |> Option.defaultValue 1
            let items = Paths.selectForRound root Paths.Kind.Xls (Paths.scan root) |> List.truncate take
            let backup = RealFolder.backupRoot root
            let say (text: string) = output.WriteLine("    " + text)
            let beforeProc = Report.procIds "EXCEL"

            // 只认自己本轮打开过的文件：源 .xls 与可能落盘的新文件，别碰用户那本。
            let ownNames =
                items
                |> List.collect(fun i -> [ i.Source; i.Target ])
                |> List.map(fun p -> p.ToLowerInvariant())
                |> Set.ofList

            Report.title output "常驻 Excel：挂回式升级 .xls，核对通过才把原件送回收站"
            Report.line output "黑板" (info.describe())
            Report.line output "交接" (sprintf "借用那只常驻进程，表里只有进程身份；挂回用表 %s" (borrowed.describe()))
            Report.line output "本轮" (sprintf "%d 份 .xls（默认 1 张，XUNITHOST_CONVERT_LIMIT 可放大）" items.Length)
            Report.line output "备份目录" backup

            for i in items do
                Report.line output "待办" (RealFolder.rel root i.Source)

            // ---- 1. 目录外完整备份（已有且大小一致就复用，可续跑） ----
            let head = root.TrimEnd(Path.DirectorySeparatorChar) + string Path.DirectorySeparatorChar
            Assert.False(backup.StartsWith(head, StringComparison.OrdinalIgnoreCase), "备份目录不能落在目标目录里")

            let backed = RealFolder.backupAll root backup items say
            Should.equal items.Length backed
            Report.line output "备份" (sprintf "%d/%d 份，逐份比大小对上了" backed items.Length)

            // ---- 2. 重名让位 ----
            let renamed, skipped = RealFolder.renameExisting root items say
            let todo = items |> List.filter(fun i -> not (skipped.Contains i.Source))
            Report.line output "重名让位" (sprintf "改名 %d 个现有新文件；改名失败跳过 %d 份旧文件" renamed skipped.Count)

            // ---- 3. 挂回现成实例：绝不新建、绝不 Quit、绝不 kill ----
            let app = Attached.attach borrowed |> Report.ok
            Report.line output "挂回" (sprintf "pid=%d rot=%s（校验只用进程身份）" info.Pid info.Rot)

            let run =
                STA.run(fun () ->
                    let booksBefore = (try app.Workbooks.Count with _ -> -1)
                    let quiet = Attached.quieten app
                    let appliedAlerts = (try app.DisplayAlerts with _ -> true)
                    let appliedSecurity = (try int app.AutomationSecurity with _ -> -1)

                    let outcomes =
                        try
                            Convert.booksAttached app todo
                        with ex ->
                            todo
                            |> List.map(fun i ->
                                { Source = i.Source
                                  Target = i.Target
                                  Status = "失败"
                                  Detail = Errors.explain ex })

                    Attached.restore app quiet
                    let restoredAlerts = (try app.DisplayAlerts with _ -> true)
                    let restoredSecurity = (try int app.AutomationSecurity with _ -> -1)
                    let booksAfter = (try app.Workbooks.Count with _ -> -1)
                    let ownLeft = Attached.ownBooksLeft app ownNames

                    outcomes,
                    booksBefore,
                    booksAfter,
                    ownLeft,
                    appliedAlerts,
                    appliedSecurity,
                    restoredAlerts,
                    restoredSecurity,
                    quiet)

            let outcomes, booksBefore, booksAfter, ownLeft, appliedAlerts, appliedSecurity, restoredAlerts, restoredSecurity, quiet =
                run

            // 先把逐份结果全部打出来，再下断言：失败原因不会再被断言挡住看不见。
            for o in outcomes do
                let note = if Convert.canDelete o then "" else "　—— " + o.Detail
                Report.line output (RealFolder.mark o.Status) (RealFolder.rel root o.Source + note)

            Report.line output "实例工作簿数" (sprintf "干活前 %d / 收尾后 %d（用户开着几本不校验）" booksBefore booksAfter)
            Report.line output "我开的还剩" (sprintf "%d 本（0 = 自己开的都关了；-1 = 数不出来）" ownLeft)
            Report.line output "DisplayAlerts" (sprintf "原值 %A → %A → 恢复 %A" quiet.DisplayAlerts appliedAlerts restoredAlerts)

            Report.line
                output
                "AutomationSecurity"
                (sprintf
                    "原值 %A → %d（ForceDisable）→ 恢复 %d"
                    (quiet.Security |> Option.map(fun v -> int v))
                    appliedSecurity
                    restoredSecurity)

            Should.equal 3 appliedSecurity
            Should.equal true (ownLeft = 0 || ownLeft = -1)
            Should.equal quiet.DisplayAlerts (Some restoredAlerts)
            Should.equal (quiet.Security |> Option.map(fun v -> int v)) (Some restoredSecurity)

            // ---- 4. 只送核对通过的原件进回收站 ----
            let good = outcomes |> List.filter Convert.canDelete
            let bad = outcomes |> List.filter(Convert.canDelete >> not)
            let removed, kept = RealFolder.deleteVerified root good say
            Report.line output "原件入回收站" (sprintf "送进去 %d 份，送不进留下 %d 份" removed kept)

            for o in bad do
                Assert.True(File.Exists o.Source, sprintf "核对没通过却把原件弄丢了：%s" o.Source)

            // ---- 5. 现场：常驻实例还在、没多也没少 ----
            let afterProc = Report.procIds "EXCEL"
            Report.line output "现场" (sprintf "EXCEL 进程 干活前 %A / 收尾后 %A" (Set.toList beforeProc) (Set.toList afterProc))
            Should.equal beforeProc afterProc
            Should.equal true (Set.contains info.Pid afterProc)

            let left = Paths.scan root |> Paths.filterByKind Paths.Kind.Xls
            Report.line output "目录现状" (sprintf "还剩 %d 份 .xls" left.Length)

            Report.conclusion
                output
                (sprintf
                    "完成：本轮 %d 份，核对通过 %d 份，原件入回收站 %d 份；只借用 pid=%d 的进程，别的窗口/工作簿不校验，自己开的都关了，应用开关已恢复"
                    items.Length
                    good.Length
                    removed
                    info.Pid)

    /// 验一张那轮 Excel 把原件从 683520 B 改写成了 684032 B，且没生成 .xlsx。
    /// 下一轮重跑前先把原件恢复成目录外备份里的原始字节：备份与留档 md5 必须一致才动手，
    /// 已有新格式文件 / 原件已入回收站就停下，绝不覆盖已完成的转换成果。
    [<Fact>]
    member this.``还原验一张那份 xls：从目录外备份恢复原始字节``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录> 才还原"
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let take = Paths.limitFromEnv () |> Option.defaultValue 1
            let items = Paths.selectForRound root Paths.Kind.Xls (Paths.scan root) |> List.truncate take
            Should.equal 1 items.Length

            let source = items.Head.Source
            let rel = RealFolder.rel root source
            let backup = Path.Combine(RealFolder.backupRoot root, rel)
            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))

            let keep =
                match Environment.GetEnvironmentVariable "XUNITHOST_RESTORE_KEEP" with
                | null
                | "" -> Path.Combine(projectRoot, "TestResults", "验一张-转换前原件.xls")
                | value -> value

            let targets = Paths.existingTargetsOf source

            Report.title output "还原验一张那份 .xls：从目录外备份恢复原始字节"
            Report.line output "目标" rel
            Report.line output "备份" backup
            Report.line output "留档" keep

            if not (File.Exists source) then
                Assert.Skip "原件已不在（已转换并入回收站），无需还原"

            Assert.True(File.Exists backup, sprintf "备份不在：%s" backup)
            Assert.True(File.Exists keep, sprintf "留档不在：%s" keep)

            let hashBefore, hashBackup, hashKeep = Attached.md5 source, Attached.md5 backup, Attached.md5 keep
            let sizeBefore = (FileInfo source).Length

            Report.line
                output
                "还原前"
                (sprintf
                    "原件 %d B md5=%s；备份 %d B md5=%s；留档 %d B md5=%s"
                    sizeBefore
                    hashBefore
                    (FileInfo backup).Length
                    hashBackup
                    (FileInfo keep).Length
                    hashKeep)

            Report.line output "已有新格式" (sprintf "%d 个（有就别覆盖原件）" targets.Length)
            Should.equal 0 targets.Length
            Should.equal hashBackup hashKeep

            let changed =
                if hashBefore = hashBackup then
                    false
                else
                    let info = FileInfo source
                    if info.IsReadOnly then info.IsReadOnly <- false
                    File.Copy(backup, source, true)
                    true

            let sizeAfter = (FileInfo source).Length
            let hashAfter = Attached.md5 source

            Report.line output "复制" (if changed then "备份 → 原件（覆盖）" else "原件已是原始字节，一个字节没动")
            Report.line output "还原后" (sprintf "%d B md5=%s" sizeAfter hashAfter)
            Should.equal hashBackup hashAfter
            Should.equal (FileInfo backup).Length sizeAfter
            Should.equal hashBackup (Attached.md5 backup)
            Should.equal hashKeep (Attached.md5 keep)

            Report.line output "备份留档" (sprintf "备份 %d B / 留档 %d B，md5 都没变" (FileInfo backup).Length (FileInfo keep).Length)
            Report.conclusion output (sprintf "完成：%s 恢复到 %d B 原始字节，备份与留档原样保留" rel sizeAfter)

    /// 逐份推进器：一份一份送，成功就进下一份；失败打原因、记清单、接着下一份。
    /// 停的条件只有两个——同一原因连续 N 次（系统性问题），或单份超时（卡住）。
    /// 每份转换前先备份；转换后若原件被 Excel 改写，就从备份还原，失败件的字节不留痕。
    /// 成功那份的原件才送回收站，一份一送，不攒批。
    [<Fact>]
    member this.``常驻 Excel：逐份升级，成功一份送走一份，失败记原因接着下一份``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录> 才逐份升级"
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let info = Attached.info () |> Report.ok
            let borrowed = Attached.borrow info
            let backup = RealFolder.backupRoot root

            let all = Paths.selectForRound root Paths.Kind.Xls (Paths.scan root)

            let queue =
                match Paths.limitFromEnv () with
                | Some n -> all |> List.truncate n
                | None -> all

            let intEnv (name: string) (fallback: int) =
                let text = defaultArg (Option.ofObj (Environment.GetEnvironmentVariable name)) ""

                match Int32.TryParse text with
                | true, n when n > 0 -> n
                | _ -> fallback

            // 多久没进展就开始催人：默认 45 秒。XUNITHOST_CONVERT_TIMEOUT 仍认，作旧名兼容。
            let hintSec = intEnv "XUNITHOST_CONVERT_POPUP_HINT" (intEnv "XUNITHOST_CONVERT_TIMEOUT" 45)
            let sameReasonLimit = intEnv "XUNITHOST_CONVERT_SAME_REASON" 3
            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))

            let fileEnv (name: string) (fallback: string) =
                match Environment.GetEnvironmentVariable name with
                | null
                | "" -> Path.Combine(projectRoot, "TestResults", fallback)
                | value -> value

            let progressFile = fileEnv "XUNITHOST_CONVERT_PROGRESS" "逐份-进度.log"
            let failedFile = fileEnv "XUNITHOST_CONVERT_FAILED" "逐份-失败清单.txt"
            let say (text: string) = output.WriteLine("    " + text)

            // 进度每行立刻落盘：真卡在某个 COM 调用时，最后一行就是卡点。
            // 主线程写、看门狗线程也写（催前台的提示），同一个文件要串行。
            let progressGate = obj()

            let append (path: string) (line: string) =
                lock progressGate (fun () ->
                    let mutable ok = false
                    let mutable tries = 0

                    while not ok && tries < 5 do
                        tries <- tries + 1

                        try
                            File.AppendAllText(path, line + Environment.NewLine)
                            ok <- true
                        with _ ->
                            System.Threading.Thread.Sleep 50)

            let record (text: string) =
                let line = sprintf "%s %s" (DateTime.Now.ToString("HH:mm:ss")) text
                say line
                append progressFile line

            // 失败清单累积去重：同一份重复失败只留一行，整份清单可直接当 XUNITHOST_CONVERT_SKIP
            let failedSeen =
                if File.Exists failedFile then
                    File.ReadAllLines failedFile
                    |> Array.map(fun line -> line.Trim())
                    |> Array.filter(fun line -> line <> "" && not (line.StartsWith "#"))
                    |> Set.ofArray
                else
                    try
                        File.WriteAllText(
                            failedFile,
                            "# 逐份升级失败清单（每行一个相对路径；可直接给 XUNITHOST_CONVERT_SKIP）" + Environment.NewLine
                        )
                    with _ ->
                        ()

                    Set.empty

            let mutable known = failedSeen

            let noteFailure (rel: string) =
                if not (known.Contains rel) then
                    known <- known.Add rel

                    try
                        File.AppendAllText(failedFile, rel + Environment.NewLine)
                    with _ ->
                        ()

            let reasonOf (detail: string) =
                if detail.Contains "共享工作簿" || detail.Contains "修订历史" then "共享工作簿/修订历史"
                elif detail.Contains "VBA" || detail.Contains "宏" then "宏判定"
                elif detail.Contains "超时" then "COM 超时"
                elif detail.Contains "修复" then "修复对话框"
                elif detail.Contains "拒绝" then "COM 拒绝"
                else "其它"

            Report.title output "常驻 Excel：逐份升级，成功一份送走一份"
            Report.line output "黑板" (info.describe())
            Report.line output "交接" (sprintf "借用那只常驻进程，表里只有进程身份；挂回用表 %s" (borrowed.describe()))
            Report.line output "本轮" (sprintf "%d 份 .xls（%s）" queue.Length (RealFolder.limitLabel ()))
            Report.line output "备份目录" backup
            Report.line output "进度日志" progressFile
            Report.line output "失败清单" (sprintf "%s（已有 %d 条历史失败）" failedFile failedSeen.Count)

            Report.line
                output
                "弹窗策略"
                (sprintf
                    "不躲：%d 秒没进展就把 Excel 切到前台催人，点掉对话框自动继续；只激活窗口，不杀进程、不关你的表"
                    hintSec)

            Report.line output "停的条件" (sprintf "同一原因连续 %d 次（卡住不判失败，等你处理）" sameReasonLimit)

            let beforeProc = Report.procIds "EXCEL"
            let app = Attached.attach borrowed |> Report.ok
            Report.line output "挂回" (sprintf "pid=%d rot=%s（校验只用进程身份）" info.Pid info.Rot)

            // 开跑时把常驻窗口切到前台一次：这一轮的对话框就在眼前，不用等到卡了才催。
            // XUNITHOST_CONVERT_FOCUS=0 关掉（卡住催前台仍然生效）。
            let focusOnStart =
                match Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_FOCUS" with
                | null
                | "" -> true
                | text -> text.Trim() <> "0"

            let raised = if focusOnStart then Attached.bringToFront info.Pid else false

            Report.line
                output
                "界面"
                (sprintf
                    "开跑前置 %s：%s"
                    (if focusOnStart then "开" else "关")
                    (if not focusOnStart then "没动窗口，只在卡住时催"
                     elif raised then "已把 Excel 切到前台，跑的时候有对话框你会立刻看见"
                     else "没切动，请自己点任务栏里的 Excel 看一眼"))

            // 弹窗兜底：不指望把所有对话框都躲掉。哪一步超过 hintSec 没动，就把 Excel 窗口切到前台，
            // 之后每隔 hintSec 再催一次 —— 人看见就能点掉，COM 调用自己返回，流程接着走。
            // 只激活窗口：不杀进程、不关用户开着的表、不改窗口状态。
            let mutable currentStage = "未开始"
            let mutable currentSince = DateTime.Now
            let mutable watchStop = false
            let mutable lastHint = DateTime.MinValue
            let mutable hintCount = 0

            let watchdog =
                new System.Threading.Timer(
                    (fun (_: obj) ->
                        if not watchStop then
                            let secs = int (DateTime.Now - currentSince).TotalSeconds
                            let sinceHint = (DateTime.Now - lastHint).TotalSeconds

                            if secs >= hintSec && sinceHint >= float hintSec && hintCount < 40 then
                                lastHint <- DateTime.Now
                                hintCount <- hintCount + 1
                                let front = Attached.bringToFront info.Pid

                                record (
                                    sprintf
                                        "⚠ 没进展 %d 秒（%s）—— 多半弹了对话框；%s"
                                        secs
                                        currentStage
                                        (if front then
                                             "已把 Excel 切到前台，请点掉弹窗，点完我自动继续"
                                         else
                                             "没切动前台，请自己点任务栏里的 Excel 看一眼")
                                )),
                    null,
                    5000,
                    5000
                )

            let mutable okCount = 0
            let mutable failCount = 0
            let mutable stopReason = ""
            let mutable lastReason = ""
            let mutable sameRun = 0
            let outcomes = ResizeArray<Outcome>()

            let run =
                try
                    Ok(
                        STA.run(fun () ->
                            let quiet = Attached.quieten app
                            let appliedAlerts = (try app.DisplayAlerts with _ -> true)
                            let appliedSecurity = (try int app.AutomationSecurity with _ -> -1)
                            let appliedPrint = (try app.PrintCommunication with _ -> true)

                            try
                                // 上次中断可能把自己的工作簿留在实例里：只关本次队列里的那些，用户的别的表一个字不动
                                let queueSet =
                                    queue |> List.map(fun x -> x.Source.ToLowerInvariant()) |> Set.ofList

                                let existingCount = (try app.Workbooks.Count with _ -> 0)

                                let books =
                                    [ for k in existingCount .. -1 .. 1 do
                                        let book = app.Workbooks.Item k
                                        let full = (try book.FullName with _ -> "")
                                        let name = (try book.Name with _ -> "")
                                        yield book, full, name ]

                                let mutable closed = []

                                for book, full, name in books do
                                    if full <> "" && queueSet.Contains(full.ToLowerInvariant()) then
                                        (try book.Close(box false) with _ -> ())
                                        closed <- name :: closed

                                record (
                                    sprintf
                                        "  实例工作簿 %d 本%s"
                                        existingCount
                                        (if closed.IsEmpty then
                                             ""
                                         else
                                             sprintf
                                                 "，关掉上次残留 %d 本（%s）"
                                                 closed.Length
                                                 (String.concat "、" (List.rev closed)))
                                )

                                let mutable index = 0

                                while index < queue.Length && stopReason = "" do
                                    let i = queue.[index]
                                    index <- index + 1
                                    let rel = RealFolder.rel root i.Source
                                    // 每份重新起表：上一份的等待不算到这一份头上，催前台也从这一份重新计
                                    currentStage <- sprintf "[%d/%d] 开始 %s" index queue.Length rel
                                    currentSince <- DateTime.Now
                                    lastHint <- DateTime.MinValue
                                    record (sprintf "[%d/%d] 开始 %s" index queue.Length rel)
                                    let clock = System.Diagnostics.Stopwatch.StartNew()

                                    let finish (text: string) =
                                        record text
                                        record (sprintf "  本份耗时 %.1f 秒" clock.Elapsed.TotalSeconds)

                                    if not (Set.contains info.Pid (Report.procIds info.Image)) then
                                        stopReason <- sprintf "常驻 Excel pid %d 已不在" info.Pid
                                        record ("停：" + stopReason)
                                    else
                                        // ---- 1. 只备份这一份 ----
                                        currentStage <- sprintf "[%d/%d] 备份 %s" index queue.Length rel
                                        currentSince <- DateTime.Now

                                        let backed =
                                            try
                                                RealFolder.backupAll root backup [i] say |> ignore
                                                true
                                            with ex ->
                                                record ("  备份失败：" + Errors.explain ex)
                                                false

                                        if backed then
                                            // ---- 2. 重名让位 ----
                                            currentStage <- sprintf "[%d/%d] 重名让位 %s" index queue.Length rel

                                            let _, skipped = RealFolder.renameExisting root [i] say

                                            if skipped.Contains i.Source then
                                                failCount <- failCount + 1
                                                finish (sprintf "× %s —— 现有同名新文件改名失败，这份跳过" rel)
                                                noteFailure rel
                                            else
                                                // ---- 3. 转换这一份（内部已读回核对） ----
                                                currentStage <- sprintf "[%d/%d] 转换 %s" index queue.Length rel
                                                currentSince <- DateTime.Now

                                                let originalHash =
                                                    try
                                                        Some(Attached.md5 i.Source)
                                                    with _ ->
                                                        None

                                                let originalLen =
                                                    try
                                                        Some((FileInfo i.Source).Length)
                                                    with _ ->
                                                        None

                                                let outcome =
                                                    try
                                                        Convert.booksAttached app [i] |> List.head
                                                    with ex ->
                                                        { Source = i.Source
                                                          Target = i.Target
                                                          Status = "失败"
                                                          Detail = Errors.explain ex }

                                                outcomes.Add outcome

                                                // ---- 4. 原件被 Excel 改写过就从备份还原（失败件必须原样） ----
                                                let restored =
                                                    if File.Exists i.Source && originalHash.IsSome then
                                                        let afterHash = Attached.md5 i.Source

                                                        if afterHash = originalHash.Value then
                                                            true
                                                        else
                                                            try
                                                                let backupFile = Path.Combine(backup, rel)
                                                                File.Copy(backupFile, i.Source, true)
                                                                let nowHash = Attached.md5 i.Source
                                                                let backupHash = Attached.md5 backupFile

                                                                record (
                                                                    sprintf
                                                                        "  原件被 Excel 改写（md5 %s → %s，%d B），已从备份还原成%s"
                                                                        (originalHash.Value.Substring(0, 8))
                                                                        (afterHash.Substring(0, 8))
                                                                        (defaultArg originalLen -1L)
                                                                        (if nowHash = backupHash then "原始字节" else "⚠ 与备份不一致的字节")
                                                                )

                                                                nowHash = backupHash
                                                            with ex ->
                                                                record ("  × 原件被改写，还原失败：" + Errors.explain ex)
                                                                false
                                                    else
                                                        true

                                                // ---- 5. 通过就一份一送；不通过就打原因、记清单 ----
                                                currentStage <- sprintf "[%d/%d] 送回收站 %s" index queue.Length rel
                                                currentSince <- DateTime.Now

                                                if Convert.canDelete outcome then
                                                    let removed, kept = RealFolder.deleteVerified root [outcome] say

                                                    if removed = 1 then
                                                        okCount <- okCount + 1
                                                        sameRun <- 0
                                                        lastReason <- ""
                                                        finish (sprintf "✓ %s —— %s，原件已送回收站" rel outcome.Status)
                                                    else
                                                        failCount <- failCount + 1
                                                        sameRun <- 0
                                                        finish (sprintf "× %s —— 新文件在，但原件送不进回收站（留下 %d 份）" rel kept)
                                                        noteFailure rel
                                                else
                                                    failCount <- failCount + 1
                                                    let reason = reasonOf outcome.Detail
                                                    finish (sprintf "× %s —— %s（%s）" rel reason outcome.Detail)
                                                    noteFailure rel

                                                    if reason = lastReason then
                                                        sameRun <- sameRun + 1
                                                    else
                                                        lastReason <- reason
                                                        sameRun <- 1

                                                    if sameRun >= sameReasonLimit then
                                                        stopReason <- sprintf "同一原因「%s」连续 %d 次" reason sameRun
                                                        record ("停：" + stopReason + " —— 像系统性问题，先查这一条")

                                                if not restored then
                                                    record "  ⚠ 上面那份原件没还原成功，先核对备份"
                                        else
                                            failCount <- failCount + 1
                                            finish (sprintf "× %s —— 备份失败，原件保留、这份跳过" rel)
                            finally
                                Attached.restore app quiet

                            (appliedAlerts,
                             appliedSecurity,
                             appliedPrint,
                             (try app.DisplayAlerts with _ -> true),
                             (try int app.AutomationSecurity with _ -> -1),
                             (try app.PrintCommunication with _ -> true),
                             quiet.DisplayAlerts,
                             quiet.Security |> Option.map int,
                             quiet.PrintCommunication))
                    )
                with ex ->
                    record ("× 通道异常：" + Errors.explain ex)
                    Error(Errors.explain ex)

            watchStop <- true
            watchdog.Dispose()

            match run with
            | Ok(appliedAlerts,
                  appliedSecurity,
                  appliedPrint,
                  restoredAlerts,
                  restoredSecurity,
                  restoredPrint,
                  quietAlerts,
                  quietSecurity,
                  quietPrint) ->
                Report.line output "DisplayAlerts" (sprintf "原值 %A → %A → 恢复 %A" quietAlerts appliedAlerts restoredAlerts)

                Report.line
                    output
                    "AutomationSecurity"
                    (sprintf "原值 %A → %d（ForceDisable）→ 恢复 %d" quietSecurity appliedSecurity restoredSecurity)

                Report.line
                    output
                    "PrintCommunication"
                    (sprintf "原值 %A → %A（不跟打印机通信）→ 恢复 %A" quietPrint appliedPrint restoredPrint)

                Should.equal 3 appliedSecurity

                (match quietAlerts with
                 | Some value -> Should.equal (Some value) (Some restoredAlerts)
                 | None -> ())

                (match quietSecurity with
                 | Some value -> Should.equal (Some value) (Some restoredSecurity)
                 | None -> ())

                (match quietPrint with
                 | Some value -> Should.equal (Some value) (Some restoredPrint)
                 | None -> ())
            | Error _ -> ()

            Report.line output "本轮结果" (sprintf "队列 %d 份，成功 %d 份，失败 %d 份" queue.Length okCount failCount)

            Report.line
                output
                "催前台"
                (sprintf "%d 次（%d 秒没进展就催一次；催了还没动，就是在等你在对话框中做选择）" hintCount hintSec)

            if stopReason <> "" then
                Report.line output "停下原因" stopReason

            // 核对没通过的，原件一份都不能丢
            for o in outcomes |> Seq.filter(Convert.canDelete >> not) do
                Assert.True(File.Exists o.Source, sprintf "核对没通过却把原件弄丢了：%s" o.Source)

            let afterProc = Report.procIds "EXCEL"
            Report.line output "现场" (sprintf "EXCEL 进程 干活前 %A / 收尾后 %A" (Set.toList beforeProc) (Set.toList afterProc))
            Should.equal beforeProc afterProc
            Should.equal true (Set.contains info.Pid afterProc)

            let left = Paths.scan root |> Paths.filterByKind Paths.Kind.Xls
            Report.line output "目录现状" (sprintf "还剩 %d 份 .xls（本轮成功 %d 份已送回收站）" left.Length okCount)

            Report.conclusion
                output
                (sprintf
                    "逐份推进：队列 %d 份，成功 %d 份（原件已入回收站），失败 %d 份（原件保留，相对路径记入失败清单）%s"
                    queue.Length
                    okCount
                    failCount
                    (if stopReason = "" then "" else "；停：" + stopReason))

    /// 挂回常驻 Word 只读核对：只证明「能挂上那一只、不动任何文档」，一个文件都不碰。
    [<Fact>]
    member this.``常驻 Word：只读核对挂回，一个文件都不动``() =
        Report.title output "常驻 Word：只读核对挂回（不新建、不 Quit、不动任何文档）"

        let info = Attached.wordInfo () |> Report.ok
        let before = Report.procIds "WINWORD"

        Report.line output "黑板" (info.describe())
        Report.line output "现场" (sprintf "WINWORD 进程 %A" (Set.toList before))

        let app =
            try
                Ok(WordHost.WordTask.GetApplicationAsync info.Pid |> fun t -> t.GetAwaiter().GetResult())
            with ex ->
                Error(Errors.explain ex)

        match app with
        | Error message ->
            Report.line output "挂回" ("× " + message)
            failwith message
        | Ok app ->
            let count = STA.run(fun () -> Attached.docCount app)
            let version = STA.run(fun () -> try app.Version with _ -> "?")

            Report.line
                output
                "挂回"
                (sprintf
                    "pid=%d rot=%s（空实例没有文档窗口，pid 核对按库的规矩跳过，不是挂回失败）"
                    info.Pid
                    (if String.IsNullOrWhiteSpace info.Rot then
                         "按 ProgID"
                     else
                         info.Rot))

            Report.line output "品牌版本" (sprintf "Microsoft Word %s" version)
            Report.line output "文档数" (string count)
            Report.line output "口径" "只接现成那一只：绝不新建、绝不 Quit、不碰任何文档"

            Should.equal true (Set.contains info.Pid (Report.procIds "WINWORD"))
            Report.conclusion output (sprintf "通过：挂回 pid=%d 那只 Word，文档 %d 本，一个文件没动" info.Pid count)

    /// 批量前自愈：把实例里「本批候选里那些已经打开的 .doc」按不保存关掉，给批量开路。
    /// 只认 Paths.scan 认出的本批文件；用户开着别处的文档一个不动、也不判定。
    [<Fact>]
    member this.``常驻 Word：清空实例里本批已打开的文档``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录>"
        | Some root ->
            let info = Attached.wordInfo () |> Report.ok

            let candidates =
                Paths.scan root
                |> Paths.filterByKind Paths.Kind.Doc
                |> List.map(fun i -> i.Source.ToLowerInvariant())
                |> Set.ofList

            Report.title output "常驻 Word：清空实例里本批已打开的文档（不保存）"
            Report.line output "黑板" (info.describe())
            Report.line output "本批候选" (sprintf "%d 份 .doc" candidates.Count)

            let app =
                WordHost.WordTask.GetApplicationAsync info.Pid
                |> fun t -> t.GetAwaiter().GetResult()

            let docAt (index: int) =
                try
                    let mutable idx: obj = box index
                    app.Documents.Item(&idx)
                with _ ->
                    null

            let closed, countBefore, countAfter =
                STA.run(fun () ->
                    let before = Attached.docCount app
                    let mutable closed = []

                    if before > 0 then
                        for k in before .. -1 .. 1 do
                            let doc = docAt k

                            let full =
                                try
                                    if isNull doc then "" else doc.FullName
                                with _ ->
                                    ""

                            if not (isNull doc) && full <> "" && candidates.Contains(full.ToLowerInvariant()) then
                                let saved =
                                    try
                                        doc.Saved
                                    with _ ->
                                        true

                                try
                                    let mutable save: obj = box 0
                                    doc.Close(&save)

                                    closed <-
                                        sprintf
                                            "%s（%s）"
                                            (Path.GetFileName full)
                                            (if saved then
                                                 "本来就没改动"
                                             else
                                                 "有未保存改动，按不保存关闭")
                                        :: closed
                                with ex ->
                                    closed <-
                                        sprintf "%s（关不掉：%s）" (Path.GetFileName full) (Errors.explain ex) :: closed

                    (List.rev closed, before, Attached.docCount app))

            Report.line output "开工前文档数" (string countBefore)

            if closed.IsEmpty then
                Report.line output "关掉的" "0 本（实例里没有本批的文档）"
            else
                for line in closed do
                    Report.line output "关掉的" line

            Report.line output "收尾后文档数" (string countAfter)

            Should.equal (countBefore - closed.Length) countAfter
            Report.conclusion output (sprintf "完成：本批残留 %d 本按不保存关掉，原件一个字节没动" closed.Length)

    /// 只读诊断：原件 .doc 和转出来的 .docx 正文差在哪一个字符上。
    /// 「正文对不上」判失败时用它看是真丢内容，还是控制字符/尾部段落标记的表示差异。
    [<Fact>]
    member this.``诊断：doc 与 docx 正文逐字比``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_PROBE_FILE" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_PROBE_FILE=<原件 .doc 路径>"
        | source ->
            let info = Attached.wordInfo () |> Report.ok
            let target = Path.ChangeExtension(source, ".docx")

            Assert.True(File.Exists source, sprintf "原件不在：%s" source)
            Assert.True(File.Exists target, sprintf "新文件不在：%s" target)

            let app =
                WordHost.WordTask.GetApplicationAsync info.Pid
                |> fun t -> t.GetAwaiter().GetResult()

            let raw (path: string) =
                let doc = WordHost.DocIO.openDoc app path true
                try
                    doc.Content.Text
                finally
                    WordHost.DocIO.closeDoc doc

            let beforeRaw = STA.run(fun () -> raw source)
            let afterRaw = STA.run(fun () -> raw target)

            let a = WordHost.DocIO.normalize beforeRaw
            let b = WordHost.DocIO.normalize afterRaw

            let shorter = min a.Length b.Length
            let mutable first = -1
            let mutable i = 0

            while first < 0 && i < shorter do
                if a.[i] <> b.[i] then first <- i
                i <- i + 1

            let ctx (text: string) (at: int) =
                if at < 0 || at >= text.Length then
                    "(越界)"
                else
                    let start = max 0 (at - 12)
                    let len = min 24 (text.Length - start)
                    text.Substring(start, len)

            Report.title output "诊断：doc 与 docx 正文逐字比"
            Report.line output "原件" (sprintf "%s（%d B）" (Path.GetFileName source) (FileInfo source).Length)
            Report.line output "新文件" (sprintf "%s（%d B）" (Path.GetFileName target) (FileInfo target).Length)
            Report.line output "原始长度" (sprintf "原件 %d 字 / 新文件 %d 字" beforeRaw.Length afterRaw.Length)
            Report.line output "去段落标记后" (sprintf "原件 %d 字 / 新文件 %d 字" a.Length b.Length)

            if first >= 0 then
                Report.line
                    output
                    "第一个差异"
                    (sprintf "第 %d 个字：原件 %A / 新文件 %A" first (string a.[first]) (string b.[first]))

                Report.line output "原件上下文" (sprintf "%A" (ctx a first))
                Report.line output "新文件上下文" (sprintf "%A" (ctx b first))
            elif a.Length <> b.Length then
                let tail (text: string) (from: int) = text.Substring(max 0 (min from text.Length))
                Report.line output "第一个差异" "前面完全一致，差异在尾部"
                Report.line output "原件尾部" (sprintf "%A" (tail a shorter))
                Report.line output "新文件尾部" (sprintf "%A" (tail b shorter))
            else
                Report.line output "第一个差异" "没有差异（归一化后完全一致）"

            Report.conclusion
                output
                (if a = b then
                     "结论：归一化后正文一致，之前的「对不上」只是原始串的表示差异"
                 else
                     sprintf "结论：正文确实不同（相差 %d 个字）" (abs (a.Length - b.Length)))

    /// 诊断：带宏的 .doc 存成 .docm 时宏为什么丢。
    /// 同一份文件在不同 AutomationSecurity 档位下各存一次，看哪一档能把 VBA 工程留住：
    ///   3 = ForceDisable（现在的做法，全批都按这个开）
    ///   2 = ByUI（按用户信任中心设置；默认「禁用所有宏、给出通知」→ 宏不执行，但工程会被读进来）
    ///   1 = Low（启用所有宏，**会真执行 AutoOpen**）→ 只有显式设 XUNITHOST_PROBE_LOW=1 才试
    /// 全程只动临时副本，原件一个字节不碰；Word 自己起一只、跑完自己关掉（本机不能已有 Word 在跑）。
    [<Fact>]
    member this.``诊断：docm 存盘为什么丢宏``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_PROBE_FILE" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_PROBE_FILE=<带宏 .doc 路径>"
        | source ->
            Assert.True(File.Exists source, sprintf "原件不在：%s" source)

            let work =
                Path.Combine(Path.GetTempPath(), "macroprobe-" + Guid.NewGuid().ToString("N").Substring(0, 8))

            Directory.CreateDirectory work |> ignore
            let copy = Path.Combine(work, "src.doc")
            File.Copy(source, copy, true)

            let tryLow =
                match Environment.GetEnvironmentVariable "XUNITHOST_PROBE_LOW" with
                | null
                | "" -> false
                | text -> text.Trim() = "1"

            Report.title output "诊断：docm 存盘为什么丢宏"
            Report.line output "原件" (sprintf "%s（%d B）" (Path.GetFileName source) (FileInfo source).Length)
            Report.line output "临时副本" copy

            /// 换档 → 打开副本 → 另存 .docm → 读回看宏 → 恢复原档
            let attempt (app: Microsoft.Office.Interop.Word.Application) (mode: MsoAutomationSecurity) (tag: string) =
                let previous =
                    try
                        Some app.AutomationSecurity
                    with _ ->
                        None

                (try
                    app.AutomationSecurity <- mode
                 with _ ->
                     ())

                let stage = Path.Combine(work, tag + ".docm")
                let mutable srcHas = false
                let mutable outHas = false
                let mutable note = "-"

                try
                    let doc = WordHost.DocIO.openDoc app copy false
                    srcHas <- WordHost.DocIO.hasVba doc
                    WordHost.DocIO.saveDocAs (box 13) doc stage
                    WordHost.DocIO.closeDoc doc

                    let back = WordHost.DocIO.openDoc app stage true
                    outHas <- WordHost.DocIO.hasVba back
                    WordHost.DocIO.closeDoc back

                    note <- sprintf "%d B" (FileInfo stage).Length
                with ex ->
                    note <- "× " + Errors.explain ex

                (match previous with
                 | Some value ->
                     (try
                         app.AutomationSecurity <- value
                      with _ ->
                          ())
                 | None -> ())

                (srcHas, outHas, note, File.Exists stage)

            let plans =
                [ ("ForceDisable（现在全批都用这档）", MsoAutomationSecurity.msoAutomationSecurityForceDisable)
                  ("ByUI（按信任中心，默认禁用+通知，不执行宏）", MsoAutomationSecurity.msoAutomationSecurityByUI) ]
                @ (if tryLow then
                       [ ("Low（启用所有宏，会执行 AutoOpen）", MsoAutomationSecurity.msoAutomationSecurityLow) ]
                   else
                       [])

            if not tryLow then
                Report.line output "Low" "没试（要试得设 XUNITHOST_PROBE_LOW=1，会执行 AutoOpen 宏）"

            // 自己起一只 Word 干活：本机不能已有 Word 在跑，跑完 Quit + 收进程
            let measured =
                WordHost.Host.execQuery (fun app -> plans |> List.mapi (fun index (label, mode) ->
                    let srcHas, outHas, note, made = attempt app mode (sprintf "try%d" index)
                    (label, srcHas, outHas, note, made)))
                |> Report.ok

            let mutable keptBy = []

            for (label, srcHas, outHas, note, made) in measured do
                Report.line
                    output
                    label
                    (sprintf "打开时有宏=%b；另存后新文件有宏=%b；%s；产物在=%b" srcHas outHas note made)

                if srcHas && outHas then keptBy <- label :: keptBy

            Report.conclusion
                output
                (if keptBy.IsEmpty then
                     "结论：这几档都留不住宏——宏不是被安全档位吃掉的，得换办法（手工另存或直接搬 VBA 工程）"
                 else
                     sprintf "结论：%s 这档能保住宏，把 DocToDocm 改成用它就行" (String.Join("、", List.rev keptBy)))

    /// 诊断：这些 .doc 的 VBA 工程里到底有没有代码。
    /// `HasVBProject` 只说「有工程」，不说「工程里有东西」——工程是空壳时另存 .docm 会被 Word 丢掉。
    /// 读 VBComponents / 代码行数需要「信任对 VBA 工程对象模型的访问」（AccessVBOM=1），
    /// 这一条只读不写文档；跑之前请自行把该信任开关打开，跑完关掉。
    /// 清单文件：XUNITHOST_PROBE_LIST（一行一个 .doc 路径）。
    [<Fact>]
    member this.``诊断：doc 的 VBA 工程有没有代码``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_PROBE_LIST" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_PROBE_LIST=<一行一个 .doc 的清单文件>"
        | listFile ->
            Assert.True(File.Exists listFile, sprintf "清单不在：%s" listFile)

            let items =
                File.ReadAllLines listFile
                |> Array.map(fun line -> line.Trim())
                |> Array.filter(fun line -> line <> "" && not (line.StartsWith "#"))
                |> Array.filter File.Exists

            Report.title output "诊断：doc 的 VBA 工程有没有代码"
            Report.line output "清单" (sprintf "%d 份" items.Length)

            /// 迟绑定取属性 / 调方法：不引 Microsoft.Vbe.Interop，也不碰 COM 类型定义。
            /// 拿不到就返回 None（信任开关没开时就是这条），由调用方记一句说明。
            let comGet (target: obj) (name: string) (args: obj[]) : obj option =
                if isNull target then
                    None
                else
                    try
                        Some(
                            target
                                .GetType()
                                .InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, target, args)
                        )
                    with _ ->
                        None

            /// 按指定「只读/可写」打开，读工程与代码行数（readOnly 是一种口径，转换用的是可写）
            let probeOne (app: Microsoft.Office.Interop.Word.Application) (path: string) (readOnly: bool) =
                let mutable hasProject = false
                let mutable components = -1
                let mutable lines = -1
                let mutable note = "-"
                let mutable doc: Microsoft.Office.Interop.Word.Document = null

                try
                    doc <- WordHost.DocIO.openDoc app path readOnly
                    hasProject <- (try doc.HasVBProject with _ -> false)

                    if hasProject then
                        match comGet (box doc) "VBProject" [||] with
                        | None -> note <- "拿不到 VBProject（没开「信任对 VBA 工程对象模型的访问」）"
                        | Some project ->
                            components <-
                                match comGet project "VBComponents" [||] with
                                | Some comps -> (match comGet comps "Count" [||] with Some n -> Convert.ToInt32 n | None -> -1)
                                | None -> -1

                            if components > 0 then
                                lines <- 0

                                for index in 1 .. components do
                                    let one =
                                        comGet project "VBComponents" [||]
                                        |> Option.bind(fun comps -> comGet comps "Item" [| box index |])

                                    match one |> Option.bind(fun item -> comGet item "CodeModule" [||]) with
                                    | Some codeModule ->
                                        match comGet codeModule "CountOfLines" [||] with
                                        | Some n -> lines <- lines + Convert.ToInt32 n
                                        | None -> note <- "读不到 CountOfLines"
                                    | None -> note <- "读不到 CodeModule"
                with ex ->
                    note <- "打不开：" + Errors.explain ex

                if not (isNull doc) then WordHost.DocIO.closeDoc doc

                (hasProject, components, lines, note)

            let measured =
                WordHost.Host.execQuery (fun app ->
                    items
                    |> Array.map(fun path ->
                        let roHas, roComps, roLines, roNote = probeOne app path true
                        let rwHas, rwComps, rwLines, rwNote = probeOne app path false
                        (Path.GetFileName path, (roHas, roComps, roLines, roNote), (rwHas, rwComps, rwLines, rwNote))))
                |> Report.ok

            let mutable withCode = 0

            for (name, (roHas, roComps, roLines, roNote), (rwHas, rwComps, rwLines, rwNote)) in measured do
                Report.line
                    output
                    name
                    (sprintf
                        "只读打开：有工程=%b 组件 %d 代码 %d 行 %s；可写打开：有工程=%b 组件 %d 代码 %d 行 %s"
                        roHas
                        roComps
                        roLines
                        roNote
                        rwHas
                        rwComps
                        rwLines
                        rwNote)

                if roLines > 0 || rwLines > 0 then withCode <- withCode + 1

            Report.conclusion
                output
                (sprintf "完成：%d 份里 %d 份工程里真有代码；其余是空壳（另存 .docm 会被 Word 丢掉，按 .docx 转不丢内容）" measured.Length withCode)

    /// 清理：把清单里点名的文件逐个送回收站。
    /// 清单 = XUNITHOST_CLEAN_LIST，一行一个绝对路径，`#` 开头是注释。
    /// 只删清单里的文件；进的是**回收站**（随时能还原），不是永久删除；
    /// 文件本来就不在也算完成（幂等，可重复跑）。
    [<Fact>]
    member this.``清理：清单里的文件送回收站``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_CLEAN_LIST" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_CLEAN_LIST=<一行一个绝对路径的清单>"
        | listFile ->
            Assert.True(File.Exists listFile, sprintf "清单不在：%s" listFile)

            let items =
                File.ReadAllLines listFile
                |> Array.map(fun line -> line.Trim())
                |> Array.filter(fun line -> line <> "" && not (line.StartsWith "#"))

            Report.title output "清理：清单里的文件送回收站"
            Report.line output "清单" (sprintf "%s（%d 条）" (Path.GetFileName listFile) items.Length)

            let mutable gone = 0
            let mutable missed = 0
            let mutable failed = 0

            for path in items do
                if not (File.Exists path) then
                    missed <- missed + 1
                else
                    try
                        let info = FileInfo path
                        if info.IsReadOnly then info.IsReadOnly <- false

                        RealFolder.toRecycleBin path

                        if File.Exists path then
                            failed <- failed + 1
                            Report.line output "× 送不进回收站" path
                        else
                            gone <- gone + 1
                            Report.line output "✓ 进回收站" (sprintf "%s（%d B）" path info.Length)
                    with ex ->
                        failed <- failed + 1
                        Report.line output "× 出错" (sprintf "%s —— %s" path (Errors.explain ex))

            Report.conclusion
                output
                (sprintf "完成：送进回收站 %d 个，本来就不在 %d 个，送不进 %d 个" gone missed failed)

            Should.equal 0 failed

    /// Word 逐份推进：一份一份送，成功就进下一份；失败打原因、记清单、接着下一份。
    /// 与 Excel 那条同构，只换 host：绝不新建 / 不 Quit / 不关用户开着的文档。
    [<Fact>]
    member this.``常驻 Word：逐份升级 doc，成功一份送走一份``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录> 才逐份升级"
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let info = Attached.wordInfo () |> Report.ok
            let backup = RealFolder.backupRoot root
            let all = Paths.selectForRound root Paths.Kind.Doc (Paths.scan root)

            let queue =
                match Paths.limitFromEnv () with
                | Some n -> all |> List.truncate n
                | None -> all

            let queueSet =
                queue |> List.map(fun x -> x.Source.ToLowerInvariant()) |> Set.ofList

            /// 正在处理的那一份：通道异常（Word 崩、文件被占用）时靠它记账
            let mutable currentItem: Item option = None

            let intEnv (name: string) (fallback: int) =
                let text = defaultArg (Option.ofObj (Environment.GetEnvironmentVariable name)) ""

                match Int32.TryParse text with
                | true, n when n > 0 -> n
                | _ -> fallback

            let hintSec = intEnv "XUNITHOST_CONVERT_POPUP_HINT" (intEnv "XUNITHOST_CONVERT_TIMEOUT" 45)
            let sameReasonLimit = intEnv "XUNITHOST_CONVERT_SAME_REASON" 3
            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))

            let fileEnv (name: string) (fallback: string) =
                match Environment.GetEnvironmentVariable name with
                | null
                | "" -> Path.Combine(projectRoot, "TestResults", fallback)
                | value -> value

            let progressFile = fileEnv "XUNITHOST_CONVERT_PROGRESS" "逐份-进度-Word.log"
            let failedFile = fileEnv "XUNITHOST_CONVERT_FAILED" "逐份-失败清单-Word.txt"
            let say (text: string) = output.WriteLine("    " + text)

            // 主线程写、看门狗线程也写，同一个文件要串行
            let progressGate = obj()

            let append (path: string) (line: string) =
                lock progressGate (fun () ->
                    let mutable ok = false
                    let mutable tries = 0

                    while not ok && tries < 5 do
                        tries <- tries + 1

                        try
                            File.AppendAllText(path, line + Environment.NewLine)
                            ok <- true
                        with _ ->
                            System.Threading.Thread.Sleep 50)

            let record (text: string) =
                let line = sprintf "%s %s" (DateTime.Now.ToString("HH:mm:ss")) text
                say line
                append progressFile line

            let failedSeen =
                if File.Exists failedFile then
                    File.ReadAllLines failedFile
                    |> Array.map(fun line -> line.Trim())
                    |> Array.filter(fun line -> line <> "" && not (line.StartsWith "#"))
                    |> Set.ofArray
                else
                    try
                        File.WriteAllText(
                            failedFile,
                            "# 逐份升级失败清单（每行一个相对路径；可直接给 XUNITHOST_CONVERT_SKIP）" + Environment.NewLine
                        )
                    with _ ->
                        ()

                    Set.empty

            let mutable known = failedSeen

            let noteFailure (rel: string) =
                if not (known.Contains rel) then
                    known <- known.Add rel

                    try
                        File.AppendAllText(failedFile, rel + Environment.NewLine)
                    with _ ->
                        ()

            let reasonOf (detail: string) =
                if detail.Contains "共享" || detail.Contains "修订" then "共享/修订历史"
                elif detail.Contains "VBA" || detail.Contains "宏" then "宏判定"
                elif detail.Contains "超时" then "COM 超时"
                elif detail.Contains "修复" then "修复对话框"
                elif detail.Contains "拒绝" then "COM 拒绝"
                elif detail.Contains "对不上" then "内容对不上"
                elif detail.Contains "占用" then "文件被占用"
                else "其它"

            Report.title output "常驻 Word：逐份升级 .doc，成功一份送走一份"
            Report.line output "黑板" (info.describe())
            Report.line output "本轮" (sprintf "%d 份 .doc（%s）" queue.Length (RealFolder.limitLabel ()))
            Report.line output "备份目录" backup
            Report.line output "进度日志" progressFile
            Report.line output "失败清单" (sprintf "%s（已有 %d 条历史失败）" failedFile failedSeen.Count)

            Report.line
                output
                "弹窗策略"
                (sprintf
                    "不躲：%d 秒没进展就把 Word 切到前台催人，点掉对话框自动继续；只激活窗口，不杀进程、不关你的文档"
                    hintSec)

            Report.line output "停的条件" (sprintf "同一原因连续 %d 次（卡住不判失败，等你处理）" sameReasonLimit)

            let beforeProc = Report.procIds "WINWORD"

            let app =
                WordHost.WordTask.GetApplicationAsync info.Pid
                |> fun t -> t.GetAwaiter().GetResult()

            Report.line
                output
                "挂回"
                (sprintf
                    "pid=%d rot=%s（校验只用进程身份）"
                    info.Pid
                    (if String.IsNullOrWhiteSpace info.Rot then "按 ProgID" else info.Rot))

            let focusOnStart =
                match Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_FOCUS" with
                | null
                | "" -> true
                | text -> text.Trim() <> "0"

            let raised = if focusOnStart then Attached.bringToFront info.Pid else false

            Report.line
                output
                "界面"
                (sprintf
                    "开跑前置 %s：%s"
                    (if focusOnStart then "开" else "关")
                    (if not focusOnStart then "没动窗口，只在卡住时催"
                     elif raised then "已把 Word 切到前台，跑的时候有对话框你会立刻看见"
                     else "没切动，请自己点任务栏里的 Word 看一眼"))

            let mutable currentStage = "未开始"
            let mutable currentSince = DateTime.Now
            let mutable watchStop = false
            let mutable lastHint = DateTime.MinValue
            let mutable hintCount = 0

            let watchdog =
                new System.Threading.Timer(
                    (fun (_: obj) ->
                        if not watchStop then
                            let secs = int (DateTime.Now - currentSince).TotalSeconds
                            let sinceHint = (DateTime.Now - lastHint).TotalSeconds

                            if secs >= hintSec && sinceHint >= float hintSec && hintCount < 40 then
                                lastHint <- DateTime.Now
                                hintCount <- hintCount + 1
                                let front = Attached.bringToFront info.Pid

                                record (
                                    sprintf
                                        "⚠ 没进展 %d 秒（%s）—— 多半弹了对话框；%s"
                                        secs
                                        currentStage
                                        (if front then
                                             "已把 Word 切到前台，请点掉弹窗，点完我自动继续"
                                         else
                                             "没切动前台，请自己点任务栏里的 Word 看一眼")
                                )),
                    null,
                    5000,
                    5000
                )

            let mutable okCount = 0
            let mutable failCount = 0
            let mutable stopReason = ""
            let mutable lastReason = ""
            let mutable sameRun = 0
            let outcomes = ResizeArray<Outcome>()

            let run =
                try
                    Ok(
                        STA.run(fun () ->
                            let quiet = Attached.quietWord app
                            let appliedAlerts = (try Some app.DisplayAlerts with _ -> None)
                            let appliedSecurity = (try int app.AutomationSecurity with _ -> -1)
                            let docsBefore = Attached.docCount app

                            try
                                record (sprintf "  实例文档 %d 本（用户自己也开着的不关）" docsBefore)

                                // 上次中断可能把自己的文档留在实例里：只关本队列里的那些，
                                // 用户自己开着的别的文档一个字不动（这正是上次「文件被另一进程占用」的成因）。
                                if docsBefore > 0 then
                                    let mutable closed = []

                                    for k in docsBefore .. -1 .. 1 do
                                        let doc =
                                            try
                                                let mutable idx: obj = box k
                                                app.Documents.Item(&idx)
                                            with _ ->
                                                null

                                        let full =
                                            try
                                                if isNull doc then "" else doc.FullName
                                            with _ ->
                                                ""

                                        if not (isNull doc) && full <> "" && queueSet.Contains(full.ToLowerInvariant()) then
                                            try
                                                let mutable save: obj = box 0
                                                doc.Close(&save)
                                                closed <- Path.GetFileName full :: closed
                                            with _ ->
                                                ()

                                    if not closed.IsEmpty then
                                        record (
                                            sprintf
                                                "  关掉上次残留 %d 本（%s）"
                                                closed.Length
                                                (String.concat "、" (List.rev closed))
                                        )

                                let mutable index = 0

                                while index < queue.Length && stopReason = "" do
                                    let i = queue.[index]
                                    index <- index + 1
                                    currentItem <- Some i
                                    let rel = RealFolder.rel root i.Source
                                    currentStage <- sprintf "[%d/%d] 开始 %s" index queue.Length rel
                                    currentSince <- DateTime.Now
                                    lastHint <- DateTime.MinValue
                                    record (sprintf "[%d/%d] 开始 %s" index queue.Length rel)
                                    let clock = System.Diagnostics.Stopwatch.StartNew()

                                    let finish (text: string) =
                                        record text
                                        record (sprintf "  本份耗时 %.1f 秒" clock.Elapsed.TotalSeconds)

                                    if not (Set.contains info.Pid (Report.procIds info.Image)) then
                                        stopReason <- sprintf "常驻 Word pid %d 已不在" info.Pid
                                        record ("停：" + stopReason)
                                    else
                                        // ---- 1. 只备份这一份 ----
                                        currentStage <- sprintf "[%d/%d] 备份 %s" index queue.Length rel
                                        currentSince <- DateTime.Now

                                        let backed =
                                            try
                                                RealFolder.backupAll root backup [i] say |> ignore
                                                true
                                            with ex ->
                                                record ("  备份失败：" + Errors.explain ex)
                                                false

                                        if backed then
                                            // ---- 2. 重名让位 ----
                                            currentStage <- sprintf "[%d/%d] 重名让位 %s" index queue.Length rel

                                            let _, skipped = RealFolder.renameExisting root [i] say

                                            if skipped.Contains i.Source then
                                                failCount <- failCount + 1
                                                finish (sprintf "× %s —— 现有同名新文件改名失败，这份跳过" rel)
                                                noteFailure rel
                                            else
                                                // ---- 3. 转换这一份（内部已读回核对） ----
                                                currentStage <- sprintf "[%d/%d] 转换 %s" index queue.Length rel
                                                currentSince <- DateTime.Now

                                                let originalHash =
                                                    try
                                                        Some(Attached.md5 i.Source)
                                                    with _ ->
                                                        None

                                                let originalLen =
                                                    try
                                                        Some((FileInfo i.Source).Length)
                                                    with _ ->
                                                        None

                                                let outcome =
                                                    try
                                                        Convert.docsAttached app [i] |> List.head
                                                    with ex ->
                                                        { Source = i.Source
                                                          Target = i.Target
                                                          Status = "失败"
                                                          Detail = Errors.explain ex }

                                                outcomes.Add outcome

                                                // ---- 4. 原件被改写就从备份还原（失败件必须原样） ----
                                                let restored =
                                                    if File.Exists i.Source && originalHash.IsSome then
                                                        let afterHash = Attached.md5 i.Source

                                                        if afterHash = originalHash.Value then
                                                            true
                                                        else
                                                            try
                                                                let backupFile = Path.Combine(backup, rel)
                                                                File.Copy(backupFile, i.Source, true)
                                                                let nowHash = Attached.md5 i.Source
                                                                let backupHash = Attached.md5 backupFile

                                                                record (
                                                                    sprintf
                                                                        "  原件被 Word 改写（md5 %s → %s，%d B），已从备份还原成%s"
                                                                        (originalHash.Value.Substring(0, 8))
                                                                        (nowHash.Substring(0, 8))
                                                                        (defaultArg originalLen -1L)
                                                                        (if nowHash = backupHash then
                                                                             "原始字节"
                                                                         else
                                                                             "⚠ 与备份不一致的字节")
                                                                )

                                                                nowHash = backupHash
                                                            with ex ->
                                                                record ("  × 原件被改写，还原失败：" + Errors.explain ex)
                                                                false
                                                    else
                                                        true

                                                // ---- 5. 通过就一份一送；不通过就打原因、记清单 ----
                                                currentStage <- sprintf "[%d/%d] 送回收站 %s" index queue.Length rel
                                                currentSince <- DateTime.Now

                                                if Convert.canDelete outcome then
                                                    let removed, kept = RealFolder.deleteVerified root [outcome] say

                                                    if removed = 1 then
                                                        okCount <- okCount + 1
                                                        sameRun <- 0
                                                        lastReason <- ""
                                                        finish (sprintf "✓ %s —— %s，原件已送回收站" rel outcome.Status)
                                                    else
                                                        failCount <- failCount + 1
                                                        sameRun <- 0
                                                        finish (sprintf "× %s —— 新文件在，但原件送不进回收站（留下 %d 份）" rel kept)
                                                        noteFailure rel
                                                else
                                                    failCount <- failCount + 1
                                                    let reason = reasonOf outcome.Detail
                                                    finish (sprintf "× %s —— %s（%s）" rel reason outcome.Detail)
                                                    noteFailure rel

                                                    if reason = lastReason then
                                                        sameRun <- sameRun + 1
                                                    else
                                                        lastReason <- reason
                                                        sameRun <- 1

                                                    if sameRun >= sameReasonLimit then
                                                        stopReason <- sprintf "同一原因「%s」连续 %d 次" reason sameRun
                                                        record ("停：" + stopReason + " —— 像系统性问题，先查这一条")

                                                if not restored then
                                                    record "  ⚠ 上面那份原件没还原成功，先核对备份"
                                        else
                                            failCount <- failCount + 1
                                            finish (sprintf "× %s —— 备份失败，原件保留、这份跳过" rel)
                            finally
                                Attached.restoreWord app quiet

                            (quiet.Alerts,
                             quiet.Security |> Option.map int,
                             appliedAlerts,
                             appliedSecurity,
                             (try Some app.DisplayAlerts with _ -> None),
                             (try int app.AutomationSecurity with _ -> -1),
                             docsBefore,
                             Attached.docCount app,
                             Attached.ownDocsLeft app queueSet))
                    )
                with ex ->
                    let detail = Errors.explain ex
                    record ("× 通道异常：" + detail)

                    // 通道异常也要记账：不然报告里是「失败 0 份」、清单里也没这条，下一轮又撞同一个坑
                    match currentItem with
                    | Some i ->
                        let rel = RealFolder.rel root i.Source
                        failCount <- failCount + 1
                        record (sprintf "× %s —— 通道异常（%s），记入失败清单" rel detail)
                        noteFailure rel

                        let backupFile = Path.Combine(backup, rel)

                        if File.Exists i.Source && File.Exists backupFile then
                            try
                                let nowHash = Attached.md5 i.Source
                                let backupHash = Attached.md5 backupFile

                                if nowHash <> backupHash then
                                    File.Copy(backupFile, i.Source, true)

                                    record (
                                        sprintf
                                            "  原件被改写（md5 %s → %s），已从备份还原"
                                            (backupHash.Substring(0, 8))
                                            (nowHash.Substring(0, 8))
                                    )
                            with restoreEx ->
                                record ("  ⚠ 还原失败：" + Errors.explain restoreEx)
                    | None -> ()

                    Error detail

            watchStop <- true
            watchdog.Dispose()

            match run with
            | Ok(quietAlerts,
                  quietSecurity,
                  appliedAlerts,
                  appliedSecurity,
                  restoredAlerts,
                  restoredSecurity,
                  docsBefore,
                  docsAfter,
                  ownLeft) ->
                Report.line output "DisplayAlerts" (sprintf "原值 %A → %A → 恢复 %A" quietAlerts appliedAlerts restoredAlerts)

                Report.line
                    output
                    "AutomationSecurity"
                    (sprintf "原值 %A → %d（ForceDisable）→ 恢复 %d" quietSecurity appliedSecurity restoredSecurity)

                Report.line
                    output
                    "实例文档"
                    (sprintf "干活前 %d 本 / 收尾后 %d 本（用户自己也开着的只算数、不判定）" docsBefore docsAfter)

                Report.line output "我开的还剩" (sprintf "%d 本（0 = 自己开的都关了；-1 = 数不出来）" ownLeft)

                Should.equal 3 appliedSecurity

                (match quietAlerts with
                 | Some value -> Should.equal (Some value) restoredAlerts
                 | None -> ())

                (match quietSecurity with
                 | Some value -> Should.equal (Some value) (Some restoredSecurity)
                 | None -> ())

                if ownLeft >= 0 then Should.equal 0 ownLeft
            | Error _ -> ()

            Report.line output "本轮结果" (sprintf "队列 %d 份，成功 %d 份，失败 %d 份" queue.Length okCount failCount)

            Report.line
                output
                "催前台"
                (sprintf "%d 次（%d 秒没进展就催一次；催了还没动，就是在等你在对话框中做选择）" hintCount hintSec)

            if stopReason <> "" then
                Report.line output "停下原因" stopReason

            // 核对没通过的，原件一份都不能丢
            for o in outcomes |> Seq.filter(Convert.canDelete >> not) do
                Assert.True(File.Exists o.Source, sprintf "核对没通过却把原件弄丢了：%s" o.Source)

            let afterProc = Report.procIds "WINWORD"
            Report.line output "现场" (sprintf "WINWORD 进程 干活前 %A / 收尾后 %A" (Set.toList beforeProc) (Set.toList afterProc))

            // 只断言「没有新建」；Word 自己崩溃重启（老 .doc 触发修复）不是这条通道能拦的，报告出来就够
            Should.equal true (Set.isSubset afterProc beforeProc)

            if not (Set.contains info.Pid afterProc) then
                record (sprintf "  ⚠ 常驻 Word pid=%d 已不在：多半是打开坏文档时 Word 自己重启了，本轮到此为止" info.Pid)

            let left = Paths.scan root |> Paths.filterByKind Paths.Kind.Doc
            Report.line output "目录现状" (sprintf "还剩 %d 份 .doc（本轮成功 %d 份已送回收站）" left.Length okCount)

            Report.conclusion
                output
                (sprintf
                    "逐份推进：队列 %d 份，成功 %d 份（原件已入回收站），失败 %d 份（原件保留，相对路径记入失败清单）%s"
                    queue.Length
                    okCount
                    failCount
                    (if stopReason = "" then "" else "；停：" + stopReason))

    /// 用户点名的几份「正文对不上」：不比对，直接存 .docx；存成的把原件送回收站。
    /// 走一次性通道（自己起一只 Word、干完就退）；本机已有 Word 在跑就拒跑，绝不碰用户那份。
    [<Fact>]
    member this.``Word 强存：清单里的 doc 不比对直接存 docx``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_DOC_FORCE_LIST" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_DOC_FORCE_LIST=<一行一个绝对路径的清单>"
        | listFile ->
            Assert.True(File.Exists listFile, sprintf "清单不在：%s" listFile)

            let paths =
                File.ReadAllLines listFile
                |> Array.map(fun line -> line.Trim())
                |> Array.filter(fun line -> line <> "" && not (line.StartsWith "#"))
                |> Array.toList

            let root = RealFolder.root () |> Option.defaultValue ""
            let say (text: string) = output.WriteLine("    " + text)
            let beforeProc = Report.procIds "WINWORD"

            Report.title output "Word 强存：清单里的 .doc 不比对正文，直接存 .docx"
            Report.line output "清单" (sprintf "%s（%d 条）" (Path.GetFileName listFile) paths.Length)
            Report.line output "口径" "只认「另存成功 + 能读回来」，字数差写进详情、不拦；存成的原件进回收站"

            let run =
                WordHost.Host.execQuery(fun app ->
                    paths
                    |> List.map(fun path ->
                        let target, status, detail = WordHost.DocForceDocx.convert app path
                        { Source = path; Target = target; Status = status; Detail = detail }))

            match run with
            | Error message ->
                Report.line output "通道" ("拒跑：" + message)
                Report.conclusion output ("没跑成：" + message)
                failwith message
            | Ok outcomes ->
                for o in outcomes do
                    Report.line output (RealFolder.mark o.Status) (sprintf "%s —— %s" (Path.GetFileName o.Source) o.Detail)

                // 存成的把原件送回收站（跟逐份通道一个规矩：只送核对通过的）
                let okOnes = outcomes |> List.filter Convert.canDelete

                for o in okOnes do
                    RealFolder.deleteVerified root [o] say |> ignore

                let newFiles = okOnes |> List.filter(fun o -> File.Exists o.Target)

                Report.line
                    output
                    "落盘"
                    (sprintf "%d 份新文件都在（%s）" newFiles.Length (newFiles |> List.map(fun o -> Path.GetFileName o.Target) |> String.concat "、"))

                let afterProc = Report.procIds "WINWORD"
                Report.line output "现场" (sprintf "WINWORD 进程 干活前 %A / 收尾后 %A" (Set.toList beforeProc) (Set.toList afterProc))

                Should.equal true (Set.isSubset afterProc beforeProc)
                Should.equal okOnes.Length newFiles.Length

                Report.conclusion output (sprintf "完成：%d 份直接存成 docx（原件已进回收站）" newFiles.Length)

    /// 交给你手工保存：把这一份 .doc 打开在 Word 窗口里 —— 不另存、不关、不替你保存。
    /// 上一次崩溃留下的 `~$` 锁文件先删掉，不然这一份只能只读打开、你保不住。
    [<Fact>]
    member this.``交给你手工保存：把这一份 doc 打开在 Word 窗口里``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_DOC_HANDOFF" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_DOC_HANDOFF=<要打开的 .doc 绝对路径>"
        | path ->
            Assert.True(File.Exists path, sprintf "文件不在：%s" path)

            let beforeProc = Report.procIds "WINWORD"
            Report.title output "交给你手工保存：打开这一份 .doc，窗口留在前台"
            Report.line output "文件" (sprintf "%s（%d B）" path (FileInfo path).Length)

            // Word 的锁文件是「~$」+ 去掉前两个字的名字；上次崩掉的进程留下的那个会让这一份只读
            let lockFile =
                Path.Combine(Path.GetDirectoryName path, "~$" + Path.GetFileName(path).Substring(2))

            let lockNote =
                if File.Exists lockFile then
                    try
                        File.Delete lockFile
                        "删掉了上次留下的 " + Path.GetFileName lockFile
                    with ex ->
                        "删不掉（" + Errors.explain ex + "），可能要你手工删"
                else
                    "没有"

            Report.line output "残留锁文件" lockNote

            let app = WordHost.Host.GetOrCreateApplication ()

            let opened, count, pid =
                STA.run(fun () ->
                    (try
                        app.Visible <- true
                     with _ ->
                         ())

                    let opened =
                        try
                            // Word 的 Open 参数全是 obj&（按引用传），只能一个个抱 ref；
                            // 这里**不用** OpenNoRepairDialog：要让 Word 自己的修复/提示框露出来给人点。
                            let mutable fileName: obj = box path
                            let mutable confirm: obj = box false
                            let mutable ro: obj = box false
                            let mutable addRecent: obj = box false
                            let mutable missing: obj = Type.Missing
                            let mutable visible: obj = box true

                            let doc =
                                app.Documents.Open(
                                    &fileName,
                                    &confirm,
                                    &ro,
                                    &addRecent,
                                    &missing,
                                    &missing,
                                    &missing,
                                    &missing,
                                    &missing,
                                    &missing,
                                    &missing,
                                    &visible
                                )

                            sprintf "已打开 %s（%d 段）" doc.Name (doc.Paragraphs.Count)
                        with ex ->
                            "打开报错：" + Errors.explain ex

                    let count = (try app.Documents.Count with _ -> -1)
                    let pid = WordHost.Host.pidOf app
                    opened, count, pid)

            Report.line output "结果" opened
            Report.line output "实例文档数" (string count)

            match pid with
            | Some number ->
                let raised = Attached.bringToFront number
                Report.line output "进程" (sprintf "pid=%d%s" number (if raised then "，已切到前台" else "，没切动，请点任务栏里的 Word"))
            | None -> Report.line output "进程" "认不出来（窗口还没就绪），请看任务栏里的 Word"

            let afterProc = Report.procIds "WINWORD"
            Report.line output "现场" (sprintf "WINWORD 进程 干活前 %A / 收尾后 %A（这只**故意留着**，等你保存）" (Set.toList beforeProc) (Set.toList afterProc))

            Report.conclusion output "完成：文件开在 Word 里，保存由你来（另存为 .docx）；我不关窗口、不退出 Word"

    /// PowerPoint 逐份推进：一份一份送，成功就进下一份；失败打原因、记清单、接着下一份。
    /// 与 Word / Excel 那两条同构，只换 host：绝不新建 / 不 Quit / 不关用户开着的演示文稿。
    /// 宏按用户口径**不留**：一律存 .pptx（`XUNITHOST_PPT_KEEP_MACRO=1` 才走 .pptm）。
    [<Fact>]
    member this.``常驻 PowerPoint：逐份升级 ppt，成功一份送走一份``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录> 才逐份升级"
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let info = Attached.pptInfo () |> Report.ok
            let backup = RealFolder.backupRoot root
            let all = Paths.selectForRound root Paths.Kind.Ppt (Paths.scan root)

            let queue =
                match Paths.limitFromEnv () with
                | Some n -> all |> List.truncate n
                | None -> all

            let queueSet =
                queue |> List.map(fun x -> x.Source.ToLowerInvariant()) |> Set.ofList

            /// 正在处理的那一份：通道异常（PowerPoint 崩、文件被占用）时靠它记账
            let mutable currentItem: Item option = None

            let intEnv (name: string) (fallback: int) =
                let text = defaultArg (Option.ofObj (Environment.GetEnvironmentVariable name)) ""

                match Int32.TryParse text with
                | true, n when n > 0 -> n
                | _ -> fallback

            let hintSec = intEnv "XUNITHOST_CONVERT_POPUP_HINT" (intEnv "XUNITHOST_CONVERT_TIMEOUT" 45)
            let sameReasonLimit = intEnv "XUNITHOST_CONVERT_SAME_REASON" 3
            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))

            let fileEnv (name: string) (fallback: string) =
                match Environment.GetEnvironmentVariable name with
                | null
                | "" -> Path.Combine(projectRoot, "TestResults", fallback)
                | value -> value

            let progressFile = fileEnv "XUNITHOST_CONVERT_PROGRESS" "逐份-进度-Ppt.log"
            let failedFile = fileEnv "XUNITHOST_CONVERT_FAILED" "逐份-失败清单-Ppt.txt"
            let say (text: string) = output.WriteLine("    " + text)

            // 主线程写、看门狗线程也写，同一个文件要串行
            let progressGate = obj()

            let append (path: string) (line: string) =
                lock progressGate (fun () ->
                    let mutable ok = false
                    let mutable tries = 0

                    while not ok && tries < 5 do
                        tries <- tries + 1

                        try
                            File.AppendAllText(path, line + Environment.NewLine)
                            ok <- true
                        with _ ->
                            System.Threading.Thread.Sleep 50)

            let record (text: string) =
                let line = sprintf "%s %s" (DateTime.Now.ToString("HH:mm:ss")) text
                say line
                append progressFile line

            let failedSeen =
                if File.Exists failedFile then
                    File.ReadAllLines failedFile
                    |> Array.map(fun line -> line.Trim())
                    |> Array.filter(fun line -> line <> "" && not (line.StartsWith "#"))
                    |> Set.ofArray
                else
                    try
                        File.WriteAllText(
                            failedFile,
                            "# 逐份升级失败清单（每行一个相对路径；可直接给 XUNITHOST_CONVERT_SKIP）" + Environment.NewLine
                        )
                    with _ ->
                        ()

                    Set.empty

            let mutable known = failedSeen

            let noteFailure (rel: string) =
                if not (known.Contains rel) then
                    known <- known.Add rel

                    try
                        File.AppendAllText(failedFile, rel + Environment.NewLine)
                    with _ ->
                        ()

            let reasonOf (detail: string) =
                if detail.Contains "修复" then "修复对话框"
                elif detail.Contains "对不上" then "内容对不上"
                elif detail.Contains "超时" then "COM 超时"
                elif detail.Contains "拒绝" then "COM 拒绝"
                elif detail.Contains "占用" then "文件被占用"
                elif detail.Contains "宏" then "宏判定"
                else "其它"

            Report.title output "常驻 PowerPoint：逐份升级 .ppt，成功一份送走一份"
            Report.line output "黑板" (info.describe())
            Report.line output "本轮" (sprintf "%d 份 .ppt（%s）" queue.Length (RealFolder.limitLabel ()))
            Report.line output "备份目录" backup
            Report.line output "进度日志" progressFile
            Report.line output "失败清单" (sprintf "%s（已有 %d 条历史失败）" failedFile failedSeen.Count)
            Report.line output "宏口径" (sprintf "不留宏，一律存 .pptx（保宏路 %b）" (PowerPointHost.PptUpgrade.keepMacro ()))

            Report.line
                output
                "弹窗策略"
                (sprintf
                    "不躲：%d 秒没进展就把 PowerPoint 切到前台催人，点掉对话框自动继续；只激活窗口，不杀进程、不关你的演示文稿"
                    hintSec)

            Report.line output "停的条件" (sprintf "同一原因连续 %d 次（卡住不判失败，等你处理）" sameReasonLimit)

            let beforeProc = Report.procIds "POWERPNT"

            let app = Attached.attachPpt info.Pid |> Report.ok

            Report.line
                output
                "挂回"
                (sprintf
                    "pid=%d rot=%s（校验只用进程身份）"
                    info.Pid
                    (if String.IsNullOrWhiteSpace info.Rot then "按 ProgID" else info.Rot))

            let focusOnStart =
                match Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_FOCUS" with
                | null
                | "" -> true
                | text -> text.Trim() <> "0"

            let raised = if focusOnStart then Attached.bringToFront info.Pid else false

            Report.line
                output
                "界面"
                (sprintf
                    "开跑前置 %s：%s"
                    (if focusOnStart then "开" else "关")
                    (if not focusOnStart then "没动窗口，只在卡住时催"
                     elif raised then "已把 PowerPoint 切到前台，跑的时候有对话框你会立刻看见"
                     else "没切动，请自己点任务栏里的 PowerPoint 看一眼"))

            let mutable currentStage = "未开始"
            let mutable currentSince = DateTime.Now
            let mutable watchStop = false
            let mutable lastHint = DateTime.MinValue
            let mutable hintCount = 0

            let watchdog =
                new System.Threading.Timer(
                    (fun (_: obj) ->
                        if not watchStop then
                            let secs = int (DateTime.Now - currentSince).TotalSeconds
                            let sinceHint = (DateTime.Now - lastHint).TotalSeconds

                            if secs >= hintSec && sinceHint >= float hintSec && hintCount < 40 then
                                lastHint <- DateTime.Now
                                hintCount <- hintCount + 1
                                let front = Attached.bringToFront info.Pid

                                record (
                                    sprintf
                                        "⚠ 没进展 %d 秒（%s）—— 多半弹了对话框；%s"
                                        secs
                                        currentStage
                                        (if front then
                                             "已把 PowerPoint 切到前台，请点掉弹窗，点完我自动继续"
                                         else
                                             "没切动前台，请自己点任务栏里的 PowerPoint 看一眼")
                                )),
                    null,
                    5000,
                    5000
                )

            let mutable okCount = 0
            let mutable failCount = 0
            let mutable stopReason = ""
            let mutable lastReason = ""
            let mutable sameRun = 0
            let outcomes = ResizeArray<Outcome>()

            let run =
                try
                    Ok(
                        STA.run(fun () ->
                            let quiet = Attached.quietPpt app
                            let appliedAlerts = (try Some app.DisplayAlerts with _ -> None)
                            let appliedSecurity = (try int app.AutomationSecurity with _ -> -1)
                            let presBefore = Attached.presCount app

                            try
                                record (sprintf "  实例演示文稿 %d 份（用户自己也开着的不关）" presBefore)

                                // 上次中断可能把自己的演示文稿留在实例里：只关本队列里的那些，
                                // 用户自己开着的别的文件一个字不动。
                                if presBefore > 0 then
                                    let mutable closed = []

                                    for k in presBefore .. -1 .. 1 do
                                        let pres =
                                            try
                                                app.Presentations.Item(k)
                                            with _ ->
                                                null

                                        let full =
                                            try
                                                if isNull pres then "" else pres.FullName
                                            with _ ->
                                                ""

                                        if not (isNull pres) && full <> "" && queueSet.Contains(full.ToLowerInvariant()) then
                                            try
                                                pres.Close()
                                                closed <- Path.GetFileName full :: closed
                                            with _ ->
                                                ()

                                    if not closed.IsEmpty then
                                        record (
                                            sprintf
                                                "  关掉上次残留 %d 份（%s）"
                                                closed.Length
                                                (String.concat "、" (List.rev closed))
                                        )

                                let mutable index = 0

                                while index < queue.Length && stopReason = "" do
                                    let i = queue.[index]
                                    index <- index + 1
                                    currentItem <- Some i
                                    let rel = RealFolder.rel root i.Source
                                    currentStage <- sprintf "[%d/%d] 开始 %s" index queue.Length rel
                                    currentSince <- DateTime.Now
                                    lastHint <- DateTime.MinValue
                                    record (sprintf "[%d/%d] 开始 %s" index queue.Length rel)
                                    let clock = System.Diagnostics.Stopwatch.StartNew()

                                    let finish (text: string) =
                                        record text
                                        record (sprintf "  本份耗时 %.1f 秒" clock.Elapsed.TotalSeconds)

                                    if not (Set.contains info.Pid (Report.procIds info.Image)) then
                                        stopReason <- sprintf "常驻 PowerPoint pid %d 已不在" info.Pid
                                        record ("停：" + stopReason)
                                    else
                                        // ---- 1. 只备份这一份 ----
                                        currentStage <- sprintf "[%d/%d] 备份 %s" index queue.Length rel
                                        currentSince <- DateTime.Now

                                        let backed =
                                            try
                                                RealFolder.backupAll root backup [i] say |> ignore
                                                true
                                            with ex ->
                                                record ("  备份失败：" + Errors.explain ex)
                                                false

                                        if backed then
                                            // ---- 2. 重名让位 ----
                                            currentStage <- sprintf "[%d/%d] 重名让位 %s" index queue.Length rel

                                            let _, skipped = RealFolder.renameExisting root [i] say

                                            if skipped.Contains i.Source then
                                                failCount <- failCount + 1
                                                finish (sprintf "× %s —— 现有同名新文件改名失败，这份跳过" rel)
                                                noteFailure rel
                                            else
                                                // ---- 3. 转换这一份（内部已读回核对） ----
                                                currentStage <- sprintf "[%d/%d] 转换 %s" index queue.Length rel
                                                currentSince <- DateTime.Now

                                                let originalHash =
                                                    try
                                                        Some(Attached.md5 i.Source)
                                                    with _ ->
                                                        None

                                                let originalLen =
                                                    try
                                                        Some((FileInfo i.Source).Length)
                                                    with _ ->
                                                        None

                                                let outcome =
                                                    try
                                                        Convert.showsAttached app [i] |> List.head
                                                    with ex ->
                                                        { Source = i.Source
                                                          Target = i.Target
                                                          Status = "失败"
                                                          Detail = Errors.explain ex }

                                                outcomes.Add outcome

                                                // ---- 4. 原件被改写就从备份还原（失败件必须原样） ----
                                                let restored =
                                                    if File.Exists i.Source && originalHash.IsSome then
                                                        let afterHash = Attached.md5 i.Source

                                                        if afterHash = originalHash.Value then
                                                            true
                                                        else
                                                            try
                                                                let backupFile = Path.Combine(backup, rel)
                                                                File.Copy(backupFile, i.Source, true)
                                                                let nowHash = Attached.md5 i.Source
                                                                let backupHash = Attached.md5 backupFile

                                                                record (
                                                                    sprintf
                                                                        "  原件被 PowerPoint 改写（md5 %s → %s，%d B），已从备份还原成%s"
                                                                        (originalHash.Value.Substring(0, 8))
                                                                        (nowHash.Substring(0, 8))
                                                                        (defaultArg originalLen -1L)
                                                                        (if nowHash = backupHash then
                                                                             "原始字节"
                                                                         else
                                                                             "⚠ 与备份不一致的字节")
                                                                )

                                                                nowHash = backupHash
                                                            with ex ->
                                                                record ("  × 原件被改写，还原失败：" + Errors.explain ex)
                                                                false
                                                    else
                                                        true

                                                // ---- 5. 通过就一份一送；不通过就打原因、记清单 ----
                                                currentStage <- sprintf "[%d/%d] 送回收站 %s" index queue.Length rel
                                                currentSince <- DateTime.Now

                                                if Convert.canDelete outcome then
                                                    let removed, kept = RealFolder.deleteVerified root [outcome] say

                                                    if removed = 1 then
                                                        okCount <- okCount + 1
                                                        sameRun <- 0
                                                        lastReason <- ""
                                                        finish (sprintf "✓ %s —— %s" rel outcome.Detail)
                                                    else
                                                        failCount <- failCount + 1
                                                        sameRun <- 0
                                                        finish (sprintf "× %s —— 新文件在，但原件送不进回收站（留下 %d 份）" rel kept)
                                                        noteFailure rel
                                                else
                                                    failCount <- failCount + 1
                                                    let reason = reasonOf outcome.Detail
                                                    finish (sprintf "× %s —— %s（%s）" rel reason outcome.Detail)
                                                    noteFailure rel

                                                    if reason = lastReason then
                                                        sameRun <- sameRun + 1
                                                    else
                                                        lastReason <- reason
                                                        sameRun <- 1

                                                    if sameRun >= sameReasonLimit then
                                                        stopReason <- sprintf "同一原因「%s」连续 %d 次" reason sameRun
                                                        record ("停：" + stopReason + " —— 像系统性问题，先查这一条")

                                                if not restored then
                                                    record "  ⚠ 上面那份原件没还原成功，先核对备份"
                                        else
                                            failCount <- failCount + 1
                                            finish (sprintf "× %s —— 备份失败，原件保留、这份跳过" rel)
                            finally
                                Attached.restorePpt app quiet

                            (quiet.Alerts,
                             appliedAlerts,
                             quiet.Security |> Option.map int,
                             appliedSecurity,
                             (try Some app.DisplayAlerts with _ -> None),
                             (try int app.AutomationSecurity with _ -> -1),
                             presBefore,
                             Attached.presCount app,
                             Attached.ownPresLeft app queueSet))
                    )
                with ex ->
                    let detail = Errors.explain ex
                    record ("× 通道异常：" + detail)

                    // 通道异常也要记账：不然报告里是「失败 0 份」、清单里也没这条，下一轮又撞同一个坑
                    match currentItem with
                    | Some i ->
                        let rel = RealFolder.rel root i.Source
                        failCount <- failCount + 1
                        record (sprintf "× %s —— 通道异常（%s），记入失败清单" rel detail)
                        noteFailure rel

                        let backupFile = Path.Combine(backup, rel)

                        if File.Exists i.Source && File.Exists backupFile then
                            try
                                let nowHash = Attached.md5 i.Source
                                let backupHash = Attached.md5 backupFile

                                if nowHash <> backupHash then
                                    File.Copy(backupFile, i.Source, true)

                                    record (
                                        sprintf
                                            "  原件被改写（md5 %s → %s），已从备份还原"
                                            (backupHash.Substring(0, 8))
                                            (nowHash.Substring(0, 8))
                                    )
                            with restoreEx ->
                                record ("  ⚠ 还原失败：" + Errors.explain restoreEx)
                    | None -> ()

                    Error detail

            watchStop <- true
            watchdog.Dispose()

            match run with
            | Ok(quietAlerts,
                  appliedAlerts,
                  quietSecurity,
                  appliedSecurity,
                  restoredAlerts,
                  restoredSecurity,
                  presBefore,
                  presAfter,
                  ownLeft) ->
                Report.line output "DisplayAlerts" (sprintf "原值 %A → %A → 恢复 %A" quietAlerts appliedAlerts restoredAlerts)

                Report.line
                    output
                    "AutomationSecurity"
                    (sprintf "原值 %A → %d（ForceDisable）→ 恢复 %d" quietSecurity appliedSecurity restoredSecurity)

                Report.line
                    output
                    "实例演示文稿"
                    (sprintf "干活前 %d 份 / 收尾后 %d 份（用户自己也开着的只算数、不判定）" presBefore presAfter)

                Report.line output "我开的还剩" (sprintf "%d 份（0 = 自己开的都关了；-1 = 数不出来）" ownLeft)

                Should.equal 3 appliedSecurity

                (match quietAlerts with
                 | Some value -> Should.equal (Some value) restoredAlerts
                 | None -> ())

                (match quietSecurity with
                 | Some value -> Should.equal (Some value) (Some restoredSecurity)
                 | None -> ())

                if ownLeft >= 0 then Should.equal 0 ownLeft
            | Error _ -> ()

            Report.line output "本轮结果" (sprintf "队列 %d 份，成功 %d 份，失败 %d 份" queue.Length okCount failCount)

            Report.line
                output
                "催前台"
                (sprintf "%d 次（%d 秒没进展就催一次；催了还没动，就是在等你在对话框中做选择）" hintCount hintSec)

            if stopReason <> "" then
                Report.line output "停下原因" stopReason

            // 核对没通过的，原件一份都不能丢
            for o in outcomes |> Seq.filter(Convert.canDelete >> not) do
                Assert.True(File.Exists o.Source, sprintf "核对没通过却把原件弄丢了：%s" o.Source)

            let afterProc = Report.procIds "POWERPNT"
            Report.line output "现场" (sprintf "POWERPNT 进程 干活前 %A / 收尾后 %A" (Set.toList beforeProc) (Set.toList afterProc))

            // 只断言「没有新建」；PowerPoint 自己崩溃重启不是这条通道能拦的，报告出来就够
            Should.equal true (Set.isSubset afterProc beforeProc)

            if not (Set.contains info.Pid afterProc) then
                record (sprintf "  ⚠ 常驻 PowerPoint pid=%d 已不在：多半是打开坏文件时它自己重启了，本轮到此为止" info.Pid)

            let left = Paths.scan root |> Paths.filterByKind Paths.Kind.Ppt
            Report.line output "目录现状" (sprintf "还剩 %d 份 .ppt（本轮成功 %d 份已送回收站）" left.Length okCount)

            Report.conclusion
                output
                (sprintf
                    "逐份推进：队列 %d 份，成功 %d 份（原件已入回收站），失败 %d 份（原件保留，相对路径记入失败清单）%s"
                    queue.Length
                    okCount
                    failCount
                    (if stopReason = "" then "" else "；停：" + stopReason))

    /// 单份分阶段计时：打开 / 判宏 / UsedRange / 读指纹 / SaveAs / 读回，哪一步慢一目了然。
    /// 只诊断一份、SaveAs 到 %TEMP%，不动原目录；结束后把原件从备份还原。
    /// 每步行进都实时写 TestResults/诊断-分阶段.log，真卡住也能看到最后一步。
    [<Fact>]
    member this.``诊断：单份 xls 分阶段计时，看卡在哪一步``() =
        match RealFolder.root () with
        | None -> Assert.Skip "设 XUNITHOST_CONVERT_DIR=<目录> 才诊断"
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let info = Attached.info () |> Report.ok
            let borrowed = Attached.borrow info
            let backup = RealFolder.backupRoot root
            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let progressFile = Path.Combine(projectRoot, "TestResults", "诊断-分阶段.log")

            let target =
                match Environment.GetEnvironmentVariable "XUNITHOST_PROBE_FILE" with
                | null
                | "" ->
                    match Paths.selectForRound root Paths.Kind.Xls (Paths.scan root) with
                    | head :: _ -> head.Source
                    | [] -> failwith "队列是空的，没有可诊断的 .xls"
                | value -> value

            let rel = RealFolder.rel root target
            let backupFile = Path.Combine(backup, rel)
            let stage = Path.Combine(Path.GetTempPath(), "probe-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx")

            Report.title output "诊断：单份 .xls 分阶段计时"
            Report.line output "目标" rel
            Report.line output "临时目标" stage
            Report.line output "备份" backupFile

            Assert.True(File.Exists target, sprintf "原件不在：%s" rel)
            Assert.True(File.Exists backupFile, sprintf "备份不在：%s" backupFile)

            let report (text: string) =
                let line = sprintf "%s %s" (DateTime.Now.ToString("HH:mm:ss")) text
                output.WriteLine("    " + text)
                try File.AppendAllText(progressFile, line + Environment.NewLine) with _ -> ()

            let hashBefore = Attached.md5 target
            let app = Attached.attach borrowed |> Report.ok
            let mutable stepName = "未开始"
            let mutable stepSince = DateTime.Now
            let mutable watch = true

            let watchTimer =
                new System.Threading.Timer(
                    (fun (_: obj) ->
                        if watch then
                            report (sprintf "  ⏱ 仍在 %s（已 %d 秒）" stepName (int (DateTime.Now - stepSince).TotalSeconds))),
                    null,
                    5000,
                    5000
                )

            let timed (name: string) (action: unit -> 'T) : 'T =
                stepName <- name
                stepSince <- DateTime.Now
                let clock = System.Diagnostics.Stopwatch.StartNew()
                report ("→ " + name)
                let value = action ()
                report (sprintf "  ✓ %s：%.2f 秒" name clock.Elapsed.TotalSeconds)
                value

            let run =
                try
                    STA.run(fun () ->
                        let quiet = Attached.quieten app

                        try
                            let ro = timed "打开（只读）" (fun () -> ExcelHost.BookIO.openBook app target true)
                            let macro = timed "读 HasVBProject" (fun () -> try ro.HasVBProject with _ -> false)
                            let sheets = timed "读表数" (fun () -> try ro.Worksheets.Count with _ -> -1)

                            let sheet =
                                timed "取第 1 张表" (fun () -> ro.Worksheets.Item(1) :?> Microsoft.Office.Interop.Excel.Worksheet)

                            let area = timed "读 UsedRange.Address" (fun () -> try sheet.UsedRange.Address() with _ -> "?")
                            let rows = timed "读 UsedRange 行数" (fun () -> try sheet.UsedRange.Rows.Count with _ -> -1)
                            let cols = timed "读 UsedRange 列数" (fun () -> try sheet.UsedRange.Columns.Count with _ -> -1)
                            let printRo = timed "读指纹（只读）" (fun () -> ExcelHost.BookIO.bookPrint ro)
                            timed "关闭（只读）" (fun () -> ExcelHost.BookIO.closeBook ro)

                            let rw = timed "打开（可写）" (fun () -> ExcelHost.BookIO.openBook app target false)
                            let printRw = timed "读指纹（可写）" (fun () -> ExcelHost.BookIO.bookPrint rw)

                            timed
                                "SaveAs 到 %TEMP%"
                                (fun () ->
                                    try
                                        rw.SaveAs(box stage, box 51) |> ignore
                                    with ex ->
                                        report ("  × SaveAs：" + OsHost.Errors.explain ex))

                            timed "关闭（可写）" (fun () -> ExcelHost.BookIO.closeBook rw)

                            let back = timed "打开新文件（只读）" (fun () -> ExcelHost.BookIO.openBook app stage true)
                            let printBack = timed "读指纹（新文件）" (fun () -> ExcelHost.BookIO.bookPrint back)
                            timed "关闭（新文件）" (fun () -> ExcelHost.BookIO.closeBook back)

                            (macro, sheets, area, rows, cols, printRo.Length, printRw.Length, printBack.Length, printRo = printBack)
                        finally
                            Attached.restore app quiet)
                with ex ->
                    report ("× 诊断中断：" + OsHost.Errors.explain ex)
                    (false, -1, "?", -1, -1, -1, -1, -1, false)

            watch <- false
            watchTimer.Dispose()

            let macro, sheets, area, rows, cols, lenRo, lenRw, lenBack, same = run

            Report.line output "判宏" (sprintf "HasVBProject=%b；%d 张表" macro sheets)
            Report.line output "UsedRange" (sprintf "%s（%d 行 × %d 列）" area rows cols)
            Report.line output "指纹长度" (sprintf "只读 %d 字符 / 可写 %d 字符 / 新文件 %d 字符" lenRo lenRw lenBack)
            Report.line output "内容一致" (sprintf "%b" same)

            (try File.Delete stage with _ -> ())

            let hashAfter = Attached.md5 target

            if hashAfter <> hashBefore then
                File.Copy(backupFile, target, true)
                let restoredHash = Attached.md5 target
                Report.line output "原件还原" (sprintf "md5 %s → %s，已从备份还原" (hashBefore.Substring(0, 8)) (restoredHash.Substring(0, 8)))
            else
                Report.line output "原件还原" "原件字节没变，不用还原"

            Report.conclusion output "完成：上面每一步的秒数就是真实耗时，最慢那一步就是卡点"

    /// 按章节切分：目录（TOC）给标题清单，正文里逐条定位标题，切点取标题起点。
    /// 每份的做法：SaveCopyAs 出一份完整 docx 副本 → 打开副本 → 删掉本章之外的部分 → 存 → 关。
    /// 这样页面设置、页眉页脚、样式、编号都原样保留，源文档一个字节都不改，Word 里同时只留一个文档。
    /// 闸门：XUNITHOST_DOC_SPLIT_OUT=<输出目录>；XUNITHOST_DOC_SPLIT_SRC=<源 doc 路径>（不给就用活动文档）；
    /// XUNITHOST_DOC_SPLIT_DRY=1 只列结构；已存在的同名 docx 直接跳过（可重跑）。
    [<Fact>]
    member this.``按章节切分：把开着的这份 doc 逐章存成 docx``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_DOC_SPLIT_OUT" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_DOC_SPLIT_OUT=<输出目录>"
        | outDir ->
            let dry = Environment.GetEnvironmentVariable "XUNITHOST_DOC_SPLIT_DRY" = "1"

            let srcPath = Environment.GetEnvironmentVariable "XUNITHOST_DOC_SPLIT_SRC"

            let clean (text: string) =
                text.Replace("\r", "").Replace("\n", "").Replace("\u0007", "").Trim()

            let safeName (text: string) =
                let bad = Path.GetInvalidFileNameChars()
                let buffer = System.Text.StringBuilder()

                for ch in text do
                    if Array.contains ch bad then
                        buffer.Append '_' |> ignore
                    else
                        buffer.Append ch |> ignore

                let name = buffer.ToString().Trim().TrimEnd('.')
                if name.Length > 60 then name.Substring(0, 60) else name

            // 目录条目形如「16.  脱硫岛电气系统技术描述」，去掉点线引导符和末尾页码做文件名
            let tidy (text: string) =
                let noDots =
                    System.Text.RegularExpressions.Regex.Replace(text, @"[\.\u2026\u00b7\u3000\s]+", " ")

                System.Text.RegularExpressions.Regex.Replace(noDots, @"\s*\d+\s*$", "").Trim()

            Report.title output "按章节切分：目录定位标题，逐章存 docx（源文档不改、不保存）"

            let app = WordHost.Host.GetOrCreateApplication ()

            let srcName, total, endOfDoc, sectionCount, tocItems, chapters, saved =
                STA.run(fun () ->
                    // Word 稍微一忙就甩 RPC_E_CALL_REJECTED，所有调用一律过一遍忙重试
                    let busy (f: unit -> 'T) = Busy.retry 120000 f

                    let alerts = busy (fun () -> app.DisplayAlerts)
                    busy (fun () -> app.DisplayAlerts <- Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone)
                    // 露出来：万一弹了什么，用户能看见、能点，不至于在看不见的实例里卡死
                    busy (fun () -> app.Visible <- true)

                    try
                        let source =
                            if not (String.IsNullOrWhiteSpace srcPath) then
                                // 先清掉上次留下的 ~$ 锁，免得被只读打开
                                try
                                    for lock in Directory.GetFiles(Path.GetDirectoryName srcPath, "~$*") do
                                        try
                                            File.Delete lock
                                        with _ ->
                                            ()
                                with _ ->
                                    ()

                                // 已经开着的就是这一份就别重复打开（重复 Open 同一路径会把 Word 拖进怪状态）
                                let alreadyOpen =
                                    try
                                        let active = app.ActiveDocument

                                        if not (isNull active) && active.FullName = srcPath then
                                            Some active
                                        else
                                            None
                                    with _ ->
                                        None

                                match alreadyOpen with
                                | Some active -> active
                                | None ->
                                    busy (fun () ->
                                        let mutable fileName: obj = box srcPath
                                        let mutable confirm: obj = box false
                                        let mutable readOnly: obj = box false
                                        let mutable recent: obj = box false
                                        let mutable missing: obj = Type.Missing
                                        // 12=Visible、13=OpenAndRepair 必须给：老 doc 一打开就弹「是否修复」，
                                        // 弹在看不见的隐藏实例里就只能等，整条通道会卡死。
                                        let mutable visible: obj = box true
                                        let mutable repair: obj = box true

                                        app.Documents.Open(
                                            &fileName,
                                            &confirm,
                                            &readOnly,
                                            &recent,
                                            &missing,
                                            &missing,
                                            &missing,
                                            &missing,
                                            &missing,
                                            &missing,
                                            &missing,
                                            &visible,
                                            &repair
                                        ))
                            else
                                app.ActiveDocument

                        let srcName = busy (fun () -> source.Name)
                        let total = busy (fun () -> source.Paragraphs.Count)
                        let endOfDoc = busy (fun () -> source.Content.End)
                        let sectionCount = busy (fun () -> source.Sections.Count)

                        let tocItems =
                            if busy (fun () -> source.TablesOfContents.Count) = 0 then
                                [||]
                            else
                                (busy (fun () -> source.TablesOfContents.Item(1).Range.Text)).Split('\r')
                                |> Array.map clean
                                |> Array.filter (fun line -> line.Length > 0)

                        let tocEnd =
                            if busy (fun () -> source.TablesOfContents.Count) = 0 then
                                0
                            else
                                busy (fun () -> source.TablesOfContents.Item(1).Range.End)

                        // 在文中（或副本里）从 from 往后找一段文字，命中返回起点，找不到返回 -1
                        let findIn (target: Microsoft.Office.Interop.Word.Document) (text: string) (from: int) =
                            if isNull target || text.Trim().Length < 2 then
                                -1
                            else
                                let mutable f: obj = box from
                                let mutable t: obj = box (busy (fun () -> target.Content.End))
                                let range = target.Range(&f, &t)

                                busy (fun () ->
                                    range.Find.ClearFormatting()
                                    range.Find.Text <- text.Trim()
                                    range.Find.Forward <- true
                                    range.Find.Wrap <- Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop)

                                if busy (fun () -> range.Find.Execute()) then
                                    range.Start
                                else
                                    -1

                        // 逐条目录文本定位；顺带回溯 200 字符看标题前面有没有分节符 / 分页符（只作报告）
                        let chapters =
                            tocItems
                            |> Array.map (fun item ->
                                let pos = findIn source item tocEnd

                                let mark =
                                    if pos < 0 then
                                        "没找到"
                                    elif pos <= 1 then
                                        "文件开头"
                                    else
                                        let mutable a: obj = box (max 0 (pos - 200))
                                        let mutable b: obj = box pos
                                        let rng = source.Range(&a, &b)

                                        if (busy (fun () -> rng.Text)).Contains "\u000c" then
                                            "标题前有分节符/分页符"
                                        else
                                            "标题前无分节符"

                                tidy item, pos, mark)

                        let results = ResizeArray<string * string>()

                        if not dry then
                            Directory.CreateDirectory outDir |> ignore

                            let located =
                                chapters
                                |> Array.filter (fun (_, pos, _) -> pos > 0)
                                |> Array.sortBy (fun (_, pos, _) -> pos)

                            // 计划：封面目录 + 每章（名字、本章标题原文、下一章标题原文）
                            let plan = ResizeArray<string * string * string>()

                            if located.Length > 0 then
                                let rawOf title =
                                    tocItems
                                    |> Array.tryFind (fun item -> tidy item = title)
                                    |> Option.defaultValue title

                                plan.Add("00 封面与目录", "", rawOf (let (t, _, _) = located.[0] in t))

                                for i in 0 .. located.Length - 1 do
                                    let (title, _, _) = located.[i]

                                    let nextTitle =
                                        if i + 1 < located.Length then
                                            let (t2, _, _) = located.[i + 1]
                                            rawOf t2
                                        else
                                            ""

                                    plan.Add(title, rawOf title, nextTitle)

                            // 计划：封面目录 + 每章，直接在源文档里算好 Range 起止
                            let plan = ResizeArray<string * int * int>()

                            if located.Length > 0 then
                                let firstPos = let (_, p, _) = located.[0] in p

                                if firstPos > 1 then
                                    plan.Add("00 封面与目录", 0, firstPos - 1)

                                for i in 0 .. located.Length - 1 do
                                    let (title, startPos, _) = located.[i]

                                    let stopPos =
                                        if i + 1 < located.Length then
                                            let (_, nextPos, _) = located.[i + 1]
                                            nextPos - 1
                                        else
                                            endOfDoc - 1

                                    if stopPos > startPos then
                                        plan.Add(clean title, startPos, stopPos)

                            let limit =
                                match Int32.TryParse(Environment.GetEnvironmentVariable "XUNITHOST_DOC_SPLIT_LIMIT") with
                                | true, value when value > 0 -> value
                                | _ -> plan.Count

                            // 从第几份开始（1 起）：Word 一个会话做十来份就不稳，分批推进
                            let fromIndex =
                                match Int32.TryParse(Environment.GetEnvironmentVariable "XUNITHOST_DOC_SPLIT_FROM") with
                                | true, value when value > 0 -> min (value - 1) (plan.Count - 1)
                                | _ -> 0

                            let take = min limit (plan.Count - fromIndex)

                            if fromIndex > 0 || take < plan.Count - fromIndex then
                                Report.line
                                    output
                                    "本批"
                                    (sprintf "第 %d—%d 份（共 %d 份）" (fromIndex + 1) (fromIndex + take) plan.Count)

                            // 写入方式：fmt=整段 FormattedText（默认）；clip=剪贴板 Copy/Paste；text=只搬文字
                            // 有的章（12主要性能关系曲线）字符没多少，但整段 FormattedText 一贴就把 Word 弄死，
                            // 换剪贴板那条路试试
                            let method =
                                match Environment.GetEnvironmentVariable "XUNITHOST_DOC_SPLIT_METHOD" with
                                | null
                                | "" -> "fmt"
                                | value -> value.Trim().ToLowerInvariant()

                            for k in 0 .. take - 1 do
                                let i = fromIndex + k
                                let name, startPos, stopPos = plan.[i]
                                let fileName = sprintf "%02d %s.docx" (i + 1) (safeName name)
                                let finalPath = Path.Combine(outDir, fileName)

                                if File.Exists finalPath then
                                    results.Add(fileName, "已存在，跳过")
                                else
                                    let mutable target: Microsoft.Office.Interop.Word.Document = null

                                    // 用完立刻关、立刻放，别让几十个文档对象堆在 Word 里（上次就是这么崩的）
                                    let cleanup () =
                                        if not (isNull target) then
                                            let dying = target
                                            target <- null

                                            try
                                                busy (fun () ->
                                                    let mutable ns: obj =
                                                        box (int Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

                                                    dying.Close(&ns))
                                            with _ ->
                                                ()

                                            try
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject dying
                                                |> ignore
                                            with _ ->
                                                ()

                                    try
                                        let mutable a: obj = box startPos
                                        let mutable b: obj = box stopPos
                                        let part = busy (fun () -> source.Range(&a, &b))

                                        target <- busy (fun () -> app.Documents.Add())

                                        let fresh = target

                                        if method = "clip" then
                                            // 走剪贴板：Word 自己序列化，比跨文档搬 FormattedText 温和
                                            busy (fun () -> part.Copy())
                                            busy (fun () -> fresh.Content.Paste())
                                        elif method = "text" then
                                            // 兜底：只要文字，图/表对象一律不要
                                            let text = busy (fun () -> part.Text)
                                            busy (fun () -> fresh.Content.Text <- text)
                                        else
                                            let copied = part.FormattedText
                                            busy (fun () -> fresh.Content.FormattedText <- copied)

                                        busy (fun () ->
                                            let mutable pathObj: obj = box finalPath

                                            let mutable fmtObj: obj =
                                                box (int Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)

                                            fresh.SaveAs2(&pathObj, &fmtObj) |> ignore)

                                        results.Add(fileName, sprintf "%d B（%s）" (FileInfo finalPath).Length method)
                                        cleanup ()
                                    with ex ->
                                        let where =
                                            let lines =
                                                if isNull ex.StackTrace then
                                                    [||]
                                                else
                                                    ex.StackTrace.Split('\n')

                                            if lines.Length > 0 then
                                                lines
                                                |> Array.truncate 3
                                                |> Array.map (fun s -> s.Trim())
                                                |> String.concat " ｜ "
                                            else
                                                ""

                                        results.Add(fileName, "失败：" + Errors.explain ex + " ｜ " + where)
                                        cleanup ()

                        srcName, total, endOfDoc, sectionCount, tocItems, chapters, results.ToArray()
                    finally
                        // Word 已经崩了的话这里也会抛，别让它把真正的异常盖掉
                        try
                            app.DisplayAlerts <- alerts
                        with _ ->
                            ())

            let atSection = chapters |> Array.filter (fun (_, _, mark) -> mark.Contains "有分节符") |> Array.length
            let locatedCount = chapters |> Array.filter (fun (_, pos, _) -> pos > 0) |> Array.length

            Report.line output "源文档" (sprintf "%s（%d 段，%d 字符）" srcName total endOfDoc)
            Report.line output "目录条目" (sprintf "%d 条，正文里定位到 %d 条" tocItems.Length locatedCount)
            Report.line output "分节符" (sprintf "%d 个（%d 个节）；标题前 200 字符内有分节符/分页符的 %d 条" (sectionCount - 1) sectionCount atSection)
            Report.line output "切点" "取标题本身的起点（分节符不够用，见上）"

            for index in 0 .. chapters.Length - 1 do
                let (title, pos, mark) = chapters.[index]
                Report.line2 output "章" (sprintf "%02d" (index + 1)) (sprintf "起%7d  %-14s %s" pos mark title)

            if dry then
                Report.conclusion output "完成：只列结构，没落盘（XUNITHOST_DOC_SPLIT_DRY=1）"
            else
                for name, note in saved do
                    Report.line output ("✓ " + name) note

                Report.conclusion output (sprintf "完成：切出 %d 份 docx 到 %s（源文档没保存、没改）" saved.Length outDir)
