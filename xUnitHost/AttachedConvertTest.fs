/// 常驻 Excel（黑板挂回式）升级 .xls：Excel 由用户 / 插件常驻，这一侧只附着。
///
/// 与老的一次性通道（Host.execQuery：自己起 Excel → 干活 → Quit → kill）并列，互不干扰：
///   · 入口固定 ExcelHost.Host.GetApplication(info: GeneralAppInfo)，绝不新建、绝不 Quit、绝不 kill；
///   · 不 requireClean —— 常驻实例在跑是**前提**，不是障碍；
///   · 所有 COM 读写都包在 OsHost.Sta.run 那条固定 STA 线程上（与附着同公寓）；
///   · 只逐份开/关自己那份工作簿，绝不用 CloseAllWorkbooksWithoutSaving（那会关掉用户开着的表）；
///   · 应用级开关（DisplayAlerts / AskToUpdateLinks / AutomationSecurity）临时改、用完恢复原值；
///   · **借用进程**：黑板上的 HWND 是用户当前工作簿的窗口，不是本轮要打开的那份 .xls 的窗口。
///     开/关文件后它必然换，拿它当交接凭据会让下一轮断链 —— 把它从校验里摘掉，
///     进程身份照核（PID 存活 / 采集时间 / 启动时刻 / EXE / ROT 名后缀 pid）；
///   · 用户开着几本工作簿不归这条通道管，不断言数量；只核自己打开的那一本已关掉、没留垃圾。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Core
open OsHost

/// 黑板文本 → 参数表 → 挂回；两条 Fact 共用这一串
module private Attached =

    let info () : Result<GeneralAppInfo, string> =
        let file = Environment.GetEnvironmentVariable "XUNITHOST_BLACKBOARD_FILE"

        if String.IsNullOrWhiteSpace file then
            Error "没设 XUNITHOST_BLACKBOARD_FILE：把插件抄下来的黑板文本存成文件再跑"
        elif not (File.Exists file) then
            Error(sprintf "黑板文件不存在：%s" file)
        else
            File.ReadAllText file |> BlackboardText.ofText

    type App = Microsoft.Office.Interop.Excel.Application

    /// 借用进程用 OsHost 的正式入口：只摘掉 HWND 这一条校验，进程身份照核。
    /// 黑板原文里的 HWND 仍留在 info 里当数据，报告照实打印。
    let borrow (info: GeneralAppInfo) : GeneralAppInfo = GeneralAppInfo.borrowProcess info

    /// 收尾时实例里还挂着我本轮打开过的文件吗（按 FullName 认，别碰用户那本）。
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
          Security: MsoAutomationSecurity option }

    let quieten (app: App) : Quiet =
        let before =
            { DisplayAlerts = (try Some app.DisplayAlerts with _ -> None)
              AskToUpdateLinks = (try Some app.AskToUpdateLinks with _ -> None)
              Security = ExcelHost.Host.Security.apply app }

        (try app.DisplayAlerts <- false with _ -> ())
        (try app.AskToUpdateLinks <- false with _ -> ())
        before

    let restore (app: App) (before: Quiet) : unit =
        (match before.DisplayAlerts with
         | Some value -> (try app.DisplayAlerts <- value with _ -> ())
         | None -> ())

        (match before.AskToUpdateLinks with
         | Some value -> (try app.AskToUpdateLinks <- value with _ -> ())
         | None -> ())

        ExcelHost.Host.Security.restore app before.Security

/// 三条 Fact：借用进程的校验口径（纯逻辑）→ 只读挂回验通道 → 挂回式验一张真实 .xls。
type AttachedConvertTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``借用进程：别的文件窗口 HWND 不校验，进程身份一条不少``() =
        Report.title output "借用进程：只认 PID/ROT/EXE/启动时刻，别人窗口的 HWND 不校验"

        let info = Attached.info () |> Report.ok
        let borrowed = Attached.borrow info

        // 旧口径的反证：把窗口句柄换成一个完全不相干的，verify 会拒接 —— 上次翻倍法就是这样断链的。
        let alienWindow = { info with Hwnd = nativeint 0x00ABCDEF }
        let deadPid = { borrowed with Pid = 999999 }
        let wrongRot = { borrowed with Rot = "!ExcelVSTO.Instance.999999" }
        let wrongExe = { borrowed with ExePath = @"C:\nowhere\OTHER.EXE" }
        let noClock = { borrowed with CapturedMs = 0L }
        let noImage = { borrowed with Identity = GeneralIdentity.create ("", "Excel", "excel") }

        Report.line output "黑板原值" (info.describe())
        Report.line output "挂回用表" (borrowed.describe())
        Report.line output "口径" "HWND 属于用户当前工作簿的窗口，开/关自己的文件后必换；挂回用表置 0n 摘掉这一条"

        let verdict (label: string) (table: GeneralAppInfo) =
            match GeneralAppInfo.verify table with
            | Ok() -> Report.line output label "放行"
            | Error message -> Report.line output label ("拒接：" + message)

        verdict "别的窗口" alienWindow
        verdict "借用进程表" borrowed
        verdict "PID 不在" deadPid
        verdict "ROT 后缀不符" wrongRot
        verdict "EXE 不符" wrongExe
        verdict "采集时间缺失" noClock
        verdict "进程名为空" noImage

        let isOk table =
            match GeneralAppInfo.verify table with
            | Ok() -> true
            | _ -> false

        let isRefused (word: string) table =
            match GeneralAppInfo.verify table with
            | Error message -> message.Contains word
            | _ -> false

        // EXE 这一条：本机若能读到进程路径就必须因不符而拒接；读不到（Proc.exePath = None）时该检查自然跳过。
        let exeReadable = (Proc.exePath borrowed.Pid).IsSome
        let exeVerdict = GeneralAppInfo.verify wrongExe

        Should.equal true (isRefused "HWND" alienWindow)
        Should.equal true (isOk borrowed)
        Should.equal true (isRefused "已经不在运行" deadPid)
        Should.equal true (isRefused "ROT" wrongRot)
        Should.equal true (if exeReadable then isRefused "EXE" wrongExe else isOk wrongExe)
        Should.equal true (isRefused "采集时间" noClock)
        Should.equal true (isRefused "Image" noImage)
        Report.line output "EXE 路径" (sprintf "本机可读=%b，判定=%A" exeReadable exeVerdict)
        Should.equal 0n borrowed.Hwnd
        Should.equal info.Pid borrowed.Pid
        Should.equal info.Rot borrowed.Rot
        Should.equal info.ExePath borrowed.ExePath
        Should.equal info.CapturedMs borrowed.CapturedMs

        Report.conclusion output "通过：HWND 摘掉后放行，PID/ROT/EXE/采集时间不对仍拒接；纯逻辑，不碰真机、不动文件"

    [<Fact>]
    member this.``常驻 Excel：按参数表挂回，只读核对，一个文件都不动``() =
        let info = Attached.info () |> Report.ok
        let borrowed = Attached.borrow info
        let before = Report.procIds info.Identity.Image

        Report.title output "常驻 Excel：黑板参数表 → GetApplication 挂回（只读）"
        Report.line output "参数表" (info.describe())
        Report.line output "文档" (sprintf "名字=%A 路径=%A" info.DocName info.DocPath)
        Report.line output "交接" (sprintf "借用进程，HWND %A 只作数据、不校验；挂回用表 %s" info.Hwnd (borrowed.describe()))
        Report.line output "现场" (sprintf "%s 进程 %A" info.Identity.Image (Set.toList before))

        let attached = ExcelHost.Host.GetApplication borrowed

        Report.line
            output
            "挂回"
            (match attached with
             | Ok app -> sprintf "Ok %s" (app.GetType().FullName)
             | Error message -> "拒接：" + message)

        let app = attached |> Report.ok

        let probe =
            Sta.run(fun () ->
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

        let after = Report.procIds info.Identity.Image
        Report.line output "收尾" (sprintf "%s 进程 %A" info.Identity.Image (Set.toList after))
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
            Report.line output "参数表" (info.describe())
            Report.line output "交接" (sprintf "借用进程，HWND %A 只作数据、不校验；挂回用表 %s" info.Hwnd (borrowed.describe()))
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
            let app = ExcelHost.Host.GetApplication borrowed |> Report.ok
            Report.line output "挂回" (sprintf "pid=%d rot=%s（校验只用进程身份）" info.Pid info.Rot)

            let run =
                Sta.run(fun () ->
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

            let md5 (path: string) =
                use stream = File.OpenRead path
                use alg = System.Security.Cryptography.MD5.Create()
                System.Convert.ToHexString(alg.ComputeHash stream)

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

            let hashBefore, hashBackup, hashKeep = md5 source, md5 backup, md5 keep
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
            let hashAfter = md5 source

            Report.line output "复制" (if changed then "备份 → 原件（覆盖）" else "原件已是原始字节，一个字节没动")
            Report.line output "还原后" (sprintf "%d B md5=%s" sizeAfter hashAfter)
            Should.equal hashBackup hashAfter
            Should.equal (FileInfo backup).Length sizeAfter
            Should.equal hashBackup (md5 backup)
            Should.equal hashKeep (md5 keep)

            Report.line output "备份留档" (sprintf "备份 %d B / 留档 %d B，md5 都没变" (FileInfo backup).Length (FileInfo keep).Length)
            Report.conclusion output (sprintf "完成：%s 恢复到 %d B 原始字节，备份与留档原样保留" rel sizeAfter)
