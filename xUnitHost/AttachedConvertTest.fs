/// 常驻 Excel（黑板挂回式）升级 .xls：Excel 由用户 / 插件常驻，这一侧只附着。
///
/// 与老的一次性通道（Host.execQuery：自己起 Excel → 干活 → Quit → kill）并列，互不干扰：
///   · 入口固定 ExcelHost.Host.GetApplication(info: GeneralAppInfo)，绝不新建、绝不 Quit、绝不 kill；
///   · 不 requireClean —— 常驻实例在跑是**前提**，不是障碍；
///   · 所有 COM 读写都包在 OsHost.Sta.run 那条固定 STA 线程上（与附着同公寓）；
///   · 只逐份开/关自己那份工作簿，绝不用 CloseAllWorkbooksWithoutSaving（那会关掉用户开着的表）；
///   · 应用级开关（DisplayAlerts / AskToUpdateLinks / AutomationSecurity）临时改、用完恢复原值。
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

/// 两条 Fact：先只读挂回验通道，再挂回式验一张真实 .xls（含原件入回收站）。
type AttachedConvertTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``常驻 Excel：按黑板信息挂回，只读核对，一个文件都不动``() =
        let info = Attached.info () |> Report.ok
        let before = Report.procIds info.Identity.Image

        Report.title output "常驻 Excel：黑板参数表 → GetApplication 挂回（只读）"
        Report.line output "参数表" (info.describe())
        Report.line output "文档" (sprintf "名字=%A 路径=%A" info.DocName info.DocPath)
        Report.line output "现场" (sprintf "%s 进程 %A" info.Identity.Image (Set.toList before))

        let attached = ExcelHost.Host.GetApplication info

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

            // 这条通道只认 .xls；默认验 1 张，XUNITHOST_CONVERT_LIMIT 可放大。
            let take = Paths.limitFromEnv () |> Option.defaultValue 1
            let items = Paths.selectForRound root Paths.Kind.Xls (Paths.scan root) |> List.truncate take
            let backup = RealFolder.backupRoot root
            let say (text: string) = output.WriteLine("    " + text)
            let beforeProc = Report.procIds "EXCEL"

            Report.title output "常驻 Excel：挂回式升级 .xls，核对通过才把原件送回收站"
            Report.line output "参数表" (info.describe())
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
            let app = ExcelHost.Host.GetApplication info |> Report.ok
            Report.line output "挂回" (sprintf "pid=%d rot=%s" info.Pid info.Rot)

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

                    outcomes,
                    booksBefore,
                    booksAfter,
                    appliedAlerts,
                    appliedSecurity,
                    restoredAlerts,
                    restoredSecurity,
                    quiet)

            let outcomes, booksBefore, booksAfter, appliedAlerts, appliedSecurity, restoredAlerts, restoredSecurity, quiet =
                run

            Report.line output "工作簿数" (sprintf "干活前 %d / 收尾后 %d" booksBefore booksAfter)
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
            Should.equal booksBefore booksAfter

            for o in outcomes do
                let note = if Convert.canDelete o then "" else "　—— " + o.Detail
                Report.line output (RealFolder.mark o.Status) (RealFolder.rel root o.Source + note)

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
                    "完成：本轮 %d 份，核对通过 %d 份，原件入回收站 %d 份；常驻 Excel pid=%d 仍在、工作簿数没变、应用开关已恢复"
                    items.Length
                    good.Length
                    removed
                    info.Pid)
