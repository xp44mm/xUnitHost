/// P0 安全开关的验证：三家各起一次实例，读回真实的设置值；KIND 分轮是纯逻辑。
/// 用户 2026-11-07 00:14 反馈：Word 会弹「修复错误」、带宏文件会弹宏提示，都要人工关；
/// 这两条 Fact 把「不弹框、不启用宏、用完恢复原值」钉成可重跑的记录。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Core
open Microsoft.Office.Interop.PowerPoint

type ConvertSafetyTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``安全开关：Word 不弹修复框、不启用宏，用完恢复原值``() =
        Report.requireClean "WINWORD"

        let probe =
            WordHost.Host.execQuery(fun app ->
                let alerts = int app.DisplayAlerts
                let security = int app.AutomationSecurity
                let confirm = app.Options.ConfirmConversions
                // restore 能不能真的改回去：手动恢复成 Low，读回来，再设回 ForceDisable
                WordHost.Host.Security.restore app (Some MsoAutomationSecurity.msoAutomationSecurityLow)
                let restored = int app.AutomationSecurity
                WordHost.Host.Security.apply app |> ignore
                alerts, security, confirm, restored)
            |> Report.ok

        let alerts, security, confirm, restored = probe

        Report.title output "安全开关：Word 不弹修复框、不启用宏"
        Report.line output "打开方式" "Documents.OpenNoRepairDialog（DocIO.openDoc），不弹「是否修复」"
        Report.line output "DisplayAlerts" (sprintf "%d（wdAlertsNone）" alerts)
        Report.line output "ConfirmConversions" (string confirm)
        Report.line output "AutomationSecurity" (sprintf "%d（ForceDisable=3）" security)
        Report.line output "恢复后读回" (sprintf "%d（Low=1，restore 生效）" restored)

        Should.equal 0 alerts
        Should.equal false confirm
        Should.equal 3 security
        Should.equal 1 restored
        Report.conclusion output "通过：宏被禁用、确认框全关，restore 能把安全级别恢复成原值"

    [<Fact>]
    member this.``安全开关：Excel 不弹修复框、不启用宏，用完恢复原值``() =
        Report.requireClean "EXCEL"

        let probe =
            ExcelHost.Host.execQuery(fun app ->
                let alerts = app.DisplayAlerts
                let security = int app.AutomationSecurity
                ExcelHost.Host.Security.restore app (Some MsoAutomationSecurity.msoAutomationSecurityLow)
                let restored = int app.AutomationSecurity
                ExcelHost.Host.Security.apply app |> ignore
                alerts, security, restored)
            |> Report.ok

        let alerts, security, restored = probe

        Report.title output "安全开关：Excel 不弹修复框、不启用宏"
        Report.line output "打开方式" "Workbooks.Open 带 CorruptLoad=xlNormalLoad（BookIO.openBook）"
        Report.line output "DisplayAlerts" (string alerts)
        Report.line output "AutomationSecurity" (sprintf "%d（ForceDisable=3）" security)
        Report.line output "恢复后读回" (sprintf "%d（Low=1，restore 生效）" restored)

        Should.equal false alerts
        Should.equal 3 security
        Should.equal 1 restored
        Report.conclusion output "通过：宏被禁用、不进修复模式，restore 能把安全级别恢复成原值"

    [<Fact>]
    member this.``安全开关：PowerPoint 不弹窗口、不启用宏，用完恢复原值``() =
        Report.requireClean "POWERPNT"

        let probe =
            PowerPointHost.Host.execQuery(fun app ->
                let alerts = app.DisplayAlerts
                let security = int app.AutomationSecurity
                PowerPointHost.Host.Security.restore app (Some MsoAutomationSecurity.msoAutomationSecurityLow)
                let restored = int app.AutomationSecurity
                PowerPointHost.Host.Security.apply app |> ignore
                alerts, security, restored)
            |> Report.ok

        let alerts, security, restored = probe

        Report.title output "安全开关：PowerPoint 不弹窗口、不启用宏"
        Report.line output "打开方式" "Presentations.Open(..., WithWindow=msoFalse)（PptIO.openPpt）"
        Report.line output "DisplayAlerts" (sprintf "%A（ppAlertsNone）" alerts)
        Report.line output "AutomationSecurity" (sprintf "%d（ForceDisable=3）" security)
        Report.line output "恢复后读回" (sprintf "%d（Low=1，restore 生效）" restored)

        Should.equal PpAlertLevel.ppAlertsNone alerts
        Should.equal 3 security
        Should.equal 1 restored
        Report.conclusion output "通过：宏被禁用、不弹窗口，restore 能把安全级别恢复成原值"

    [<Fact>]
    member this.``安全开关：KIND 只放行那一类，分三轮互不串味``() =
        let dir = Path.Combine(Path.GetTempPath(), "convert-kind")

        let sample: Item list =
            [ ".doc"; ".xls"; ".ppt" ]
            |> List.map(fun ext ->
                let src = Path.Combine(dir, "样例" + ext)
                { Source = src; Target = Paths.targetOf src })

        Report.title output "安全开关：XUNITHOST_CONVERT_KIND 按类型分三轮"
        Report.line output "解析 doc" (sprintf "%A" (Paths.parseKind "doc"))
        Report.line output "解析 XLS" (sprintf "%A" (Paths.parseKind "XLS"))
        Report.line output "解析 .ppt" (sprintf "%A" (Paths.parseKind ".ppt"))
        Report.line output "未设 / 不认得" (sprintf "%A / %A" (Paths.parseKind null) (Paths.parseKind "whatever"))

        for kind in [ Paths.Kind.Doc; Paths.Kind.Xls; Paths.Kind.Ppt; Paths.Kind.All ] do
            let picked = Paths.filterByKind kind sample
            let exts = picked |> List.map(fun i -> Path.GetExtension i.Source) |> String.concat " "
            Report.line output (Paths.kindLabel kind) (sprintf "%d 份：%s" picked.Length exts)
            Should.equal (if kind = Paths.Kind.All then 3 else 1) picked.Length

        Should.equal Paths.Kind.Doc (Paths.parseKind "doc")
        Should.equal Paths.Kind.Xls (Paths.parseKind "XLS")
        Should.equal Paths.Kind.Ppt (Paths.parseKind ".ppt")
        Should.equal Paths.Kind.All (Paths.parseKind null)
        Report.conclusion output "通过：doc / xls / ppt 各放行一类，不设 KIND 才全跑"

    [<Fact>]
    member this.``安全开关：翻倍法的 LIMIT 稳定取前 N 份，且跳过已知失败``() =
        let dir =
            Path.Combine(Path.GetTempPath(), "convert-limit-" + Guid.NewGuid().ToString("N").Substring(0, 8))

        Directory.CreateDirectory dir |> ignore
        let skipFile = Path.Combine(dir, "skip.txt")

        let item name =
            let src = Path.Combine(dir, name)
            { Source = src; Target = Paths.targetOf src }

        // 故意乱序：验证「前 N 份」是按路径排过的，不是文件系统碰巧给的顺序
        let sample: Item list =
            [ item "丙.doc"; item "甲.doc"; item "乙.doc"; item "表.xls"; item "片.ppt" ]

        let oldLimit = Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_LIMIT"
        let oldSkip = Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_SKIP"

        let names xs = xs |> List.map(fun (i: Item) -> Path.GetFileName i.Source) |> String.concat " "

        try
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_SKIP", null)
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", null)
            let all = Paths.selectForRound dir Paths.Kind.Doc sample
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", "1")
            let one = Paths.selectForRound dir Paths.Kind.Doc sample
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", "2")
            let two = Paths.selectForRound dir Paths.Kind.Doc sample
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", "99")
            let over = Paths.selectForRound dir Paths.Kind.Doc sample

            File.WriteAllText(skipFile, "# 已知失败\r\n甲.doc\r\n", System.Text.Encoding.UTF8)
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_SKIP", skipFile)
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", "1")
            let skippedOne = Paths.selectForRound dir Paths.Kind.Doc sample
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", null)
            let skippedAll = Paths.selectForRound dir Paths.Kind.Doc sample

            Report.title output "安全开关：XUNITHOST_CONVERT_LIMIT 翻倍法取份"
            Report.line output "不限" (sprintf "%d 份：%s" all.Length (names all))
            Report.line output "LIMIT=1" (sprintf "%d 份：%s" one.Length (names one))
            Report.line output "LIMIT=2" (sprintf "%d 份：%s" two.Length (names two))
            Report.line output "LIMIT=99" (sprintf "%d 份：%s" over.Length (names over))
            Report.line output "跳过甲.doc 后 LIMIT=1" (sprintf "%d 份：%s" skippedOne.Length (names skippedOne))
            Report.line output "跳过甲.doc 后 不限" (sprintf "%d 份：%s" skippedAll.Length (names skippedAll))

            Should.equal "甲.doc 乙.doc 丙.doc" (names all)
            Should.equal "甲.doc" (names one)
            Should.equal "甲.doc 乙.doc" (names two)
            Should.equal "甲.doc 乙.doc 丙.doc" (names over)
            Should.equal "乙.doc" (names skippedOne)
            Should.equal "乙.doc 丙.doc" (names skippedAll)
            Report.conclusion output "通过：排序稳定、LIMIT 只取前 N 份、已知失败不再挡住后续轮次"
        finally
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_LIMIT", oldLimit)
            Environment.SetEnvironmentVariable("XUNITHOST_CONVERT_SKIP", oldSkip)
            (try
                Directory.Delete(dir, true)
             with _ ->
                 ())
