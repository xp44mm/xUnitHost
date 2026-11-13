/// 在用户正开着的 Excel（ExcelVSTO 黑板 PID）里，把「不同类型的燃烧器选型计算」一次性打开。
///
/// 范围：C:\崔胜利\群有 全树（跳过 . 开头目录与 ~$ 临时文件），实测 116 个 xls/xlsx/xlsm，
/// 逐个看工作表名之后按「燃料类型 / 计算口径」分类，只挑真正做选型计算的六份：
///
///   煤气          煤气燃烧计算.xlsx                  大气与烟气参数 / 煤气 / 过剩空气 / 烟气分析
///   天然气        天然气燃烧计算.xlsx                同上四表，燃料换天然气
///   燃煤          煤燃烧计算.xlsx                    同上四表，煤质表
///   天然气(吸热)  河北浩锐陶瓷\天然气燃烧计算.xlsx   体积法scr烟气吸热 / 天然气自身吸热 / 耗量 / 燃烧器
///   烟气再热      旭日厂\flue reheat.xlsx            烟气公用 / 煤气 / 天然气（双燃料）
///   功率+选型     广西盛隆\燃烧器功率计算.xlsx       烟气公用 / 煤气 / 设备选型
///
/// 不打开：`天然气燃烧计算-旧式数组`/`.CSE备份`/`煤燃烧计算（崔胜利）` 是同模板旧副本；
/// `500万大卡…报价单`/`燃烧器重量估算`/`ZL-…配置一览表`/`招标书热风炉系统摘要` 是报价、重量、清单，不做选型计算。
///
/// 只打开、不 Save、不关任何书（含用户原来那本）；打开期间临时关掉 DisplayAlerts 与
/// AskToUpdateLinks，结束后原值还原。全程不新建 Excel、不 Quit。
///
///   XUNITHOST_EXCEL_PID=11468 XUNITHOST_OPEN_BURNERS=1 \
///     xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*燃烧器选型*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal BurnerList =

    let root = @"C:\崔胜利\群有"

    /// 燃料类型 / 计算口径 —— 打开后原样打回报告，方便对照
    type Item =
        { Kind: string
          Path: string
          /// 同名冲突时的打开路径：Excel 同一会话里起不出两本同名工作簿
          /// （COM 打开第二个同名文件会静默失败、集合里根本没有它），
          /// 这条先把源复制成带来源标记的副本再开
          Alias: string option
          Why: string }

    let aliasRoot = Path.Combine(Path.GetTempPath(), "xUnitHost-burner-templates")

    /// 逐本记账：这本是什么口径、是本来就开着还是这次新开、里面有哪些表
    type Opened =
        { Kind: string
          State: string
          Path: string
          Sheets: string }

    /// 一次开完的现场：开之前有几本、开之后有几本、应用开关是否还原
    type Catalog =
        { Rows: Opened list
          Before: string list
          After: string list
          AlertsOn: bool
          Active: string }

    let items =
        [ { Kind = "煤气"
            Path = Path.Combine(root, @"煤气燃烧计算.xlsx")
            Alias = None
            Why = "焦炉/混合煤气，四表结构（参数→煤气→过剩空气→烟气分析）" }
          { Kind = "天然气"
            Path = Path.Combine(root, @"天然气燃烧计算.xlsx")
            Alias = None
            Why = "同构模板换燃料：验证四种燃料共用一套计算骨架" }
          { Kind = "燃煤"
            Path = Path.Combine(root, @"煤燃烧计算.xlsx")
            Alias = None
            Why = "煤质表口径：燃料从气体换成固体，产物算法不同" }
          { Kind = "天然气(吸热+耗量)"
            Path = Path.Combine(root, @"河北浩锐陶瓷\天然气燃烧计算.xlsx")
            Alias = Some(Path.Combine(aliasRoot, "天然气燃烧计算-河北浩锐陶瓷-吸热耗量口径.xlsx"))
            Why = "另一套口径：体积法烟气吸热 → 自身吸热 → 耗量 → 燃烧器" }
          { Kind = "烟气再热(双燃料)"
            Path = Path.Combine(root, @"旭日厂\flue reheat.xlsx")
            Alias = None
            Why = "烟气公用 + 煤气/天然气两张，同一个再热负荷下比两种燃料" }
          { Kind = "功率+设备选型"
            Path = Path.Combine(root, @"广西盛隆\燃烧器功率计算.xlsx")
            Alias = None
            Why = "唯一带「设备选型」表的：功率 → 设备规格" } ]

    /// 同名文件靠上一层目录区分（根目录与河北浩锐陶瓷各有一本「天然气燃烧计算.xlsx」）
    let shortPath (p: string) =
        let dir = Path.GetFileName(Path.GetDirectoryName p)
        if String.IsNullOrEmpty dir then Path.GetFileName p else sprintf "%s\\%s" dir (Path.GetFileName p)

    /// 明确排除的（报告里列出来，省得再问）
    let skipped =
        [ @"天然气燃烧计算-旧式数组.xlsx", "同模板旧数组写法（Range.Formula 时代）"
          @"天然气燃烧计算.CSE备份.xlsx", "同模板备份"
          @"煤燃烧计算（崔胜利）.xlsx", "同模板个人副本"
          @"旭日厂\500万大卡高炉煤气燃烧器报价单.xlsx", "报价单，非选型计算"
          @"旭日厂\500万大卡天然气燃烧器报价单.xlsx", "报价单，非选型计算"
          @"旭日厂\燃烧器重量估算.xlsx", "重量估算，非热力选型"
          @"旭日厂\招标书热风炉系统摘要.xlsx", "招标摘要"
          @"湖州酸洗线脱硝\孟1103\ZL-900配置一览表-燃烧系统.xlsx", "设备配置清单"
          @"湖州酸洗线脱硝\孟1103\ZL-300~350配置一览表-燃烧系统.xlsx", "设备配置清单"
          @"旭日厂\后英设计计算-0911.xlsx", "1.2MB SCR 全专业设计书，燃烧器只占两页（煤气量/天然气量）" ]

type BurnerCatalogTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``燃烧器选型 不同燃料的计算模板一起挂进当前 Excel``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_OPEN_BURNERS" with
        | null
        | "" ->
            Assert.Skip "设 XUNITHOST_OPEN_BURNERS=1 才真的往用户 Excel 里开文件（只开不关、不 Save）"
        | _ ->
            let pidText = Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID"

            if String.IsNullOrWhiteSpace pidText then
                Assert.Skip "没设 XUNITHOST_EXCEL_PID：这条要挂在活着的 Excel 上"

            let pid = int pidText

            Report.title output "燃烧器选型　按燃料类型打开不同口径的计算模板（只开不关）"
            Report.line output "范围" BurnerList.root
            Report.line output "挂回" (sprintf "PID %d（只附着，不新建、不 Quit）" pid)

            // 清单里哪个文件不在就早报，别等 Excel 抛
            let missing =
                BurnerList.items
                |> List.filter (fun i -> not (File.Exists i.Path))
                |> List.map (fun i -> i.Path)

            if not missing.IsEmpty then
                failwithf "清单里有文件不在：%s" (String.concat "；" missing)

            // 同名工作簿在同一会话里起不出第二本：需要对照的那份先复制成带来源标记的副本
            BurnerList.items
            |> List.iter (fun item ->
                match item.Alias with
                | Some alias ->
                    let dir = Path.GetDirectoryName alias

                    if not (Directory.Exists dir) then
                        Directory.CreateDirectory dir |> ignore

                    let stale =
                        not (File.Exists alias)
                        || File.GetLastWriteTime item.Path > File.GetLastWriteTime alias

                    if stale then
                        File.Copy(item.Path, alias, true)

                    Report.line output "同名副本" (sprintf "%s → %s" (BurnerList.shortPath item.Path) alias)
                | None -> ())

            let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await
            Should.equal (Some pid) (STA.run (fun () -> ExcelHost.Host.pidOf app))

            let result =
                STA.run (fun () ->
                    let before =
                        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                        |> List.map (fun i -> try app.Workbooks.Item(i).FullName with _ -> "?")

                    let oldAlerts = app.DisplayAlerts
                    let oldLinks = try app.AskToUpdateLinks with _ -> true

                    // 打开期间别弹「正在更新链接」「是否启用编辑」卡住
                    app.DisplayAlerts <- false
                    (try
                        app.AskToUpdateLinks <- false
                     with _ ->
                         ())

                    // 按全名取书：同名不同路径的两本（根目录与河北浩锐陶瓷各一本「天然气燃烧计算.xlsx」）
                    // 只能靠 FullName 区分，Workbooks.Open 的返回值在这里不可靠
                    let findBook (path: string) =
                        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                        |> List.map (fun i -> app.Workbooks.Item(i))
                        |> List.tryFind (fun b ->
                            let full = try b.FullName with _ -> ""
                            String.Equals(full, path, StringComparison.OrdinalIgnoreCase))

                    let rows =
                        try
                            BurnerList.items
                            |> List.map (fun item ->
                                // 同名冲突那条开的是副本，其余就是原路径
                                let actual = item.Alias |> Option.defaultValue item.Path

                                let existed =
                                    before
                                    |> List.exists (fun p -> String.Equals(p, actual, StringComparison.OrdinalIgnoreCase))

                                let state, sheets =
                                    try
                                        let book =
                                            if existed then
                                                // 已经开着就别再开一份：直接取回那一本
                                                match findBook actual with
                                                | Some b -> b
                                                | None -> failwith "记着已打开，但集合里按全名找不到"
                                            else
                                                // UpdateLinks=0：别去更新外部链接（带链接的书会因此卡在提示框）
                                                let _ = app.Workbooks.Open(actual, 0)

                                                match findBook actual with
                                                | Some b -> b
                                                | None ->
                                                    let now =
                                                        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                                                        |> List.map (fun i -> try app.Workbooks.Item(i).Name with _ -> "?")
                                                        |> String.concat "、"

                                                    failwithf "Open 之后按全名仍取不到这一本；当前集合：%s" now

                                        (if existed then "本来已打开" else "新打开"),
                                        ([ 1 .. (try book.Worksheets.Count with _ -> 0) ]
                                         |> List.map (fun i ->
                                             let s = book.Worksheets.Item(i) :?> Worksheet
                                             s.Name)
                                         |> String.concat " / ")
                                    with ex ->
                                        sprintf "失败：%s" ex.Message, ""

                                let opened: BurnerList.Opened =
                                    { Kind = item.Kind
                                      State = state
                                      Path = actual
                                      Sheets = sheets }

                                opened)
                        with ex ->
                            // 中途失败也要把开关还原，别把用户 Excel 留在「不提示」状态
                            (try
                                app.DisplayAlerts <- oldAlerts
                                app.AskToUpdateLinks <- oldLinks
                             with _ ->
                                 ())

                            reraise ()

                    (try
                        app.DisplayAlerts <- oldAlerts
                        app.AskToUpdateLinks <- oldLinks
                     with _ ->
                         ())

                    let after =
                        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                        |> List.map (fun i -> try app.Workbooks.Item(i).FullName with _ -> "?")

                    let catalog: BurnerList.Catalog =
                        { Rows = rows
                          Before = before
                          After = after
                          AlertsOn = app.DisplayAlerts
                          Active = (try app.ActiveWorkbook.Name with _ -> "?") }

                    catalog)

            Report.line output "打开前" (sprintf "%d 本" result.Before.Length)

            result.Before
            |> List.iter (fun p -> Report.line output "  ·" (BurnerList.shortPath p))

            Report.line output "本次清单" ""

            result.Rows
            |> List.iter (fun row ->
                Report.line output (sprintf "  %s ｜%s" row.Kind row.State) (BurnerList.shortPath row.Path)
                Report.line output "" (sprintf "      %s" row.Sheets))

            Report.line
                output
                "打开后"
                (sprintf "%d 本（原来 %d 本，新开 %d 本，一本没关）" result.After.Length result.Before.Length (result.After.Length - result.Before.Length))

            result.After |> List.iter (fun p -> Report.line output "  ·" (BurnerList.shortPath p))

            Report.line output "活动工作簿" result.Active

            // 这些模板的核心公式都是 XLL 函数：先探一个，别让人对着满屏 #NAME? 找原因
            let probe =
                STA.run (fun () ->
                    try
                        sprintf "%A" (app.Evaluate(box "molar(\"CO2\")"))
                    with ex ->
                        sprintf "异常：%s" ex.Message)

            Report.line output "XLL 探测" (sprintf "molar(\"CO2\") = %s" probe)
            Report.line output "DisplayAlerts 已还原" (sprintf "%b" result.AlertsOn)
            Report.line output "明确没打开" ""

            BurnerList.skipped
            |> List.iter (fun (f, why) -> Report.line output "  ·" (sprintf "%s —— %s" f why))

            // 一本没关：原来的都在，且只多出「新打开」那几本
            let fresh = result.Rows |> List.filter (fun row -> row.State = "新打开") |> List.length

            let failed = result.Rows |> List.filter (fun row -> row.State.StartsWith "失败") |> List.length

            Should.equal true result.AlertsOn
            Should.equal 0 failed
            Should.equal (result.Before.Length + fresh) result.After.Length

            result.Before
            |> List.iter (fun p ->
                let kept =
                    result.After
                    |> List.exists (fun q -> String.Equals(q, p, StringComparison.OrdinalIgnoreCase))

                Should.equal true kept)

            Report.conclusion
                output
                (sprintf "通过：%d 份不同口径的燃烧器选型计算全在进程里（新开 %d 本），原有工作簿一本没关" result.Rows.Length fresh)
