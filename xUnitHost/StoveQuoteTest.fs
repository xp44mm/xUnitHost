/// 在用户正开着的「热风炉报价表.xlsx」里生成中冶南方都市环保项目的管道内烟气升温燃烧系统报价单。
///
/// 结构与数量照抄旭日厂「500万大卡高炉煤气燃烧器报价单.xlsx」（济南志晟格式，本单品牌改为「群有」）：
/// 6 个分类、30 条明细（第 3~36 行）、合计 37 行、合同优惠价 38 行。
/// 只改四处：① 燃烧器按本项目 920.93 万大卡/h（10.71 MW）重新选型，型号前缀取群有拼音首字母 QY → QYL-1000；
/// ② 燃烧器本体下方增列「燃烧室本体 φ1800×1200」（原旭日单不含炉膛，本项目按图单列）；
/// ③ 助燃风机由 5500m³/h·3kPa·7.5kW 改为 12000m³/h·4.5kPa·30kW；
/// ④ 煤气阀组 DN300 → DN250（本项目煤气支管内径 250mm，13.98 m/s）；
/// ⑤ 助燃风侧按总管内径 550mm 放大：风门调节阀 DN300 → DN500（12000m³/h 过 DN300 是 47m/s，过 DN500 为 17m/s）；
/// ⑥ 燃烧器外管内径 600mm（原摘要 φ518），规格串按计算书新值写。
///
/// 单价依据（本轮 C:\崔胜利\群有 全树检索 + 网上比对，详见报告行）：
///   燃烧器 7.5 万（旭日 ZSL-450/500万大卡）→ 12.0 万（按 (920.93/500)^0.6 折算 11.0，
///       另按祥云 35MW 煤粉燃烧器 22 万反推 11.2，取 12.0 含套管加大）
///   燃烧室本体 4.8 万：外壳 Q235B δ8 约 535kg（1.12）+ 耐材 0.98t（0.33）+ 施工烘炉 0.40
///       + 310S 锚固 0.18 + 保温 0.13 + 附件 0.30 = 成本 2.46 万，毛利 45~55% → 4.5~5.5，取 4.8；
///       交叉验证：群有湖州酸洗线混风箱 3.23/6.40 万两档，本项目 φ1800×1200 带耐材落在中间
///   助燃风机 0.7 万（旭日 5500/3kPa/7.5kW）、1.4 万（旭日 8400/3kPa/11kW）→ 2.6 万；
///       网上 30kW 4-68 不锈钢 1.7 万、碳钢送风机 1.2~1.5 万，取 2.6 含防爆电机
///   阀组 DN250：网价硬密封电动蝶阀 0.13~0.485 万/台（阀体），英菲百控成套 DN300 为 1.5 万，
///       DN250 取 1.3~1.35 万
///   控制/火检/点火件与尺寸无关，沿用旭日成交价
///
/// 闸门：XUNITHOST_EXCEL_PID=<pid>，没设就 Skip；只挂回、不新建、不 Quit、不关用户的书。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal Quote =

    /// 用户点名的这本（ExcelVSTO 黑板：PID 7720 会话里的工作簿）
    let target = @"C:\崔胜利\群有\中冶南方都市环保\热风炉报价表.xlsx"

    let title = "管道内烟气升温燃烧系统报价单-中冶南方都市环保（焦炉煤气）"

    /// 合同优惠价：合计 33.63 让到 33.1（-1.6%，让价 0.53 万与上一版一致），口径与旭日单一致
    let discount = 33.1

    /// 一行：Price = None 表示分类标题行（只有 A/B 两格）
    type Line =
        { Row: int
          No: string
          Name: string
          Spec: string
          Unit: string
          Qty: float
          Price: float option
          Brand: string }

    let private item row no name spec unit qty price brand =
        { Row = row
          No = no
          Name = name
          Spec = spec
          Unit = unit
          Qty = qty
          Price = Some price
          Brand = brand }

    let private head row no name =
        { Row = row
          No = no
          Name = name
          Spec = ""
          Unit = ""
          Qty = 0.0
          Price = None
          Brand = "" }

    let lines =
        [ head 3 "一" "燃烧器系统"
          item
              4
              "1"
              "燃烧器（套管式）"
              "QYL-1000；额定 920.93 万大卡/h（10.71MW）；煤气 2302Nm³/h@4kPa（支管 φ250，13.98m/s）；助燃风 12000m³/h（总管 φ550，13.92m/s）；燃烧器外管 φ600；310S/304/Q235"
              "套"
              1.0
              12.0
              "群有"
          item
              5
              "2"
              "燃烧室本体"
              "φ1800×1200；外壳 Q235B δ8；耐材内衬 δ120（轻质莫来石，高温段刚玉）；310S 锚固钉；含看火孔/测温座/防爆门/外保温"
              "套"
              1.0
              4.8
              "群有"
          item 6 "3" "防爆主风机" "Q=12000m³/h；P=4.5kPa；30kW" "套" 1.0 2.6 "国产防爆"
          item 7 "4" "火焰检测器" "C7027A" "套" 2.0 0.3 "霍尼韦尔"
          item 8 "5" "风门调节阀" "DN500；防爆" "套" 1.0 2.0 "英菲百控"
          item 9 "6" "火焰放大器" "BC1000" "套" 1.0 0.2 "霍尼韦尔"
          head 10 "二" "点火装置"
          item 11 "1" "点火枪" "" "套" 1.0 0.2 "群有"
          item 12 "2" "点火变压器" "2×7kV" "套" 1.0 0.2 "意大利飞达"
          item 13 "3" "高压点火电极" "" "套" 1.0 0.1 "意大利飞达"
          item 14 "4" "点火电磁阀" "DN25" "套" 1.0 0.2 "德国GUNGS"
          item 15 "5" "火种火焰检测" "" "套" 1.0 0.1 "西门子"
          item 16 "6" "火焰放大器" "BC1000" "套" 1.0 0.3 "霍尼韦尔"
          head 17 "三" "控制系统"
          item 18 "1" "控制柜其中控制" "" "套" 1.0 1.2 "群有"
          item 19 "2" "PLC（含拓展模块）" "S7-200  Smart" "套" 1.0 0.6 "西门子"
          item 20 "3" "触摸屏" "10寸中文彩色" "套" 1.0 0.3 "西门子"
          item 21 "4" "防爆电控柜" "1200*800*600" "套" 1.0 0.7 "定制"
          item 22 "5" "其他附属元件" "" "套" 1.0 0.3 "施耐德"
          head 23 "四" "焦炉煤气阀组"
          item 24 "1" "手动金属碟阀" "D71H-10C DN250" "套" 1.0 0.35 "英菲百控"
          item 25 "2" "防爆气动切断蝶阀" "D671H-10C DN250" "套" 2.0 1.3 "英菲百控"
          item 26 "3" "防爆电动调节阀" "D941-10C DN250" "套" 1.0 1.35 "英菲百控"
          item 27 "4" "气动放散阀" "Q641H-10C DN40" "套" 1.0 0.2 "英菲百控"
          item 28 "5" "手动放散阀" "Q41W-1C DN40" "套" 1.0 0.2 "英菲百控"
          item 29 "6" "检漏压力开关" "GW50.." "套" 1.0 0.2 "德国GUNGS"
          item 30 "7" "空气压力开关" "GW50.." "套" 1.0 0.2 "德国GUNGS"
          item 31 "8" "燃气压力表" "0-25kpa" "套" 1.0 0.03 "上海神睫"
          item 32 "9" "补偿器" "TB250-4" "套" 1.0 0.25 "国产"
          item 33 "10" "煤气外置捡漏装置" "QB2000T/N" "套" 1.0 0.2 "河南驰诚电气"
          item 34 "11" "温度变送器" "WRN-230 0~1300℃" "套" 1.0 0.15 "上海神睫"
          item 35 "五" "点火调试" "" "套" 1.0 1.0 "我方负责"
          item 36 "六" "包装及运输" "" "套" 1.0 0.5 "我方负责" ]

    /// 明细合计（万元），浮点用容差比
    let total =
        lines
        |> List.sumBy (fun l ->
            match l.Price with
            | Some p -> p * l.Qty
            | None -> 0.0)

    /// 分类小计：按行号区间取，别按「有没有下一个分类标题」推
    /// ——「五 点火调试」「六 包装及运输」是单行分类（A 列序号 + B 列名称 + 带单价），
    /// 它们本身有价、后面也没有下一个标题，用 takeWhile 会一路吃到表尾
    let sections =
        [ "燃烧器系统", 3, 9
          "点火装置", 10, 16
          "控制系统", 17, 22
          "焦炉煤气阀组", 23, 34
          "点火调试", 35, 35
          "包装及运输", 36, 36 ]

    let subtotal (name: string) =
        let _, first, last = sections |> List.find (fun (n, _, _) -> n = name)

        lines
        |> List.filter (fun l -> l.Row >= first && l.Row <= last && l.Price.IsSome)
        |> List.sumBy (fun l -> l.Price.Value * l.Qty)

    let num (sheet: Worksheet) (addr: string) : float =
        let value = sheet.Range(addr).Value2
        if isNull value then nan else Convert.ToDouble(value, CultureInfo.InvariantCulture)

    let text (sheet: Worksheet) (addr: string) : string =
        let value = sheet.Range(addr).Text
        if isNull value then "" else Convert.ToString(value, CultureInfo.InvariantCulture)

    let pid () =
        match Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID" with
        | null
        | "" -> None
        | s ->
            match Int32.TryParse s with
            | true, p -> Some p
            | _ -> None

    /// 生成前留一份原稿；同名备份只留最早那份，重复跑不再叠新文件
    let backupPath () : string =
        let dir = Path.GetDirectoryName target

        let existing =
            Directory.GetFiles(dir, "热风炉报价表[生成前]*.xlsx")

        match existing with
        | [||] ->
            let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmm")
            Path.Combine(dir, sprintf "热风炉报价表[生成前]%s.xlsx" stamp)
        | xs -> xs |> Array.sortBy File.GetLastWriteTime |> Array.head

    let findBook (app: Application) : Workbook option =
        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
        |> List.map (fun i -> app.Workbooks.Item(i))
        |> List.tryFind (fun b ->
            String.Equals((try b.FullName with _ -> ""), target, StringComparison.OrdinalIgnoreCase))

    /// 只读清点：书在不在会话里、这份表是不是还空着、来源单（旭日）能不能读到
    type Survey =
        { Books: string list
          SheetName: string
          A1: string
          UsedRange: string
          SourceTotal: float
          SourceDiscount: float }

    let survey () : Survey =
        let pid = pid () |> Option.get
        let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await

        STA.run (fun () ->
            let books =
                [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                |> List.map (fun i -> try app.Workbooks.Item(i).Name with _ -> "?")

            match findBook app with
            | None -> failwithf "眼前这只 Excel（PID %d）里没有打开 %s" pid target
            | Some book ->
                let sheet = book.Worksheets.[1] :?> Worksheet

                // 来源单：旭日厂那份，只读它的合计与优惠价做口径对照
                let src =
                    @"C:\崔胜利\群有\旭日厂\500万大卡高炉煤气燃烧器报价单.xlsx"

                let srcTotal, srcDiscount =
                    if not (File.Exists src) then
                        nan, nan
                    else
                        let s = ExcelHost.BookIO.openBook app src true
                        let sh = s.Worksheets.[1] :?> Worksheet
                        let t = num sh "G36"
                        let d = num sh "G37"
                        ExcelHost.BookIO.closeBook s
                        t, d

                { Books = books
                  SheetName = sheet.Name
                  A1 = text sheet "A1"
                  UsedRange = sheet.UsedRange.Address(false, false)
                  SourceTotal = srcTotal
                  SourceDiscount = srcDiscount })

    type Result =
        { Backup: string
          BackedUp: bool
          SavedAt: string
          Total: float
          Discount: float
          Rows: (string * string * string) list }

    /// 一段：备份 → 写 A:H → 公式 → 保存 → 只读读回
    let write () : Result =
        match pid () with
        | None -> failwith "没设 XUNITHOST_EXCEL_PID"
        | Some pid ->
            let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await

            let backup = backupPath ()
            let backedUp = not (File.Exists backup)
            if backedUp then File.Copy(target, backup, false)

            STA.run (fun () ->
                let book = findBook app |> Option.defaultWith (fun () -> failwith "书不在了")

                let oldAlerts = app.DisplayAlerts
                app.DisplayAlerts <- false
                try
                    let sheet = book.Worksheets.[1] :?> Worksheet

                    // 空表才写：A1 有字说明用户已经动过，停手
                    let a1 = text sheet "A1"

                    Assert.True(
                        a1 = "" || a1 = title,
                        sprintf "A1 已有别的内容「%s」，先人工确认" a1
                    )

                    sheet.Range("A1:H1").Merge()
                    sheet.Range("A1").Value2 <- box title

                    // 明细区整块重写（A3:H40 全是本表自生成的内容，先清掉避免旧行残留）
                    sheet.Range("A3:H40").ClearContents() |> ignore

                    [ "序号"; "名称 "; "型号规格或材质"; "单位"; "数量"; "单价\n（万元）"; "总价\n（万元）"; "配套商" ]
                    |> List.iteri (fun i h -> sheet.Range(sprintf "%s2" (string (char (65 + i)))).Value2 <- box h)

                    for l in lines do
                        sheet.Range(sprintf "A%d" l.Row).Value2 <- box l.No
                        sheet.Range(sprintf "B%d" l.Row).Value2 <- box l.Name

                        match l.Price with
                        | Some p ->
                            if l.Spec <> "" then
                                sheet.Range(sprintf "C%d" l.Row).Value2 <- box l.Spec

                            sheet.Range(sprintf "D%d" l.Row).Value2 <- box l.Unit
                            sheet.Range(sprintf "E%d" l.Row).Value2 <- box l.Qty
                            sheet.Range(sprintf "F%d" l.Row).Value2 <- box p
                            sheet.Range(sprintf "G%d" l.Row).Formula <- sprintf "=E%d*F%d" l.Row l.Row
                            sheet.Range(sprintf "H%d" l.Row).Value2 <- box l.Brand
                        | None -> ()

                    sheet.Range("B37").Value2 <- box "合计"
                    sheet.Range("G37").Formula <- "=SUM(G3:G36)"
                    sheet.Range("B38").Value2 <- box "合同优惠价"
                    sheet.Range("G38").Value2 <- box discount

                    // 只做可读性：列宽、标题行高、表头加粗、金额两位小数
                    [ "A", 5.0; "B", 18.0; "C", 46.0; "D", 6.0; "E", 6.0; "F", 10.0; "G", 10.0; "H", 13.0 ]
                    |> List.iter (fun (c, w) -> (sheet.Columns.[c] :?> Range).ColumnWidth <- w)

                    (sheet.Rows.[1] :?> Range).RowHeight <- 26.0
                    sheet.Range("A1").Font.Bold <- true
                    sheet.Range("A1").Font.Size <- 14.0
                    sheet.Range("A2:H2").Font.Bold <- true
                    sheet.Range("F4:G38").NumberFormat <- "0.00"

                    app.CalculateFull()
                    book.Save()
                    let savedAt = File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss")

                    // 读回：逐行 A/B/E/F/G/H 的真实显示值
                    let rows =
                        lines
                        |> List.map (fun l ->
                            (sprintf "%d" l.Row,
                             sprintf "%s %s" (text sheet (sprintf "A%d" l.Row)) (text sheet (sprintf "B%d" l.Row)),
                             if l.Price.IsSome then
                                 sprintf "×%s　单价 %s　总价 %s"
                                     (text sheet (sprintf "E%d" l.Row))
                                     (text sheet (sprintf "F%d" l.Row))
                                     (text sheet (sprintf "G%d" l.Row))
                             else ""))

                    { Backup = backup
                      BackedUp = backedUp
                      SavedAt = savedAt
                      Total = num sheet "G37"
                      Discount = num sheet "G38"
                      Rows = rows }
                finally
                    app.DisplayAlerts <- oldAlerts)

/// 真机两条：先只读清点，再生成
type StoveQuoteTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``热风炉报价：清点空表与旭日来源单（一个字都不写）``() =
        match Quote.pid () with
        | None -> Assert.Skip("不硬编码 PID：设 XUNITHOST_EXCEL_PID=<pid> 才清点")
        | Some pid ->
            let snap = Quote.survey ()

            Report.title output (sprintf "清点 PID %d 的 %s（只读）" pid Quote.target)
            Report.line output "会话里的书" (String.concat " ｜ " snap.Books)
            Report.line2 output "目标表" "工作表/已用区" (sprintf "%s ｜ %s" snap.SheetName snap.UsedRange)
            Report.line2 output "目标表" "A1" (sprintf "「%s」（空表才能写）" snap.A1)
            Report.line2 output "来源单" "旭日 500万大卡" (sprintf "合计 %.2f 万　优惠价 %.2f 万" snap.SourceTotal snap.SourceDiscount)
            Report.line2 output "本项目" "清单/合计" (sprintf "%d 行（含 6 个分类）　合计 %.2f 万　优惠价 %.2f 万" Quote.lines.Length Quote.total Quote.discount)

            Should.equal true (snap.A1 = "" || snap.A1 = Quote.title)
            Report.conclusion output "只清点：目标表为空表或本工具生成的表，旭日来源单可读，一个字没动"

    [<Fact>]
    member this.``热风炉报价：生成本项目管道内烟气升温燃烧系统报价单``() =
        match Quote.pid () with
        | None -> Assert.Skip("不硬编码 PID：设 XUNITHOST_EXCEL_PID=<pid> 才生成")
        | Some pid ->
            Report.title output (sprintf "在 PID %d 的「热风炉报价表.xlsx」里生成报价单" pid)
            Report.line output "口径" "结构照旭日单（30 条明细/6 分类，较旭日多一条燃烧室本体），燃烧器 920.93 万大卡、外管 φ600、风机 12000m³/h、煤气管 φ250→阀组 DN250、助燃风总管 φ550→风门阀 DN500"

            let r = Quote.write ()

            Report.line
                output
                "备份"
                (if r.BackedUp then r.Backup else "已有备份，未重复生成：" + r.Backup)

            for (row, label, detail) in r.Rows do
                Report.line2 output (sprintf "%2s 行" row) label detail

            for name in [ "燃烧器系统"; "点火装置"; "控制系统"; "焦炉煤气阀组"; "点火调试"; "包装及运输" ] do
                Report.line2 output "分类小计" name (sprintf "%.2f 万" (Quote.subtotal name))

            Report.line2 output "合计" "G37=SUM(G3:G36)" (sprintf "%.2f 万（算出来 %.2f）" r.Total Quote.total)
            Report.line2 output "合同优惠价" "G38" (sprintf "%.2f 万" r.Discount)
            Report.line output "保存" r.SavedAt

            Assert.True(abs (r.Total - Quote.total) < 1e-6, sprintf "合计对不上：%.6f vs %.6f" r.Total Quote.total)
            Should.equal Quote.discount r.Discount
            Report.conclusion output "通过：30 条明细单价与总价都写进用户正开着的那本，合计与分项小计自洽，生成前已备份"
