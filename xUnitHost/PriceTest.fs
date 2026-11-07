/// 网查到的三个 PP/PPH 电动挡板门单价，回填到三张报价表「入口电动挡板门」行（第 4 行）。
///
/// 三格数字是用户点名的结果；同一行原本空的 J/K/L/M 公式照本表第 5 行的惯例补齐，
/// 否则小计 SUM(J3:J66) 不会把这三笔钱算进去。报价系数不自己定，直接沿用本节首行 M5。
/// 路径 / 取表 / 读格 / 认人 复用 DnTest.fs 的 DnReal；真表两条 Fact 靠 XUNITHOST_DN_DIR 叫起来。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel

module internal Price =

    type Item =
        { N: int
          ExpectedSpec: string
          UnitPrice: float
          Source: string }

    /// 上一轮网上查到的三个单价（元/个），与 DN400 / DN500 / DN800 一一对应。
    let items =
        [ { N = 1
            ExpectedSpec = "DN400"
            UnitPrice = 230.0
            Source = "百度爱采购·亿兆环保「PP电动风阀 管道式 阻燃耐酸碱 DN400mm」" }
          { N = 2
            ExpectedSpec = "DN500"
            UnitPrice = 500.0
            Source = "中国供应商·华涵「PP风阀 电动 公称通径500mm」1–9个价" }
          { N = 3
            ExpectedSpec = "DN800"
            UnitPrice = 1070.0
            Source = "四川轻化工大学废气处理项目中标公告·鑫宏洋 PP电动风阀 Φ800（38个）" } ]

    let note (item: Item) = sprintf "网查PP电动风阀参考单价；%s" item.Source

    let dir () = DnReal.dir ()

    /// 填价前留一份，标签与上一轮「原稿」区分开，重复跑不覆盖已有备份
    let backupPath (dir: string) (target: string) : string =
        let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmm")
        Path.Combine(dir, sprintf "3条酸洗线设备清单-1104[填价前]%s.xlsx" stamp)

    let num (sheet: Worksheet) (addr: string) : float =
        let value = sheet.Range(addr).Value2
        Assert.False(isNull value, sprintf "%s 没有值" addr)
        Convert.ToDouble(value, CultureInfo.InvariantCulture)

    /// 找「小计」行，不假定它在第几行（1# 67 / 2# 43 / 3# 66）
    let subtotalRow (sheet: Worksheet) : int =
        [ 1..sheet.UsedRange.Rows.Count ]
        |> List.tryPick(fun r -> if DnReal.cellText sheet (sprintf "A%d" r) = "小计" then Some r else None)
        |> function
            | Some r -> r
            | None -> failwith "找不到「小计」行"

    /// 只读清点：确认挡板门行还在、规格对得上、单价格还空着
    let survey (dir: string) =
        Host.execOnceQuery Kind.Excel (fun app ->
            match app with
            | ExcelApp xl ->
                xl.AskToUpdateLinks <- false
                let book = ExcelHost.BookIO.openBook xl (DnReal.quotePath dir) true

                try
                    items
                    |> List.map(fun item ->
                        let sheet = DnReal.findSheet book (DnReal.sheetName item.N)

                        {| N = item.N
                           Name = DnReal.cellText sheet "B4"
                           Spec = DnReal.cellText sheet "C4"
                           Qty = num sheet "E4"
                           Unit = DnReal.cellText sheet "I4"
                           Factor = num sheet "M5"
                           SubtotalAddr = sprintf "J%d" (subtotalRow sheet)
                           Subtotal = num sheet (sprintf "J%d" (subtotalRow sheet)) |})
                finally
                    ExcelHost.BookIO.closeBook book
            | _ -> failwith "kind 不对：这条命令只走 Excel")
        |> Report.ok

    type Row =
        { Item: Item
          Qty: float
          Factor: float
          Before: string
          AfterUnit: float
          AfterCost: float
          AfterQuoteUnit: float
          AfterQuoteTotal: float
          AfterFactor: float
          Note: string
          SubBefore: float
          SubAfter: float }

    type Outcome =
        { Rows: Row list
          Identity: string
          Backup: string
          BackedUp: bool
          SavedAt: string }

    /// 一段 Excel 命令里做完：备份 → 逐表认行 → 填单价 + 补公式 → 存盘 → 只读读回
    let write (dir: string) : Outcome =
        Host.execOnceQuery Kind.Excel (fun app ->
            match app with
            | ExcelApp xl ->
                xl.AskToUpdateLinks <- false
                let target = DnReal.quotePath dir
                Assert.True(File.Exists target, sprintf "找不到设备清单：%s" target)

                let backup = backupPath dir target
                let backedUp = not (File.Exists backup)
                if backedUp then File.Copy(target, backup, false)

                // 用户点名的现有工作簿：动笔前先认人，把身份贴进报告
                let book = ExcelHost.BookIO.openBook xl target false
                let identity = Doc.ofCom (fun () -> book.FullName) (fun () -> book.Path) |> Doc.describe

                // 先把三张表都检查一遍，任何一张不对就一个字都不写
                let prepared =
                    items
                    |> List.map(fun item ->
                        let sheet = DnReal.findSheet book (DnReal.sheetName item.N)
                        Assert.Equal("入口电动挡板门", DnReal.cellText sheet "B4")
                        Assert.Equal(item.ExpectedSpec, DnReal.cellText sheet "C4")

                        let before = DnReal.cellText sheet "I4"
                        let expected = Convert.ToString(item.UnitPrice, CultureInfo.InvariantCulture)

                        Assert.True(
                            before = "" || before = expected,
                            sprintf "%d# I4 已有别的单价「%s」，先人工确认再动" item.N before
                        )

                        let subAddr = sprintf "J%d" (subtotalRow sheet)

                        {| Item = item
                           Sheet = sheet
                           Qty = num sheet "E4"
                           Factor = num sheet "M5"
                           Before = before
                           SubAddr = subAddr
                           SubBefore = num sheet subAddr |})

                for p in prepared do
                    let sheet = p.Sheet
                    sheet.Range("I4").Value2 <- box p.Item.UnitPrice
                    sheet.Range("H4").Value2 <- note p.Item
                    // 系数沿用本节首行 M5，不自作主张定价
                    sheet.Range("M4").Value2 <- box p.Factor
                    // 与第 5 行完全同形的四个公式，只换行号；小计 SUM 本来就覆盖第 4 行
                    sheet.Range("J4").Formula <- "=I4*E4"
                    sheet.Range("K4").Formula <- "=I4*M4"
                    sheet.Range("L4").Formula <- "=K4*E4"

                xl.CalculateFull()
                book.Save()
                let savedAt = File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss")
                ExcelHost.BookIO.closeBook book

                // 存完只读重新开一遍，逐格读回来（公式格读的是算出来的值）
                let check = ExcelHost.BookIO.openBook xl target true

                let rows =
                    prepared
                    |> List.map(fun p ->
                        let sheet = DnReal.findSheet check (DnReal.sheetName p.Item.N)

                        { Item = p.Item
                          Qty = p.Qty
                          Factor = p.Factor
                          Before = p.Before
                          AfterUnit = num sheet "I4"
                          AfterCost = num sheet "J4"
                          AfterQuoteUnit = num sheet "K4"
                          AfterQuoteTotal = num sheet "L4"
                          AfterFactor = num sheet "M4"
                          Note = DnReal.cellText sheet "H4"
                          SubBefore = p.SubBefore
                          SubAfter = num sheet p.SubAddr })

                ExcelHost.BookIO.closeBook check

                { Rows = rows
                  Identity = identity
                  Backup = backup
                  BackedUp = backedUp
                  SavedAt = savedAt }
            | _ -> failwith "kind 不对：这条命令只走 Excel")
        |> Report.ok

/// 真表两条：先只读清点，再写盘
type PriceFolderTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``真表：清点三张报价表的挡板门行（一个文件都不动）``() =
        match Price.dir () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_DN_DIR=<目录> 才清点")
        | Some dir ->
            Report.requireClean "EXCEL"
            Assert.True(Directory.Exists dir, sprintf "目录不存在：%s" dir)
            let snap = Price.survey dir

            Report.title output (sprintf "清点 %s 的挡板门行（只读，一个文件都不动）" dir)

            for s in snap do
                let item = Price.items |> List.find(fun i -> i.N = s.N)

                Report.line2
                    output
                    (sprintf "%d#" s.N)
                    "B4/C4/E4"
                    (sprintf "「%s」/ %s / %.0f个　I4 单价=「%s」" s.Name s.Spec s.Qty s.Unit)

                Report.line2
                    output
                    (sprintf "%d#" s.N)
                    "待填"
                    (sprintf "¥%.0f/个　报价系数沿用 M5=%.2f　%s 小计=¥%.0f" item.UnitPrice s.Factor s.SubtotalAddr s.Subtotal)

            Should.equal [ "DN400"; "DN500"; "DN800" ] (snap |> List.map(fun s -> s.Spec))
            Should.equal [ ""; ""; "" ] (snap |> List.map(fun s -> s.Unit))
            Report.conclusion output "只清点：三张表挡板门行都在、规格与 DN 一致、单价格还空着，一个字没动"

    [<Fact>]
    member this.``真表：回填三张报价表的入口电动挡板门网查单价``() =
        match Price.dir () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_DN_DIR=<目录> 才回填")
        | Some dir ->
            Report.requireClean "EXCEL"
            Assert.True(Directory.Exists dir, sprintf "目录不存在：%s" dir)

            let result = Price.write dir
            Report.title output "真表：网查三个电动门单价 → 报价表「入口电动挡板门」"
            Report.line output "认人" result.Identity

            Report.line
                output
                "备份"
                (if result.BackedUp then result.Backup else "已有备份，未重复生成：" + result.Backup)

            for row in result.Rows do
                Report.line2
                    output
                    (sprintf "%d#" row.Item.N)
                    "单价"
                    (sprintf
                        "%s　¥%.0f/个 × %.0f个　来源：%s"
                        row.Item.ExpectedSpec
                        row.Item.UnitPrice
                        row.Qty
                        row.Item.Source)

                Report.line2
                    output
                    (sprintf "%d#" row.Item.N)
                    "读回"
                    (sprintf
                        "原 I4=「%s」→ I4=¥%.0f　成本总价 J4=¥%.0f　报价单价 K4=¥%.2f　报价总价 L4=¥%.0f（系数 %.2f）"
                        row.Before
                        row.AfterUnit
                        row.AfterCost
                        row.AfterQuoteUnit
                        row.AfterQuoteTotal
                        row.AfterFactor)

                Report.line2
                    output
                    (sprintf "%d#" row.Item.N)
                    "小计"
                    (sprintf "¥%.0f → ¥%.0f（+%.0f，含 15%% 取费行联动）" row.SubBefore row.SubAfter (row.SubAfter - row.SubBefore))

                Should.equal row.Item.UnitPrice row.AfterUnit
                Should.equal (row.Item.UnitPrice * row.Qty) row.AfterCost
                Should.equal (row.Item.UnitPrice * row.Factor) row.AfterQuoteUnit
                Should.equal (row.Item.UnitPrice * row.Factor * row.Qty) row.AfterQuoteTotal
                Assert.Equal(row.Factor, row.AfterFactor)
                Assert.Contains("网查PP电动风阀参考单价", row.Note)
                Assert.True(row.SubAfter > row.SubBefore)

            Report.line output "保存" result.SavedAt
            Report.conclusion output "通过：三个单价写进去又读回来一致，成本/报价总价与成本小计都联动上了，填价前已备份"
