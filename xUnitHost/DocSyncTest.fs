/// 把设备清单里新增的「入口电动挡板门」同步进技术方案 docx 的三张设备清单表（表6.1/6.2/6.3）。
///
/// 用户拍板：只补这一行，其余 60+ 行一字不动；主文件原地更新，改前先另存 [原稿] 备份。
/// 因此不套 Doc.ensureFresh（那是给「还没落盘的新文档」的保护，这里用户点名授权改磁盘上的现有文件），
/// 但仍用 Doc.describe 把身份贴进报告。
/// 真表两条 Fact 沿用 XUNITHOST_DN_DIR（与 DnTest / PriceTest 同一个目录），不设就跳过。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Word

module internal DocSync =

    let dir () = DnReal.dir ()

    let targetPath (dir: string) = Path.Combine(dir, "湖州3条酸洗线脱硝技术方案-20261104.docx")

    /// 备份名带原稿的落盘时刻，重复跑不会覆盖已有备份
    let backupPath (dir: string) (target: string) : string =
        let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmm")
        Path.Combine(dir, sprintf "湖州3条酸洗线脱硝技术方案-20261104[原稿]%s.docx" stamp)

    let gate = "入口电动挡板门"

    /// Word 的互操作参数一律 obj&（按引用），先备好 mutable 再取地址
    let tablesOf (doc: Document) : Table list = [ for i in 1..doc.Tables.Count -> doc.Tables.Item(i) ]

    let rowsOf (table: Table) : Row list = [ for i in 1..table.Rows.Count -> table.Rows.Item(i) ]

    let cellsOf (row: Row) : Cell list = [ for i in 1..row.Cells.Count -> row.Cells.Item(i) ]

    /// Word 单元格文本尾部带 \r\a，统一去掉再比
    let rowTexts (row: Row) : string list = cellsOf row |> List.map(fun c -> WordHost.DocIO.normalize c.Range.Text)

    /// 三张设备清单表：表头 + 行数认出来，顺序就是 1#/2#/3#（表6.1/6.2/6.3）
    let equipment (doc: Document) : Table list =
        let found =
            tablesOf doc
            |> List.filter(fun t ->
                let head = rowTexts (rowsOf t).[0]
                t.Rows.Count > 20 && head.Length >= 3 && head.[1] = "设备名称" && head.[2] = "规格")

        Assert.Equal(3, found.Length)
        found

    type Item =
        { N: int
          Spec: string
          Tower: string
          Material: string
          Qty: string
          Unit: string }

    /// 与 Excel 三张报价表挡板门行一致；材质按文档惯例大写 PPH（Excel 里写 pph）
    let items =
        [ { N = 1
            Spec = "DN400"
            Tower = "Φ1800×5500"
            Material = "PPH"
            Qty = "2"
            Unit = "个" }
          { N = 2
            Spec = "DN500"
            Tower = "Φ2000×5500"
            Material = "PPH"
            Qty = "2"
            Unit = "个" }
          { N = 3
            Spec = "DN800"
            Tower = "Φ3500×7500"
            Material = "PPH"
            Qty = "2"
            Unit = "个" } ]

    /// 只读清点：三张表在第几张、多少行、第 3 行是不是纤维除雾器、有没有已经补过
    let survey (dir: string) =
        Host.execOnceQuery Kind.Word (fun app ->
            match app with
            | WordApp word ->
                word.DisplayAlerts <- WdAlertLevel.wdAlertsNone
                let doc = WordHost.DocIO.openDoc word (targetPath dir) true

                try
                    equipment doc
                    |> List.mapi(fun i t ->
                        let item = items.[i]
                        let rows = rowsOf t
                        let already = rows |> List.exists(fun r -> rowTexts r |> List.contains gate)
                        // 补过之后纤维除雾器行顺延到第 4 行，认表要跟着挪
                        let first = rowTexts rows.[(if already then 3 else 2)]
                        Assert.Contains(item.Tower, first.[2])

                        {| N = item.N
                           Rows = rows.Length
                           Section = rowTexts rows.[1] |> String.concat "/"
                           First = first.[1]
                           Tower = item.Tower
                           Already = already |})
                finally
                    WordHost.DocIO.closeDoc doc
            | _ -> failwith "kind 不对：这条命令只走 Word")
        |> Report.ok

    type RowResult =
        { N: int
          Spec: string
          RowsBefore: int
          RowsAfter: int
          Before: string list
          After: string list
          FirstBefore: string
          FirstAfter: string
          Already: bool }

    type Outcome =
        { Rows: RowResult list
          Identity: string
          Backup: string
          BackedUp: bool
          Changed: bool
          SavedAt: string }

    /// 一段 Word 命令里做完：备份 → 逐表认行 → 第 3 行前插一行 → 存盘 → 只读读回
    /// 改都在内存里，全部检查通过才 Save；中途失败关文档丢改动，磁盘文件不变。
    let write (dir: string) : Outcome =
        Host.execOnceQuery Kind.Word (fun app ->
            match app with
            | WordApp word ->
                word.DisplayAlerts <- WdAlertLevel.wdAlertsNone
                let target = targetPath dir
                Assert.True(File.Exists target, sprintf "找不到技术方案：%s" target)

                let backup = backupPath dir target
                let backedUp = not (File.Exists backup)
                if backedUp then File.Copy(target, backup, false)

                let doc = WordHost.DocIO.openDoc word target false
                let identity = Doc.ofCom (fun () -> doc.FullName) (fun () -> doc.Path) |> Doc.describe

                let prepared =
                    equipment doc
                    |> List.mapi(fun i t ->
                        let item = items.[i]
                        let rows = rowsOf t
                        let already = rows |> List.exists(fun r -> rowTexts r |> List.contains gate)
                        // 双重确认没认错表：纤维除雾塔那一行必须带本线塔径
                        let firstRow = rowTexts rows.[(if already then 3 else 2)]
                        Assert.Contains(item.Tower, firstRow.[2])

                        {| Item = item
                           Table = t
                           RowsBefore = rows.Length
                           Before = rowTexts rows.[2]
                           FirstBefore = firstRow.[2]
                           Already = already |})

                let changed = prepared |> List.exists(fun p -> not p.Already)

                if changed then
                    for p in prepared do
                        if not p.Already then
                            // 插在 Word 表格第 3 行（编号 1 纤维除雾器）之前，与 Excel 第 4 行位置对应
                            let mutable beforeRow: obj = box (rowsOf p.Table).[2]
                            let newRow = p.Table.Rows.Add(&beforeRow)
                            let cs = cellsOf newRow
                            Assert.Equal(8, cs.Length)
                            // 序号列留空（Excel 里挡板门行也没编号），品牌/备注留空
                            cs.[1].Range.Text <- gate
                            cs.[2].Range.Text <- p.Item.Spec
                            cs.[3].Range.Text <- p.Item.Material
                            cs.[4].Range.Text <- p.Item.Qty
                            cs.[5].Range.Text <- p.Item.Unit

                    doc.Save()

                let savedAt = File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss")
                WordHost.DocIO.closeDoc doc

                // 存完只读重新打开，逐表读回第 3 行（新行）与第 4 行（原第 3 行）
                let check = WordHost.DocIO.openDoc word target true

                let rowsResult =
                    equipment check
                    |> List.mapi(fun i t ->
                        let p = prepared.[i]
                        let rows = rowsOf t
                        let inserted = rowTexts rows.[2]
                        let shifted = rowTexts rows.[3]

                        { N = p.Item.N
                          Spec = p.Item.Spec
                          RowsBefore = p.RowsBefore
                          RowsAfter = rows.Length
                          Before = p.Before
                          After = inserted
                          FirstBefore = p.FirstBefore
                          FirstAfter = shifted.[2]
                          Already = p.Already })

                WordHost.DocIO.closeDoc check

                { Rows = rowsResult
                  Identity = identity
                  Backup = backup
                  BackedUp = backedUp
                  Changed = changed
                  SavedAt = savedAt }
            | _ -> failwith "kind 不对：这条命令只走 Word")
        |> Report.ok

/// 真表两条：先只读清点，再写盘
type DocSyncTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``真表：清点技术方案三张设备清单表（一个文件都不动）``() =
        match DocSync.dir () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_DN_DIR=<目录> 才清点")
        | Some dir ->
            Report.requireClean "WINWORD"
            Assert.True(Directory.Exists dir, sprintf "目录不存在：%s" dir)
            let snap = DocSync.survey dir

            Report.title output (sprintf "清点 %s 的技术方案（只读，一个文件都不动）" dir)

            for s in snap do
                let item = DocSync.items |> List.find(fun i -> i.N = s.N)

                Report.line2
                    output
                    (sprintf "表6.%d" s.N)
                    "现状"
                    (sprintf "%d 行　「%s」　首个设备=「%s」　塔径 %s" s.Rows s.Section s.First s.Tower)

                Report.line2
                    output
                    (sprintf "表6.%d" s.N)
                    "待补"
                    (sprintf "「%s / %s / PPH / 2 / 个」　已补过=%b" DocSync.gate item.Spec s.Already)

            Should.equal [ 1; 2; 3 ] (snap |> List.map(fun s -> s.N))
            Report.conclusion output "只清点：三张设备清单表都认出来了，挡板门行状态如实列出，一个字没动"

    [<Fact>]
    member this.``真表：补入口电动挡板门行到技术方案三张设备清单表``() =
        match DocSync.dir () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_DN_DIR=<目录> 才补行")
        | Some dir ->
            Report.requireClean "WINWORD"
            Assert.True(Directory.Exists dir, sprintf "目录不存在：%s" dir)

            let result = DocSync.write dir
            Report.title output "真表：技术方案三张设备清单表补「入口电动挡板门」"
            Report.line output "认人" result.Identity

            Report.line
                output
                "备份"
                (if result.BackedUp then result.Backup else "已有备份，未重复生成：" + result.Backup)

            for row in result.Rows do
                Report.line2
                    output
                    (sprintf "表6.%d" row.N)
                    "插行"
                    (sprintf
                        "%d 行 → %d 行　第 3 行=「%s」"
                        row.RowsBefore
                        row.RowsAfter
                        (row.After |> List.filter(fun s -> s <> "") |> String.concat " / "))

                Report.line2
                    output
                    (sprintf "表6.%d" row.N)
                    "原行"
                    (sprintf "纤维除雾器行规格%s" (if row.FirstBefore = row.FirstAfter then "未变" else "！被改了"))

                Should.equal (row.RowsBefore + (if row.Already then 0 else 1)) row.RowsAfter
                Should.equal [ ""; DocSync.gate; row.Spec; "PPH"; "2"; "个"; ""; "" ] row.After
                Should.equal row.FirstBefore row.FirstAfter

            Report.line output "保存" (if result.Changed then result.SavedAt else "三张表已补过，未重复插行：" + result.SavedAt)
            Report.conclusion output "通过：三张表各补一行、读回八列一致，原有行顺延且规格未变，原稿已备份"
