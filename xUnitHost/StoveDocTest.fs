/// 在用户正开着的「热风炉计算书摘要.docx」里，把表 2 的管径两行改成本项目最新取值
/// （煤气支管内径 250 / 燃烧器外管内径 600，来自计算书 xlsx 的新结果）。
///
/// 两条 Fact：
///   1) 只读清点：打印表 2 里跟管径有关的行原文，核对替换锚点（`-m "*只读清点*"`）
///   2) 就地改写：整格换文字 → Save → 读回核对；改前把磁盘原件备份到同目录带时间戳
///      （`-m "*改管径*"`）
///
/// 闸门：`XUNITHOST_WORD_PID`；没设就 Skip。挂回用 `WordTask.GetApplicationAsync`，
/// 不新建、不 Quit、不关用户的文档。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open OsHost
open WordHost
open Microsoft.Office.Interop.Word

type StoveDocTest(output: ITestOutputHelper) =

    let target = @"C:\崔胜利\群有\中冶南方都市环保\热风炉计算书摘要.docx"
    let docName = "热风炉计算书摘要.docx"

    /// 表 2 里要改的锚点：格子里出现这些字才读/改
    let keywords = [ "支管"; "套管"; "助燃风" ]

    let wordPid () =
        let text = Environment.GetEnvironmentVariable "XUNITHOST_WORD_PID"

        if String.IsNullOrWhiteSpace text then
            Assert.Skip("没设 XUNITHOST_WORD_PID：这条要挂在用户正开着的那只 Word 上")
            failwith "unreachable"

        int text

    /// 在打开着的文档集合里挑出目标文档；Documents.Item 收 obj&，byref 不能进 lambda
    let pickDoc (app: Application) : Document =
        let active = app.ActiveDocument

        if active.Name = docName then
            active
        else
            let mutable found: Document = Unchecked.defaultof<Document>

            for i in 1 .. app.Documents.Count do
                let mutable key: obj = box i
                let d = app.Documents.Item(&key)

                if isNull (box found) && d.Name = docName then
                    found <- d

            if isNull (box found) then
                let names = ResizeArray<string>()

                for i in 1 .. app.Documents.Count do
                    let mutable key: obj = box i
                    names.Add(app.Documents.Item(&key).Name)

                failwithf "这只 Word 里没有打开「%s」，现有：%A" docName (List.ofSeq names)

            found

    /// 逐表逐行取单元格纯文本（合并单元格拿不到 RowIndex 的归到 -1 组，只用于诊断）
    let scanTables (doc: Document) =
        [ for ti in 1 .. doc.Tables.Count ->
              let t = doc.Tables.Item(ti)
              let cells = t.Range.Cells
              let n = cells.Count
              let rows = Collections.Generic.Dictionary<int, ResizeArray<string>>()

              for i in 1 .. n do
                  let c = cells.Item(i)
                  let ri = try c.RowIndex with _ -> -1

                  if not (rows.ContainsKey ri) then
                      rows.[ri] <- ResizeArray()

                  rows.[ri].Add(DocIO.normalize (c.Range.Text))

              ti, t.Rows.Count, t.Columns.Count, [ for kv in rows -> kv.Key, List.ofSeq kv.Value ] ]

    /// 找到「第 2 列含 keyword」的那一格，把它右边一格（数值列）整格换成 newText。
    /// 原文不含 oldAnchor 就不动：已经是 newAnchor（新值）就跳过并说明，别的一律停手。
    /// 返回 (标签, 改前, 改后, 是否真的改了)。
    let replaceValue (table: Table) (keyword: string) (oldAnchor: string) (newAnchor: string) (newText: string) =
        let cells = table.Range.Cells
        let n = cells.Count
        let mutable label: Cell = Unchecked.defaultof<Cell>
        let mutable value: Cell = Unchecked.defaultof<Cell>
        let mutable i = 1

        while i <= n && isNull (box label) do
            let c = cells.Item(i)

            if (DocIO.normalize (c.Range.Text)).Contains keyword then
                label <- c
                value <- cells.Item(i + 1)

            i <- i + 1

        if isNull (box label) then
            failwithf "没找到含「%s」的格子，未改动任何内容" keyword

        let before = DocIO.normalize (value.Range.Text)

        if before.Contains oldAnchor then
            value.Range.Text <- newText
            DocIO.normalize (label.Range.Text), before, DocIO.normalize (value.Range.Text), true
        elif before.Contains newAnchor then
            DocIO.normalize (label.Range.Text), before, before, false
        else
            failwithf "「%s」那一格的原文既不是旧值也不是新值，未改动：%s" keyword before

    [<Fact>]
    member this.``计算书摘要 只读清点表2的管径三行``() =
        Report.title output "计算书摘要　只读清点：管径那三格的原文"

        let pid = wordPid ()
        Report.line output "目标 PID" (string pid)

        let app = WordTask.GetApplicationAsync pid |> Report.await

        let name, path, tables =
            STA.run(fun () ->
                let doc = pickDoc app
                doc.Name, doc.FullName, scanTables doc)

        Report.line output "文档" (sprintf "%s  (%s)" name path)
        Report.line output "表数" (string (List.length tables))

        for ti, rowCount, colCount, _ in tables do
            Report.line output (sprintf "表 %d" ti) (sprintf "%d 行 × %d 列" rowCount colCount)

        let hits =
            [ for ti, _, _, rows in tables do
                  for ri, cells in rows do
                      let line = String.Join(" ｜ ", cells)

                      if keywords |> List.exists (fun k -> line.Contains k) then
                          yield ti, ri, cells ]

        Report.line output "命中的行" (string (List.length hits))

        for ti, ri, cells in hits do
            Report.line2 output (sprintf "表 %d" ti) (sprintf "第 %2d 行" ri) (String.Join(" ｜ ", cells))

        Should.equal true (hits |> List.exists (fun (_, _, cells) -> cells |> List.exists (fun c -> c.Contains "支管")))

        Report.conclusion output "通过：管径那三格的原文已取到，可据此定替换锚点"

    [<Fact>]
    member this.``计算书摘要 改管径两行并保存``() =
        Report.title output "计算书摘要　改管径：煤气支管 φ250、燃烧器外管 φ600"

        let pid = wordPid ()
        Report.line output "目标 PID" (string pid)

        // 1) 先备份磁盘上的原件（Word 锁着只影响写，读副本没问题）
        let stamp = DateTime.Now.ToString "yyyyMMdd-HHmm"
        let dir = Path.GetDirectoryName target

        let backup =
            Path.Combine(dir, sprintf "热风炉计算书摘要[改管径前]%s.docx" stamp)

        File.Copy(target, backup, true)
        Report.line output "备份" (sprintf "%s（%d 字节）" backup (FileInfo(backup).Length))

        let app = WordTask.GetApplicationAsync pid |> Report.await

        let results =
            STA.run(fun () ->
                let doc = pickDoc app
                let t = doc.Tables.Item(2)
                let mutable rows = []

                // 煤气支管：内径 260 → 250（DN250 不变），流速按 20℃ 工况 2470.9 m³/h 重算
                rows <-
                    rows
                    @ [ replaceValue t "煤气支管" "内径 260 mm" "内径 250 mm" "DN250（内径 250 mm）→ 13.98 m/s" ]

                // 套管环隙：外内径 518 → 600，面积 π/4×(0.600²−0.273²)=0.2242 m²，
                // 同一流量口径（20℃ 工况 11902.6 m³/h，与助燃风总管 φ550 的 13.92 m/s 一致）→ 14.75 m/s
                rows <-
                    rows
                    @ [ replaceValue
                            t
                            "套管环隙"
                            "外内径 518"
                            "外管内径 600"
                            "外管内径 600 / 内管外径 273 → 0.2242 m²，14.75 m/s" ]

                let changed = rows |> List.filter (fun (_, _, _, c) -> c) |> List.length

                if changed > 0 then
                    doc.Save()

                doc.Name, doc.FullName, rows, changed, doc.Saved)

        let name, path, rows, changed, saved = results

        for label, before, after, didChange in rows do
            Report.line2 output label "改前" before
            Report.line2 output label (if didChange then "改后" else "已是新值") after

        Report.line output "本次改动" (sprintf "%d 格" changed)
        Report.line output "保存" (sprintf "%s  %s  已落盘=%b" name path saved)
        Report.line output "保存时刻" (Report.clock (File.GetLastWriteTime(target) |> Clock.unixMs))

        // 读回：两处都必须已是新值，旧锚点必须消失
        STA.run(fun () ->
                   let doc = pickDoc app
                   let t = doc.Tables.Item(2)
                   let cells = t.Range.Cells
                   let all = [ for i in 1 .. cells.Count -> DocIO.normalize (cells.Item(i).Range.Text) ]
                   let text = String.Join(" ｜ ", all)
                   Report.line output "读回" (sprintf "表 2 全文 %d 格：%s" all.Length text)

                   Should.equal true (text.Contains "DN250（内径 250 mm）→ 13.98 m/s")
                   Should.equal true (text.Contains "外管内径 600 / 内管外径 273 → 0.2242 m²，14.75 m/s")
                   Should.equal false (text.Contains "518")
                   Should.equal false (text.Contains "260 mm"))

        // 独立核磁盘字节：docx 是 zip，去掉 XML 标签后比文字（Word 常把一段文字拆成多个 run）
        // Word 存完还占着文件，必须共享读打开，ZipFile.OpenRead 拿不到独占句柄会直接抛
        let diskPlain =
            use stream =
                new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

            use zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read)
            let entry = zip.GetEntry "word/document.xml"

            use reader = new StreamReader(entry.Open())
            Text.RegularExpressions.Regex.Replace(reader.ReadToEnd(), "<[^>]+>", "")

        let has (needle: string) = diskPlain.Contains needle

        Report.line
            output
            "磁盘核对"
            (sprintf "含「内径 250 mm」=%b，含「外管内径 600」=%b，含「0.2242」=%b，含「518」=%b"
                (has "内径 250 mm")
                (has "外管内径 600")
                (has "0.2242")
                (has "518"))

        Should.equal true (has "内径 250 mm")
        Should.equal true (has "外管内径 600")
        Should.equal true (has "0.2242")
        Should.equal false (has "518")

        Report.conclusion output "通过：表 2 两格已按计算书新口径改写并保存，读回与磁盘字节都对得上"

    [<Fact>]
    member this.``计算书摘要 插入助燃风总管一行``() =
        Report.title output "计算书摘要　插入「助燃风总管」一行（φ550 / 13.92 m/s）"

        let pid = wordPid ()
        Report.line output "目标 PID" (string pid)

        let stamp = DateTime.Now.ToString "yyyyMMdd-HHmm"
        let backup = Path.Combine(Path.GetDirectoryName target, sprintf "热风炉计算书摘要[插行前]%s.docx" stamp)
        File.Copy(target, backup, true)
        Report.line output "备份" (sprintf "%s（%d 字节）" backup (FileInfo(backup).Length))

        let app = WordTask.GetApplicationAsync pid |> Report.await

        let results =
            STA.run(fun () ->
                let doc = pickDoc app
                let t = doc.Tables.Item(2)
                let cells = t.Range.Cells
                let plain i = DocIO.normalize (cells.Item(i).Range.Text)

                // 幂等：已经有这一行就不插
                let exists = seq { 1 .. cells.Count } |> Seq.exists (fun i -> (plain i).Contains "助燃风总管")

                if exists then
                    doc.Name, 0, 0, t.Rows.Count, doc.Saved
                else
                    // 定位「套管环隙」所在的 RowIndex，新行插在它下一行之前
                    let mutable anchorRow = 0

                    for i in 1 .. cells.Count do
                        if anchorRow = 0 && (plain i).Contains "套管环隙" then
                            anchorRow <- cells.Item(i).RowIndex

                    if anchorRow = 0 then
                        failwith "没找到「套管环隙」那一行，未改动"

                    let k = anchorRow + 1
                    let mutable beforeRow: obj = box (t.Rows.Item(k))
                    let newRow = t.Rows.Add(&beforeRow)

                    if newRow.Cells.Count <> 3 then
                        failwithf "新行有 %d 格，表 2 应该是 3 格（序号/项目/数值）" newRow.Cells.Count

                    newRow.Cells.Item(1).Range.Text <- string (k - 1)
                    newRow.Cells.Item(2).Range.Text <- "助燃风总管"
                    newRow.Cells.Item(3).Range.Text <- "内径 550 mm → 13.92 m/s"

                    // 插行后整表序号按 RowIndex−1 重排（表头在第 1 行）
                    for r in (k + 1) .. t.Rows.Count do
                        t.Rows.Item(r).Cells.Item(1).Range.Text <- string (r - 1)

                    doc.Save()
                    doc.Name, k, anchorRow, t.Rows.Count, doc.Saved)

        let name, at, anchor, rowCount, saved = results

        Report.line output "插入位置" (if at = 0 then "已存在，未插" else sprintf "第 %d 行（「套管环隙」第 %d 行之后）" at anchor)
        Report.line output "表 2 行数" (string rowCount)
        Report.line output "保存" (sprintf "%s  已落盘=%b" name saved)
        Report.line output "保存时刻" (Report.clock (File.GetLastWriteTime(target) |> Clock.unixMs))

        // 读回：新行在位、序号连续、原文没丢
        let dump =
            STA.run(fun () ->
                let doc = pickDoc app
                let t = doc.Tables.Item(2)
                let rows = t.Rows
                let lines =
                    [ for r in 1 .. rows.Count ->
                          let cs = rows.Item(r).Cells
                          String.Join(" ｜ ", [ for c in 1 .. cs.Count -> DocIO.normalize (cs.Item(c).Range.Text) ]) ]

                let header = lines.Head
                let body = lines.Tail
                // 明细行第一格应是 1,2,3…连续的序号
                let serials = [ for l in body -> l.Split('｜').[0].Trim() ]
                let expected = [ for i in 1 .. List.length body -> string i ]
                String.Join(" ／ ", lines), serials = expected, List.length body, header)

        let all, serialOk, bodyCount, header = dump

        Report.line output "读回" all
        Report.line output "表头" header

        Should.equal true (all.Contains "助燃风总管")
        Should.equal true (all.Contains "内径 550 mm → 13.92 m/s")
        Should.equal true (all.Contains "外管内径 600 / 内管外径 273 → 0.2242 m²，14.75 m/s")
        Should.equal true serialOk
        Should.equal 12 bodyCount

        // 磁盘字节：剥掉 XML 标签后再比
        let diskPlain =
            use stream =
                new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

            use zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read)
            let entry = zip.GetEntry "word/document.xml"

            use reader = new StreamReader(entry.Open())
            Text.RegularExpressions.Regex.Replace(reader.ReadToEnd(), "<[^>]+>", "")

        Report.line output "磁盘核对" (sprintf "含「助燃风总管」=%b，含「内径 550 mm」=%b" (diskPlain.Contains "助燃风总管") (diskPlain.Contains "内径 550 mm"))
        Should.equal true (diskPlain.Contains "助燃风总管")
        Should.equal true (diskPlain.Contains "内径 550 mm")

        Report.conclusion output "通过：表 2 多了一行「助燃风总管」，序号重排连续，原有各格未受影响"
