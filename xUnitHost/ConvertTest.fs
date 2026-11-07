/// 转换工具的测试：在临时目录里造一份 .doc 和一份 .xls，转成新格式，逐字核对。
/// 先在这里证明「转得对、原件不动」，才敢拿它去动真目录里的 28 份文件。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Core
open Microsoft.Office.Interop.PowerPoint

type ConvertTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``转换：.doc 转 .docx、.xls 转 .xlsx，内容一字不差，原件不动``() =
        // 转换要自己起 Word / Excel：本机已经有它们在跑就没法做（也绝不能碰你手上的文档）
        Report.requireClean "WINWORD"
        Report.requireClean "EXCEL"

        let dir =
            Path.Combine(Path.GetTempPath(), "docconvert-test-" + Guid.NewGuid().ToString("N").Substring(0, 8))

        Directory.CreateDirectory dir |> ignore

        let docPath = Path.Combine(dir, "样例.doc")
        let xlsPath = Path.Combine(dir, "样例.xls")
        let body = "第一段：转换要一字不差\r第二段：含中文和数字 12345"

        try
            Report.title output "转换：.doc → .docx、.xls → .xlsx，内容一字不差，原件不动"

            // ---- 造两份老格式文件 ----
            WordHost.Host.exec(fun app ->
                let mutable template: obj = Type.Missing
                let mutable newTemplate: obj = box false
                let mutable docType: obj = box Microsoft.Office.Interop.Word.WdNewDocumentType.wdNewBlankDocument
                let mutable visible: obj = box false
                let doc = app.Documents.Add(&template, &newTemplate, &docType, &visible)
                doc.Content.Text <- body
                let mutable fileName: obj = box docPath
                let mutable format: obj = box 0 // wdFormatDocument = .doc
                doc.SaveAs(&fileName, &format) |> ignore
                let mutable save: obj = box 0
                doc.Close(&save))
            |> Report.ok

            ExcelHost.Host.exec(fun app ->
                let book = app.Workbooks.Add()
                let sheet = book.Worksheets.[1] :?> Microsoft.Office.Interop.Excel.Worksheet
                sheet.Range("A1").Value2 <- "项目"
                sheet.Range("B1").Value2 <- 123.45
                sheet.Range("A2").Value2 <- "合计"
                sheet.Range("B2").Formula <- "=B1*2"
                book.SaveAs(box xlsPath, box -4143) |> ignore // xlExcel8 = .xls
                book.Close(box false))
            |> Report.ok

            Report.line output "样例" (sprintf "造好 %s（%d 字节）、%s（%d 字节）" (Path.GetFileName docPath) (FileInfo docPath).Length (Path.GetFileName xlsPath) (FileInfo xlsPath).Length)

            // ---- 转换 ----
            let items: Item list =
                [ { Source = docPath; Target = Path.ChangeExtension(docPath, ".docx") }
                  { Source = xlsPath; Target = Path.ChangeExtension(xlsPath, ".xlsx") } ]

            let outcomes = Convert.all items

            for o in outcomes do
                Report.line output (Path.GetFileName o.Source) (sprintf "%s —— %s" o.Status o.Detail)

            let word = outcomes |> List.find(fun o -> o.Source = docPath)
            let excel = outcomes |> List.find(fun o -> o.Source = xlsPath)

            // ---- 断言：转成功、原件还在 ----
            Should.equal "转了" word.Status
            Should.equal "转了" excel.Status
            Assert.True(File.Exists(Path.ChangeExtension(docPath, ".docx")), "没生成 docx")
            Assert.True(File.Exists(Path.ChangeExtension(xlsPath, ".xlsx")), "没生成 xlsx")
            Assert.True(File.Exists docPath, "转换不该动原件（删是另一条命令的事）")
            Assert.True(File.Exists xlsPath, "转换不该动原件")

            // ---- 断言：内容一字不差 ----
            let beforeText = Convert.readWordText docPath |> Report.ok
            let afterText = Convert.readWordText(Path.ChangeExtension(docPath, ".docx")) |> Report.ok
            Should.equal beforeText afterText
            Should.equal (body.Replace("\r", "")) afterText
            Report.line output "核对 docx" (sprintf "正文 %d 字，与原件逐字相同「%s」" afterText.Length afterText)

            let beforePrint = Convert.readBookPrint xlsPath |> Report.ok
            let afterPrint = Convert.readBookPrint(Path.ChangeExtension(xlsPath, ".xlsx")) |> Report.ok
            Should.equal beforePrint afterPrint
            Assert.Contains("合计", afterPrint)
            Report.line output "核对 xlsx" (sprintf "每张表的格子值全对上（指纹 %d 字符）" afterPrint.Length)

            Report.conclusion output "通过：转得对、原件一个字没动"
        finally
            (try
                Directory.Delete(dir, true)
             with _ ->
                 ())

    [<Fact>]
    member this.``转换：四条路各管各的后缀和格式码，互不串味``() =
        // 纯逻辑，不碰 Word / Excel：验证四条转换各自的规格，两个统一入口各自选路
        let dir = Path.Combine(Path.GetTempPath(), "convert-spec")
        let doc = Path.Combine(dir, "样例.doc")
        let xls = Path.Combine(dir, "样例.xls")

        let docx = WordHost.DocToDocx.targetOf doc
        let docm = WordHost.DocToDocm.targetOf doc
        let xlsx = ExcelHost.XlsToXlsx.targetOf xls
        let xlsm = ExcelHost.XlsToXlsm.targetOf xls

        Report.title output "转换：四条路各管各的后缀和格式码，互不串味"

        Report.line output ".doc 无宏 → docx" (sprintf "%s　格式码 %d　WordHost" (Path.GetFileName docx) (WordHost.DocToDocx.format :?> int))
        Report.line output ".doc 有宏 → docm" (sprintf "%s　格式码 %d　WordHost" (Path.GetFileName docm) (WordHost.DocToDocm.format :?> int))
        Report.line output ".xls 无宏 → xlsx" (sprintf "%s　格式码 %d　ExcelHost" (Path.GetFileName xlsx) (ExcelHost.XlsToXlsx.format :?> int))
        Report.line output ".xls 有宏 → xlsm" (sprintf "%s　格式码 %d　ExcelHost" (Path.GetFileName xlsm) (ExcelHost.XlsToXlsm.format :?> int))
        Report.line output "统一入口" "DocUpgrade / XlsUpgrade：有宏保宏（.docm / .xlsm），没宏 .docx / .xlsx；判不出宏就失败、原件保留"
        Report.line output "扫描时估的" (sprintf "%s / %s（还没判宏，先按无宏占位）" (Path.GetFileName(Paths.targetOf doc)) (Path.GetFileName(Paths.targetOf xls)))

        Should.equal (Path.ChangeExtension(doc, ".docx")) docx
        Should.equal (Path.ChangeExtension(doc, ".docm")) docm
        Should.equal (Path.ChangeExtension(xls, ".xlsx")) xlsx
        Should.equal (Path.ChangeExtension(xls, ".xlsm")) xlsm
        Should.equal 12 (WordHost.DocToDocx.format :?> int)
        Should.equal 13 (WordHost.DocToDocm.format :?> int)
        Should.equal 51 (ExcelHost.XlsToXlsx.format :?> int)
        Should.equal 52 (ExcelHost.XlsToXlsm.format :?> int)
        Should.equal (Path.ChangeExtension(doc, ".docx")) (Paths.targetOf doc)
        Should.equal (Path.ChangeExtension(xls, ".xlsx")) (Paths.targetOf xls)
        Should.equal WordHost.DocUpgrade.Docm (WordHost.DocUpgrade.route true)
        Should.equal WordHost.DocUpgrade.Docx (WordHost.DocUpgrade.route false)
        Should.equal ExcelHost.XlsUpgrade.Xlsm (ExcelHost.XlsUpgrade.route true)
        Should.equal ExcelHost.XlsUpgrade.Xlsx (ExcelHost.XlsUpgrade.route false)

        Report.conclusion output "通过：两个统一入口各在各的 host 库，后缀、格式码与选路规则互不串味"

// ---------------- 真目录：平时不跑，靠环境变量叫起来 ----------------
//   XUNITHOST_CONVERT_DIR     要处理哪个目录（不设就跳过，正常跑测试不受影响）
//   XUNITHOST_CONVERT_BATCH   送回收站 / 报告的批次，默认 10（安全上限，别放大）
//   XUNITHOST_CONVERT_CHUNK   一个 COM 进程里连续转多少份，默认 200
//   XUNITHOST_CONVERT_BACKUP  备份目录；默认 <root>-bak<戳>，在目标目录之外
//   XUNITHOST_CONVERT_KIND    doc|xls|ppt，只跑一类，建议分三轮；不设 = 全跑
//   XUNITHOST_CONVERT_LIMIT   本轮最多处理 N 份（翻倍法 1→2→4→…）；不设 = 不限
//   XUNITHOST_CONVERT_SKIP    已知失败清单（每行一个相对路径），本轮排除、不再反复重试
// 顺序：完整备份 → 重名让位改名 → 分块转换 → 逐份读回核对 → 只把核对通过的原件送回收站。

/// 真目录里的杂活：报告里用相对路径、删原件只认「核对通过」
module internal RealFolder =

    open Microsoft.VisualBasic.FileIO

    let rel (root: string) (path: string) =
        let head = root.TrimEnd(Path.DirectorySeparatorChar) + string Path.DirectorySeparatorChar

        if path.StartsWith(head, StringComparison.OrdinalIgnoreCase) then
            path.Substring head.Length
        else
            path

    let mark (status: string) =
        match status with
        | "转了" -> "✓ 转了"
        | "已有·一致" -> "✓ 已有"
        | "跳过" -> "- 跳过"
        | _ -> "× 失败"

    let root () =
        match Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_DIR" with
        | null
        | "" -> None
        | value -> Some(Path.GetFullPath value)

    /// 本轮只跑一类（doc / xls / ppt）；不设 KIND 就全跑。
    let kind () = Paths.kindFromEnv ()

    /// 该类型还剩多少（清点 / 收尾统计用，不受本轮上限影响）
    let items (root: string) : Item list =
        Paths.scan root |> Paths.filterByKind (kind ()) |> Paths.sortItems

    /// 本轮真正要动的：翻倍法靠 LIMIT 取前 N 份，已知失败的不再挡在队首。
    let roundItems (root: string) : Item list =
        Paths.selectForRound root (kind ()) (Paths.scan root)

    let limitLabel () =
        match Paths.limitFromEnv () with
        | Some n -> sprintf "%d 份" n
        | None -> "不限（全量）"

    let skipCount (root: string) = (Paths.skipFromEnv root).Count

    let private positive (name: string) (fallback: int) =
        let text = defaultArg (Option.ofObj (Environment.GetEnvironmentVariable name)) ""

        match Int32.TryParse text with
        | true, n when n > 0 -> n
        | _ -> fallback

    /// 删原件 / 报告的批次，安全上限：每批最多 10 份
    let batchSize () = positive "XUNITHOST_CONVERT_BATCH" 10

    /// 一个 COM 进程里连续转多少份（换进程是为了别让 Word 攒太多内存）
    let chunkSize () = positive "XUNITHOST_CONVERT_CHUNK" 200

    /// 备份目录默认在目标目录之外：<root>-bak<戳>
    let backupRoot (root: string) =
        match Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_BACKUP" with
        | null
        | "" ->
            let r = root.TrimEnd(Path.DirectorySeparatorChar)
            Path.Combine(Path.GetDirectoryName r, Path.GetFileName r + "-bak" + Paths.RenameStamp)
        | value -> Path.GetFullPath value

    /// 完整备份：保持目录结构、逐份比大小。任何一份对不上就抛，绝不带着残缺备份往下走。
    /// 已经备份过且大小一致的直接复用（中断后可续跑）。
    let backupAll (root: string) (backup: string) (items: Item list) (say: string -> unit) : int =
        let mutable n = 0

        for i in items do
            let dest = Path.Combine(backup, rel root i.Source)
            Directory.CreateDirectory(Path.GetDirectoryName dest) |> ignore
            let src = FileInfo i.Source

            if File.Exists dest && (FileInfo dest).Length = src.Length then
                ()
            else
                File.Copy(i.Source, dest, true)
                let copied = FileInfo dest

                if copied.Length <> src.Length then
                    failwithf "备份对不上：%s（原 %d / 备份 %d）" (rel root i.Source) src.Length copied.Length

            n <- n + 1

            if n % 200 = 0 then
                say (sprintf "备份进度 %d/%d" n items.Length)

        n

    /// 重名让位：现有同名新格式文件改名成 X-bak<戳>.docx，旧文件才能照常「转了」。
    /// 改名失败的旧文件原样跳过，绝不带病转换。
    let renameExisting (root: string) (items: Item list) (say: string -> unit) : int * Set<string> =
        let mutable renamed = 0
        let mutable skipped = Set.empty<string>

        for i in items do
            for old in Paths.existingTargetsOf i.Source do
                let mutable target = Paths.renameOf old
                let mutable n = 1

                while File.Exists target do
                    target <-
                        Path.Combine(
                            Path.GetDirectoryName old,
                            Path.GetFileNameWithoutExtension old
                            + "-bak" + Paths.RenameStamp + "-" + string n
                            + Path.GetExtension old
                        )

                    n <- n + 1

                try
                    let info = FileInfo old
                    if info.IsReadOnly then info.IsReadOnly <- false
                    File.Move(old, target)
                    renamed <- renamed + 1
                    say (sprintf "改名 %s → %s" (rel root old) (Path.GetFileName target))
                with ex ->
                    skipped <- skipped.Add i.Source
                    say (sprintf "× 改名失败，这份旧文件跳过：%s —— %s" (rel root old) ex.Message)

        renamed, skipped

    /// 不是 File.Delete：送进回收站，随时能还原
    let toRecycleBin (path: string) : unit =
        FileSystem.DeleteFile(path, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin)

    /// 只送核对通过的：新文件必须确实在，原件只读先清属性
    let deleteVerified (root: string) (outcomes: Outcome list) (say: string -> unit) : int * int =
        let mutable removed = 0
        let mutable kept = 0

        for o in outcomes do
            let target =
                if File.Exists o.Target then o.Target else Paths.existingTargetOf o.Source

            if not (File.Exists o.Source) then
                say (sprintf "- 本来就不在了：%s" (rel root o.Source))
            elif not (File.Exists target) then
                kept <- kept + 1
                say (sprintf "× 新文件不见了，原件留下：%s" (rel root o.Source))
            else
                try
                    let info = FileInfo o.Source
                    if info.IsReadOnly then info.IsReadOnly <- false
                    toRecycleBin o.Source
                    removed <- removed + 1
                    say (sprintf "回收站 %s（新文件 %s 在）" (Path.GetFileName o.Source) (Path.GetFileName target))
                with ex ->
                    kept <- kept + 1
                    say (sprintf "× 送不进回收站：%s —— %s" (rel root o.Source) ex.Message)

        removed, kept

/// 真目录上的两条 Fact：一条只清点，一条真转真送回收站。
type ConvertFolderTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``真目录：先清点，看看要转哪些（一个文件都不动）``() =
        match RealFolder.root () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_CONVERT_DIR=<目录> 才清点")
        | Some root ->
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let items = RealFolder.items root
            let existing = items |> List.filter(fun i -> Paths.existingTargetsOf i.Source |> List.isEmpty |> not)

            Report.title output (sprintf "清点 %s" root)
            Report.line output "本轮类型" (Paths.kindLabel (RealFolder.kind ()))
            Report.line output "本轮上限" (RealFolder.limitLabel ())
            Report.line output "已知失败跳过" (sprintf "%d 份" (RealFolder.skipCount root))

            Report.line
                output
                "待办"
                (sprintf
                    "%d 份（.doc %d / .xls %d / .ppt %d），其中 %d 份已经有同名新文件"
                    items.Length
                    (items |> List.filter(fun i -> Paths.isWord i.Source) |> List.length)
                    (items |> List.filter(fun i -> Paths.isExcel i.Source) |> List.length)
                    (items |> List.filter(fun i -> Paths.isPpt i.Source) |> List.length)
                    existing.Length)

            for i in items do
                let hasTarget = Paths.existingTargetsOf i.Source |> List.isEmpty |> not
                Report.line output (if hasTarget then "已有新文件" else "待转") (RealFolder.rel root i.Source)

            Report.conclusion output (sprintf "只清点，一个文件都没动：%d 份待办，%d 份已有新文件" items.Length existing.Length)

    [<Fact>]
    member this.``真目录：把老格式转成新格式，核对通过才把原件送回收站``() =
        match RealFolder.root () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_CONVERT_DIR=<目录> 才真转")
        | Some root ->
            // 转换要自己起 Word / Excel / PowerPoint：本机已经有它们在跑就先别动
            Report.requireClean "WINWORD"
            Report.requireClean "EXCEL"
            Report.requireClean "POWERPNT"
            Assert.True(Directory.Exists root, sprintf "目录不存在：%s" root)

            let items = RealFolder.roundItems root
            let batch = RealFolder.batchSize ()
            let chunk = RealFolder.chunkSize ()
            let backup = RealFolder.backupRoot root
            let say (text: string) = output.WriteLine("    " + text)

            Report.title output (sprintf "转 %s" root)

            Report.line
                output
                "待办"
                (sprintf
                    "%d 份（.doc %d / .xls %d / .ppt %d），每块 %d 份、每批 %d 份"
                    items.Length
                    (items |> List.filter(fun i -> Paths.isWord i.Source) |> List.length)
                    (items |> List.filter(fun i -> Paths.isExcel i.Source) |> List.length)
                    (items |> List.filter(fun i -> Paths.isPpt i.Source) |> List.length)
                    chunk
                    batch)

            Report.line output "本轮类型" (Paths.kindLabel (RealFolder.kind ()))
            Report.line output "本轮上限" (RealFolder.limitLabel ())
            Report.line output "已知失败跳过" (sprintf "%d 份" (RealFolder.skipCount root))
            Report.line output "备份目录" backup

            // ---- 1. 完整备份：备份目录必须在目标目录之外，逐份比大小 ----
            let head = root.TrimEnd(Path.DirectorySeparatorChar) + string Path.DirectorySeparatorChar
            Assert.False(backup.StartsWith(head, StringComparison.OrdinalIgnoreCase), "备份目录不能落在目标目录里")

            let backed = RealFolder.backupAll root backup items say
            Should.equal items.Length backed
            Report.line output "备份" (sprintf "%d/%d 份，逐份比大小对上了" backed items.Length)

            // ---- 2. 重名让位：现有新文件改名，旧文件照常原位升级 ----
            let renamed, skipped = RealFolder.renameExisting root items say
            let todo = items |> List.filter(fun i -> not (skipped.Contains i.Source))
            Report.line output "重名让位" (sprintf "改名 %d 个现有新文件；改名失败跳过 %d 份旧文件" renamed skipped.Count)

            // ---- 3. 分块转换 + 逐份报告 + 每批 10 份送回收站 ----
            let mutable all = []
            let mutable totalRemoved = 0
            let mutable totalKept = 0

            for index, block in todo |> List.chunkBySize chunk |> List.indexed do
                let label = sprintf "第 %d 块" (index + 1)
                Report.line2 output label "份数" (string block.Length)

                let outcomes = Convert.all block

                for o in outcomes do
                    let note = if Convert.canDelete o then "" else "　—— " + o.Detail
                    Report.line output (RealFolder.mark o.Status) (RealFolder.rel root o.Source + note)

                all <- all @ outcomes

                for sub in outcomes |> List.filter Convert.canDelete |> List.chunkBySize batch do
                    let removed, kept = RealFolder.deleteVerified root sub say
                    totalRemoved <- totalRemoved + removed
                    totalKept <- totalKept + kept

                Report.line2
                    output
                    label
                    "小计"
                    (outcomes
                     |> List.countBy(fun o -> o.Status)
                     |> List.map(fun (status, n) -> sprintf "%s %d" status n)
                     |> String.concat " / ")

            let good = all |> List.filter Convert.canDelete
            let bad = all |> List.filter(Convert.canDelete >> not)

            Report.line output "核对通过" (sprintf "%d 份" good.Length)
            Report.line output "没通过" (sprintf "%d 份（原件都还在）" bad.Length)

            for o in bad do
                Report.line output "　原因" (sprintf "%s —— %s" (RealFolder.rel root o.Source) o.Detail)

            Report.line output "原件入回收站" (sprintf "送进去 %d 份，送不进留下 %d 份" totalRemoved totalKept)
            let left = RealFolder.items root
            Report.line output "目录现状" (sprintf "本轮类型还剩 %d 份老格式" left.Length)

            // 失败的一份都不能把原件弄丢；备份目录必须真的有东西
            for o in bad do
                Assert.True(File.Exists o.Source, sprintf "核对没通过却把原件弄丢了：%s" o.Source)

            Assert.True(Directory.Exists backup, "备份目录不见了")
            Report.conclusion
                output
                (sprintf
                    "完成：%d 份转了/已有一致，原件入回收站 %d 份；没通过 %d 份（原件保留）；备份在 %s"
                    good.Length totalRemoved bad.Length backup)

// ---------------- 新增通道的验证：.ppt 批量分派 + 三道安全兜底 ----------------
// 放在文件末尾：上面的 RealFolder 已经定义好，安全兜底那条才调得到。

type ConvertBatchTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``转换：.ppt 也走批量分派，幻灯片文字一字不差，原件不动``() =
        Report.requireClean "POWERPNT"

        let beforeProc = Report.procIds "POWERPNT"
        let dir = Path.Combine(Path.GetTempPath(), "convert-ppt-" + Guid.NewGuid().ToString("N").Substring(0, 8))
        Directory.CreateDirectory dir |> ignore
        let ppt = Path.Combine(dir, "样例.ppt")
        let body = "批量分派：第一页\r含中文和数字 12345"

        try
            Report.title output "转换：.ppt 走 Convert.all 批量分派，幻灯片文字一字不差，原件不动"

            // 造一份老格式 .ppt（一个 PowerPoint 进程）
            PowerPointHost.Host.exec(fun app ->
                let pres = app.Presentations.Add(MsoTriState.msoFalse)
                let slide = pres.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

                let shape =
                    slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        50.0f, 50.0f, 400.0f, 100.0f)

                shape.TextFrame.TextRange.Text <- body
                pres.SaveAs(ppt, PpSaveAsFileType.ppSaveAsPresentation, MsoTriState.msoTriStateMixed)
                pres.Close())
            |> Report.ok

            // 扫描要认得它，分派要把它送进 PowerPoint 那条路
            let items = Paths.scan dir
            Should.equal 1 items.Length
            Should.equal true (Paths.isPpt items.Head.Source)

            let outcome = (Convert.all items).Head
            Report.line output "分派结果" (sprintf "%s —— %s" outcome.Status outcome.Detail)

            Should.equal "转了" outcome.Status
            Assert.True(File.Exists(Path.ChangeExtension(ppt, ".pptx")), "没生成 pptx")
            Assert.True(File.Exists ppt, "转换不该动原件")

            let before = Convert.readPptPrint ppt |> Report.ok
            let after = Convert.readPptPrint(Path.ChangeExtension(ppt, ".pptx")) |> Report.ok
            Should.equal before after
            Report.line output "核对 pptx" (sprintf "幻灯片文字 %d 字符，与原件逐字相同" after.Length)
            Report.conclusion output "通过：Convert.all 认得 .ppt，转完读回逐字比对，原件一个字没动"
        finally
            let leftovers = Set.difference (Report.procIds "POWERPNT") beforeProc
            if not leftovers.IsEmpty then OsHost.Proc.kill leftovers
            (try Directory.Delete(dir, true) with _ -> ())

    [<Fact>]
    member this.``转换：备份、重名让位、送回收站，在临时目录里全跑一遍``() =
        let dir = Path.Combine(Path.GetTempPath(), "convert-safe-" + Guid.NewGuid().ToString("N").Substring(0, 8))
        let root = Path.Combine(dir, "root")
        let backup = Path.Combine(dir, "backup")
        Directory.CreateDirectory root |> ignore

        let src = Path.Combine(root, "样例.doc")
        let existing = Path.Combine(root, "样例.docx")
        File.WriteAllText(src, "旧格式内容 0123456789", System.Text.Encoding.UTF8)
        File.WriteAllText(existing, "已经存在的新格式内容", System.Text.Encoding.UTF8)

        let items: Item list = [ { Source = src; Target = Path.ChangeExtension(src, ".docx") } ]
        let say (text: string) = output.WriteLine("    " + text)

        try
            Report.title output "转换前的三道保险：完整备份 / 重名让位 / 送回收站"

            // 1. 备份：保持目录结构、逐份比大小
            let n = RealFolder.backupAll root backup items say
            Should.equal 1 n
            let copy = Path.Combine(backup, "样例.doc")
            Assert.True(File.Exists copy, "备份没落盘")
            Should.equal (FileInfo src).Length (FileInfo copy).Length
            Report.line output "备份" (sprintf "%s（%d 字节，与原件一致）" copy (FileInfo copy).Length)

            // 2. 重名让位：现有新文件改名，旧文件才有位置「转了」
            let renamed, skipped = RealFolder.renameExisting root items say
            Should.equal 1 renamed
            Should.equal 0 skipped.Count
            let moved = Paths.renameOf existing
            Assert.True(File.Exists moved, "改名后的新文件不在")
            Assert.False(File.Exists existing, "原位置还占着名字")
            Report.line output "改名" (sprintf "%s → %s" (Path.GetFileName existing) (Path.GetFileName moved))

            // 3. 回收站：不是 File.Delete，备份还在
            RealFolder.toRecycleBin src
            Assert.False(File.Exists src, "送回收站后原路径还在")
            Assert.True(File.Exists copy, "备份还在")

            Report.conclusion output "通过：备份逐份比大小、重名改名让位、原件走回收站而不是永久删除"
        finally
            (try Directory.Delete(dir, true) with _ -> ())
