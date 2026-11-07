/// 两家之间打个招呼：一句话从 A 传到 B，B 读回来再声明自己收到了。
///
/// **一个 Fact 只管一段通道**，只开这一段真正需要的两程序：
///   · Word → Excel
///   · Excel → AutoCAD（AutoCAD 没在跑时默认不新起，除非显式 XUNITHOST_LAUNCH_ACAD=1）
///
/// 动笔之前一律先「认人」：核对**文件名 + 最后实质保存时间**，只往「还没落盘的新文档」里写，
/// 落到磁盘上的一律停手 —— 那是用户手上的文件。
///
/// 公用的报告、认人、忙重试在前部 `Common.fs`；线程/进程/窗口那些在 `OsHost`。
namespace XUnitHost

open System
open System.Diagnostics
open System.IO
open System.Threading
open Xunit
open OsHost

/// 各自「说一句」：新建一份临时文档，先认人，再写，读回来。
/// 这是测试自己的嘴，不属于任何 host 库的本职。
/// 新建的那一份**要交回去**：收尾时由调用方关掉（不保存）—— 碰到的要是用户开着的实例，
/// 也不能在人家窗口里留一份垃圾。
module internal Say =

    /// Word：新建一张空白文档，**先认人再写**，返回（身份，写进去的话，那份文档）
    let word
        (app: Microsoft.Office.Interop.Word.Application)
        (greeting: string)
        : Doc.Identity * string * obj =
        let mutable template: obj = Type.Missing
        let mutable newTemplate: obj = box false
        let mutable docType: obj = box Microsoft.Office.Interop.Word.WdNewDocumentType.wdNewBlankDocument
        let mutable visible: obj = box false
        let doc = app.Documents.Add(&template, &newTemplate, &docType, &visible)

        // 新建的文档不该有路径；有路径就说明点错了人
        let id = Doc.ofCom (fun () -> doc.FullName) (fun () -> doc.Path)
        Doc.ensureFresh "Word 文档" id

        doc.Content.Text <- greeting
        // Word 的正文末尾带段落标记，读回来统一去掉
        id, doc.Content.Text.Trim Report.marks, box doc

    /// Excel：新建一本工作簿，先认人、再把话塞进 A1，返回（身份，单元格里的内容，那本工作簿）
    /// Excel 的 interop 参数是可省略的 `obj`（不是 `obj ref`），所以不用摆 ref 变量
    let excel
        (app: Microsoft.Office.Interop.Excel.Application)
        (heard: string)
        (greeting: string)
        : Doc.Identity * string * obj =
        let book = app.Workbooks.Add()
        let id = Doc.ofCom (fun () -> book.FullName) (fun () -> book.Path)
        Doc.ensureFresh "Excel 工作簿" id

        let said =
            if String.IsNullOrEmpty heard then greeting else heard + " " + greeting

        let sheet = book.Worksheets.[1] :?> Microsoft.Office.Interop.Excel.Worksheet
        let cell = sheet.Range("A1")
        cell.Value2 <- said
        id, (string cell.Value2).Trim(), box book

    /// AutoCAD：**新开一张图**（绝不往当前图纸里写），先认人再写，
    /// 返回（身份，读回来的文字，那张图）。
    /// 取实例用传进来的 getApp —— 连接真断了要重新取一个新的（旧的 RCW 已经跟着进程没了）。
    let acad
        (getApp: unit -> Autodesk.AutoCAD.Interop.AcadApplication)
        (heard: string)
        (greeting: string)
        : Doc.Identity * string * obj =
        let once () =
            let app = getApp ()

            // 冷启动：GetAcadState().IsQuiescent 是 ActiveX 专给进程外调用方准备的探针。
            // 要连着 1.5 秒都安静才算真就绪 —— 刚起来时它会短暂地报 true，紧接着动它就 RPC_E_SERVERFAULT / 进程直接崩。
            let clock = Stopwatch.StartNew()
            let mutable streak = 0

            while streak < 5 && clock.ElapsedMilliseconds < 240000L do
                Thread.Sleep 300

                if
                    (try
                        app.GetAcadState().IsQuiescent
                     with _ ->
                         false)
                then
                    streak <- streak + 1
                else
                    streak <- 0

            if streak < 5 then
                failwith "AutoCAD 240 秒内仍未就绪"

            Thread.Sleep 1500

            // 一律新开一张临时图：AutoCAD 只能往「当前图纸」里写，直接写就等于污染用户手上的图纸
            let doc = Busy.retry 180000 (fun () -> app.Documents.Add())

            let id = Doc.ofCom (fun () -> doc.FullName) (fun () -> doc.Path)

            Doc.ensureFresh "AutoCAD 图纸" id

            let text =
                Busy.retry 180000 (fun () ->
                    doc.ModelSpace.AddText(heard + " " + greeting, box [| 0.0; 0.0; 0.0 |], 2.5))

            id, (Busy.retry 120000 (fun () -> text.TextString)).Trim(), (box doc)

        let rec attempt (left: int) =
            try
                once ()
            with ex when left > 0 && Busy.isDeath ex ->
                Thread.Sleep 3000
                attempt (left - 1)

        attempt 2

// ---------------- 测试 ----------------
// 一段通道一个 Fact：只开这段真正需要的两家，各自起、各自收。

type HostGreetingTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``认人：没落盘的才写，磁盘上的一律停手``() =
        let dir =
            Path.Combine(Path.GetTempPath(), "xUnitHost-认人-" + Guid.NewGuid().ToString("N").Substring(0, 8))

        Directory.CreateDirectory dir |> ignore
        let path = Path.Combine(dir, "别人的方案.docx")

        try
            File.WriteAllText(path, "这是用户手上的文件")
            File.SetLastWriteTime(path, DateTime(2026, 11, 5, 10, 30, 0))

            let onDisk = Doc.identify path (Path.GetDirectoryName path)

            Report.title output "动笔之前先认人：文件名 + 最后实质保存时间"
            Report.line output "磁盘上" (Doc.describe onDisk)

            // 磁盘上的文件必须被拦下来，而且报错里要能看见是哪个文件、什么时候存的
            let blocked =
                try
                    Doc.ensureFresh "Word 文档" onDisk
                    None
                with ex ->
                    Some ex.Message

            match blocked with
            | None -> failwith "磁盘上的文件没被拦住，一个字都不该写"
            | Some message -> Report.line output "拦下" message

            // 测试自己新建、还没存盘的（Word 里叫「文档1」）：放行
            let fresh = Doc.identify "文档1" ""
            Doc.ensureFresh "Word 文档" fresh
            Report.line output "新文档" (Doc.describe fresh + " → 放行")

            Report.conclusion output "通过：先看名字和保存时间，磁盘上的一律不碰"
        finally
            (try
                Directory.Delete(dir, true)
             with _ ->
                 ())

    [<Fact>]
    member this.``Word 和 Excel 打招呼：一句话传过去，Excel 接着回一句``() =
        let beforeWord = Report.procIds "WINWORD"
        let beforeExcel = Report.procIds "EXCEL"

        let wordGreeting = "Hello Excel. -- Word"
        let excelGreeting = "Excel 收到。Hello Word. -- Excel"

        /// 收尾时被强制结束的进程 —— 正常退出应该一个都没有
        let killed = ResizeArray<string>()

        let track (result: string option) =
            match result with
            | Some name -> killed.Add name
            | None -> ()

        Report.title output "Word → Excel：一句话传过去，Excel 接着回一句（只开这两家）"

        Sta.here (fun () ->
            let word = WordHost.Host.GetOrCreateApplication ()
            let excel = ExcelHost.Host.GetOrCreateApplication ()

            // 别弹任何确认框，否则测试会卡在模态框上；记下原值，走的时候还原（万一碰到的是别人的实例）
            let excelAlerts =
                try
                    let original = excel.DisplayAlerts
                    excel.DisplayAlerts <- false
                    original
                with _ ->
                    true

            (try
                word.DisplayAlerts <- Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
             with _ ->
                 ())

            /// 这次自己新建的那两份：收尾关掉（不保存），用户窗口里不留垃圾
            let mutable wordDoc: obj = null

            let mutable excelBook: obj = null

            try
                // 1. Word 新建一份、先认人、再开口
                let wordId, heardByExcel, made = Say.word word wordGreeting
                wordDoc <- made
                Report.line output "Word" (Doc.describe wordId)
                Report.line output "Word 说" heardByExcel

                // 2. Excel 新建一本、先认人、把听到的那句写进 A1
                let excelId, heardBack, made = Say.excel excel heardByExcel excelGreeting
                excelBook <- made
                Report.line output "Excel" (Doc.describe excelId)
                Report.line output "Excel 说" heardBack

                // 话必须首尾相接：Excel 复述的正是 Word 说的那句原话
                Assert.Equal(wordGreeting, heardByExcel)
                Assert.StartsWith(wordGreeting, heardBack)
                Assert.Contains(excelGreeting, heardBack)
                Report.line output "对照" "Excel 复述的那一句，跟 Word 说的逐字相同"
            finally
                // 先关掉这次新建的两份（丢弃改动）：碰到的要是用户开着的实例，也不能留垃圾
                if not (isNull excelBook) then
                    (try
                        (excelBook :?> Microsoft.Office.Interop.Excel.Workbook).Close(box false)
                     with _ ->
                         ())

                if not (isNull wordDoc) then
                    (try
                        let mutable saveChanges: obj = box false
                        (wordDoc :?> Microsoft.Office.Interop.Word.Document).Close(&saveChanges)
                     with _ ->
                         ())

                track (
                    Proc.shutdown
                        "WINWORD"
                        20000
                        beforeWord
                        (fun () ->
                            let mutable sc: obj =
                                box Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges

                            let mutable fmt: obj = Type.Missing
                            let mutable route: obj = Type.Missing
                            word.Quit(&sc, &fmt, &route))
                )

                track (
                    Proc.shutdown
                        "EXCEL"
                        8000
                        beforeExcel
                        (fun () ->
                            (try
                                excel.DisplayAlerts <- excelAlerts
                             with _ ->
                                 ())

                            ExcelHost.Host.QuitWithoutSaving excel)
                ))

        // Word 必须自己退：它要卡只会卡在**看得见的**「是否保存」模态框上。
        // Excel 在这台机器上有个已知毛病 —— 客户端还活着时 Quit 后进程不肯退，
        // 但它是隐藏进程、不弹窗、不影响人，所以只记录不强判失败。
        Report.line
            output
            "现场"
            (sprintf
                "WINWORD %d 个 / EXCEL %d 个"
                (Report.procIds "WINWORD").Count
                (Report.procIds "EXCEL").Count)

        Report.line
            output
            "收尾"
            (if killed.Count = 0 then
                 "两家都是自己退的，没有对话框卡住"
             else
                 sprintf "%s 没自己退，已结束进程" (String.Join(" / ", killed)))

        Assert.False(
            killed.Contains "WINWORD",
            "Word 二十秒内没退出，被强制结束 —— 退出时多半又弹保存对话框了"
        )

        Report.conclusion output "通过：Word 说的，Excel 一字不差地收到了"

    [<Fact>]
    member this.``Excel 和 AutoCAD 打招呼：CAD 没开就不硬起，开了也只写自己新开的图``() =
        let beforeExcel = Report.procIds "EXCEL"
        let beforeAcad = Report.procIds "acad"
        let acadWasRunning = not beforeAcad.IsEmpty

        // 本机没有 AutoCAD 在跑时默认**不新起**：为了一条这次根本不用的通道弹出 CAD 窗口，
        // 是看不见的副作用。确实要测 AutoCAD 通道，跑之前设 XUNITHOST_LAUNCH_ACAD=1。
        let useAcad = acadWasRunning || Acad.mayLaunch ()

        Report.title output "Excel → AutoCAD：一句话传过去，只写自己新开的那张图"

        if not useAcad then
            Report.line output "前提" "本机没有 AutoCAD 在跑"
            Report.line output "处理" "不替你起一个（真要测这条通道，设 XUNITHOST_LAUNCH_ACAD=1 再跑）"
            Report.conclusion output "跳过：这条通道这次不测，机器上什么都没动"
            Assert.Skip "AutoCAD 没在跑；按约定不替用户新建实例"
        else
            let excelGreeting = "Hello AutoCAD. -- Excel"
            let acadGreeting = "AutoCAD 收到。Hello Excel. -- AutoCAD"

            let killed = ResizeArray<string>()

            let track (result: string option) =
                match result with
                | Some name -> killed.Add name
                | None -> ()

            Report.line
                output
                "AutoCAD"
                (if acadWasRunning then
                     "复用你开着的那只：另开一张临时图，写在那张上；你的图纸不碰"
                 else
                     "本次新起的一只")

            Sta.here (fun () ->
                let excel = ExcelHost.Host.GetOrCreateApplication ()

                // AutoCAD 可能中途崩掉重来，所以每次用都重新取，并留住最后一次拿到的那个供清理
                let mutable acad: Autodesk.AutoCAD.Interop.AcadApplication = null
                // 测试新开的那张临时图：收尾关掉，绝不留在用户实例里
                let mutable acadDoc: obj = null
                /// 这次自己新建的那本工作簿：收尾也关掉
                let mutable excelBook: obj = null

                let getAcad () =
                    let app = AutoCADHost.Host.GetOrCreateApplication ()
                    acad <- app
                    app

                let excelAlerts =
                    try
                        let original = excel.DisplayAlerts
                        excel.DisplayAlerts <- false
                        original
                    with _ ->
                        true

                try
                    // 1. Excel 新建一本、先认人、说一句
                    let excelId, heard, made = Say.excel excel "" excelGreeting
                    excelBook <- made
                    Report.line output "Excel" (Doc.describe excelId)
                    Report.line output "Excel 说" heard

                    // 2. AutoCAD 新开一张图、先认人、把听到的那句写成 TEXT
                    let acadId, lastWord, doc = Say.acad getAcad heard acadGreeting
                    acadDoc <- doc
                    Report.line output "AutoCAD" (Doc.describe acadId)
                    Report.line output "AutoCAD 说" lastWord

                    Assert.StartsWith(heard, lastWord)
                    Assert.Contains(acadGreeting, lastWord)
                    Report.line output "对照" "CAD 读回来的前一句，跟 Excel 说的逐字相同"
                finally
                    // 先把我们新建的工作簿关掉（丢弃改动）
                    if not (isNull excelBook) then
                        (try
                            (excelBook :?> Microsoft.Office.Interop.Excel.Workbook).Close(box false)
                         with _ ->
                             ())

                    // 再把我们新开的临时图关掉（不保存）：用户那只 acad 里不该多出一张图
                    if not (isNull acadDoc) then
                        (try
                            (acadDoc :?> Autodesk.AutoCAD.Interop.AcadDocument).Close(box false, Type.Missing)
                         with _ ->
                             ())

                    if not (isNull acad) then
                        // 关图纸 + 退出一律交给 AutoCADHost：它先用 ActiveDocument.Close(false) 丢掉改动，
                        // 再 Quit。少了这一步，带改动的 Drawing1 会让 Quit 弹出「是否保存」模态框，测试就卡死。
                        track (Proc.shutdown "acad" 20000 beforeAcad (fun () -> AutoCADHost.Host.QuitWithoutSaving acad))

                    track (
                        Proc.shutdown
                            "EXCEL"
                            8000
                            beforeExcel
                            (fun () ->
                                (try
                                    excel.DisplayAlerts <- excelAlerts
                                 with _ ->
                                     ())

                                ExcelHost.Host.QuitWithoutSaving excel)
                    ))

            Report.line
                output
                "现场"
                (sprintf "acad %d 个 / EXCEL %d 个" (Report.procIds "acad").Count (Report.procIds "EXCEL").Count)

            Report.line
                output
                "收尾"
                (if killed.Count = 0 then
                     "都是自己退的，没有对话框卡住"
                 else
                     sprintf "%s 没自己退，已结束进程" (String.Join(" / ", killed)))

            Assert.False(
                killed.Contains "acad",
                "AutoCAD 二十秒内没退出，被强制结束 —— 退出时多半又弹「是否保存 Drawing1」了"
            )

            Report.conclusion output "通过：Excel 那句话传到了 CAD，写在自己新开的图上，你的图纸一个字没动"
