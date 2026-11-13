/// 「逐页复制」升级 .xls：一张一张抄进新工作簿，抄完在副本里清掉新格式带不走的东西，
/// 只留公式和单元格普通格式。原件只读、一个字不改。
///
/// 真机闸门（用户开着 Excel、那份 .xls 就是活动工作簿）：
///   XUNITHOST_EXCEL_PID=12345 \
///   XUNITHOST_XLS_PAGE='C:\…\脱硫系统材料清单.xls' \
///   xUnitHost.exe -method "*逐页复制*" -noColor -showLiveOutput
///
/// 只想看看里面有什么、不动手：再加 XUNITHOST_XLS_PAGE_DRY=1
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open OsHost
open Microsoft.Office.Interop.Excel

type XlsPageCopyTest(output: ITestOutputHelper) =

    let pid () =
        match Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID" with
        | null
        | "" -> (Assert.Skip "设 XUNITHOST_EXCEL_PID=<Excel 的进程号>"; 0)
        | text ->
            match Int32.TryParse text with
            | true, value when value > 0 -> value
            | _ -> (Assert.Skip "XUNITHOST_EXCEL_PID 不是正整数"; 0)

    let source () =
        match Environment.GetEnvironmentVariable "XUNITHOST_XLS_PAGE" with
        | null
        | "" -> (Assert.Skip "设 XUNITHOST_XLS_PAGE=<那份 .xls 的完整路径>"; "")
        | text -> if File.Exists text then text else (Assert.Skip(sprintf "文件不在：%s" text); "")

    /// 接上那只 Excel（按 PID，绝不新建、绝不 Quit）
    let attach (value: int) =
        let work = ExcelHost.ExcelTask.GetApplicationAsync value
        work.GetAwaiter().GetResult()

    /// 顺手设置（关提示、关屏幕刷新）：Excel 忙着的时候会拒，拒了就当没设过
    let mute (app: Application) =
        try
            app.DisplayAlerts <- false
        with _ ->
            ()

        try
            app.ScreenUpdating <- false
        with _ ->
            ()

    let unmute (app: Application) (alerts: bool) =
        try
            app.ScreenUpdating <- true
        with _ ->
            ()

        try
            app.DisplayAlerts <- alerts
        with _ ->
            ()

    /// 找到已经开着的那本（没开就自己只读开一份，用完关掉）
    let withSource (app: Application) (path: string) (work: Workbook -> 'T) : 'T =
        let mutable found: Workbook = null

        for i in 1 .. app.Workbooks.Count do
            let candidate = app.Workbooks.Item(i)

            if not (isNull candidate) && String.Equals(candidate.FullName, path, StringComparison.OrdinalIgnoreCase) then
                found <- candidate

        if not (isNull found) then
            work found
        else
            let opened = ExcelHost.BookIO.openBook app path true

            try
                work opened
            finally
                ExcelHost.BookIO.closeBook opened

    [<Fact>]
    member this.``逐页复制：先清点这份 xls 里有哪些不支持的格式``() =
        let pidValue = pid ()
        let src = source ()

        Report.title output "逐页复制　清点（只读，不动文件）"
        Report.line output "Excel PID" (string pidValue)

        let stamp = (File.GetLastWriteTime src).ToString("yy-MM-dd HH:mm")
        let kb = FileInfo(src).Length / 1024L
        Report.line output "源文件" (sprintf "%s（%s，%d KB）" (Path.GetFileName src) stamp kb)

        let app = attach pidValue

        let facts, chartSheets =
            STA.run(fun () ->
                let alerts = app.DisplayAlerts
                app.DisplayAlerts <- false

                try
                    withSource app src (fun book ->
                        let facts =
                            [ for i in 1 .. book.Worksheets.Count do
                                  ExcelHost.SheetCopy.factsOf (book.Worksheets.Item(i) :?> Worksheet) ]

                        let charts =
                            [ for i in 1 .. book.Sheets.Count do
                                  let sheet = book.Sheets.Item(i)

                                  if (sheet :? Chart) then
                                      (sheet :?> Chart).Name ]

                        facts, charts)
                finally
                    app.DisplayAlerts <- alerts)

        Report.line output "工作表" (sprintf "%d 张（另图表工作表 %d 张）" facts.Length chartSheets.Length)

        for i in 0 .. facts.Length - 1 do
            let f = facts.[i]

            Report.line2
                output
                (sprintf "%02d" (i + 1))
                f.Name
                (sprintf
                    "%s 行%d 列%d 公式%d ｜ 图形%d 图表%d OLE%d 批注%d 条件格式%d 验证%d 透视%d 智能表%d 查询%d 链接%d"
                    f.Origin
                    f.Rows
                    f.Cols
                    f.Formulas
                    f.Shapes
                    f.Charts
                    f.Ole
                    f.Comments
                    f.CondFormats
                    f.Validations
                    f.Pivots
                    f.Lists
                    f.Queries
                    f.Hyperlinks)

        let total (pick: ExcelHost.SheetCopy.Facts -> int) = facts |> List.sumBy pick

        Report.line
            output
            "合计"
            (sprintf
                "公式 %d ｜ 图形 %d 图表 %d OLE %d 批注 %d 条件格式 %d 验证 %d 透视 %d 智能表 %d 查询 %d 链接 %d"
                (total (fun f -> f.Formulas))
                (total (fun f -> f.Shapes))
                (total (fun f -> f.Charts))
                (total (fun f -> f.Ole))
                (total (fun f -> f.Comments))
                (total (fun f -> f.CondFormats))
                (total (fun f -> f.Validations))
                (total (fun f -> f.Pivots))
                (total (fun f -> f.Lists))
                (total (fun f -> f.Queries))
                (total (fun f -> f.Hyperlinks)))

        Should.equal true (facts.Length > 0)
        Report.conclusion output "完成：只清点，没有落任何文件"

    [<Fact>]
    member this.``逐页复制：做成新 xlsx，只留公式和普通格式``() =
        let pidValue = pid ()
        let src = source ()

        Report.title output "逐页复制　一张一张抄表 → 副本里清掉带不走的东西 → 存 xlsx"

        let target =
            let plain = Path.ChangeExtension(src, ".xlsx")

            if File.Exists plain then
                Path.Combine(Path.GetDirectoryName plain, Path.GetFileNameWithoutExtension plain + "（逐页复制）.xlsx")
            else
                plain

        let srcKb = FileInfo(src).Length / 1024L
        Report.line output "源文件" (sprintf "%s（%d KB，只读，不改）" (Path.GetFileName src) srcKb)
        Report.line output "新文件" target

        let app = attach pidValue

        let before, lines, status, detail =
            STA.run(fun () ->
                let alerts = app.DisplayAlerts
                app.DisplayAlerts <- false
                app.ScreenUpdating <- false

                try
                    // 先把原件的内容指纹拍下来，抄完跟新文件比
                    let print = withSource app src (fun book -> ExcelHost.BookIO.bookPrint book)
                    let targetPath, lines, status, detail = ExcelHost.SheetCopy.convert app src target
                    print, lines, status, detail, targetPath
                finally
                    app.ScreenUpdating <- true
                    app.DisplayAlerts <- alerts)
            |> fun (print, lines, status, detail, _) -> print, lines, status, detail

        for i in 0 .. lines.Length - 1 do
            let name, note = lines.[i]
            Report.line2 output (sprintf "%02d" (i + 1)) name note

        Report.line output "结果" (sprintf "%s ｜ %s" status detail)

        if status = "转了" then
            let after = STA.run(fun () -> ExcelHost.BookIO.readPrint app target)
            Report.line output "指纹" (sprintf "原件 %d 字符 ／ 新文件 %d 字符" before.Length after.Length)
            Should.equal before after
            Report.conclusion output (sprintf "通过：%d 张表逐页抄完，格子值与原文件一致，只留公式和普通格式 → %s" lines.Length target)
        else
            Report.conclusion output ("失败：" + detail)
            Should.equal "转了" status

    [<Fact>]
    member this.``整本保存：删掉 OLE 后整个工作簿直接另存为 xlsx``() =
        let pidValue = pid ()
        let src = source ()

        Report.title output "整本保存　拷临时副本 → 副本上删掉带不走的东西 → 整本另存 xlsx"

        let target =
            Path.Combine(Path.GetDirectoryName src, Path.GetFileNameWithoutExtension src + "（整本）.xlsx")

        let srcKb = FileInfo(src).Length / 1024L
        Report.line output "源文件" (sprintf "%s（%d KB，只读，不改）" (Path.GetFileName src) srcKb)
        Report.line output "新文件" target

        let app = attach pidValue

        let before, lines, status, detail =
            STA.run(fun () ->
                let alerts = try app.DisplayAlerts with _ -> true
                mute app

                try
                    let print = withSource app src (fun book -> ExcelHost.BookIO.bookPrint book)
                    let ignored, lines, status, detail = ExcelHost.SheetCopy.wholeSave app src target
                    ignore ignored
                    print, lines, status, detail
                finally
                    unmute app alerts)

        for i in 0 .. lines.Length - 1 do
            let name, note = lines.[i]
            Report.line2 output (sprintf "%02d" (i + 1)) name note

        Report.line output "结果" (sprintf "%s ｜ %s" status detail)

        if status = "转了" then
            let after = STA.run(fun () -> ExcelHost.BookIO.readPrint app target)
            Report.line output "指纹" (sprintf "原件 %d 字符 ／ 新文件 %d 字符" before.Length after.Length)
            Should.equal before after

            let facts =
                STA.run(fun () ->
                    withSource app target (fun book ->
                        [ for i in 1 .. book.Worksheets.Count do
                              ExcelHost.SheetCopy.factsOf (book.Worksheets.Item(i) :?> Worksheet) ]))

            Report.line
                output
                "回读"
                (sprintf
                    "%d 张表，公式 %d，图形 %d 图表 %d OLE %d"
                    facts.Length
                    (facts |> List.sumBy (fun f -> f.Formulas))
                    (facts |> List.sumBy (fun f -> f.Shapes))
                    (facts |> List.sumBy (fun f -> f.Charts))
                    (facts |> List.sumBy (fun f -> f.Ole)))

            Report.conclusion output (sprintf "通过：整本另存 %d 张表，格子值一致 → %s" lines.Length target)
        else
            Report.conclusion output ("失败，这一份得回去用逐页复制：" + detail)
            Should.equal "转了" status
