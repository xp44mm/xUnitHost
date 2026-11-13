/// 「逐页复制」升级：把 .xls 的每张工作表一张一张抄进新工作簿，
/// 抄进**副本**之后再删掉新格式里不该带走的东西（图表 / 图片 / 控件 / OLE / 批注 /
/// 条件格式 / 数据验证 / 超链接 / 查询表 / 智能表 / 数据透视表），只留公式和单元格普通格式。
///
/// 为什么要这么绕：这些 .xls 里有老对象，直接 `SaveAs` 转 .xlsx 常常报错或把 Excel 弄崩；
/// 一张一张 `Worksheet.Copy` 抄过去，出问题能定位到具体那张表。
/// **原件一个字都不改**：只读打开、只在副本里删东西。
namespace ExcelHost

open System
open System.IO
open Microsoft.Office.Interop.Excel

module SheetCopy =

    /// Excel 互操作参数按值传 object
    let private missing: obj = Type.Missing

    /// 一张表的现场：名字、起点、尺寸、公式数、各类「不支持的格式」有多少
    type Facts =
        { Name: string
          Origin: string
          Rows: int
          Cols: int
          Formulas: int
          Shapes: int
          Charts: int
          Ole: int
          Comments: int
          CondFormats: int
          Validations: int
          Pivots: int
          Lists: int
          Queries: int
          Hyperlinks: int }

    /// 参数化属性一律返回 obj，取出来先转成强类型集合
    let private chartObjects (ws: Worksheet) = ws.ChartObjects(missing) :?> ChartObjects

    let private oleObjects (ws: Worksheet) = ws.OLEObjects(missing) :?> OLEObjects

    let private pivotTables (ws: Worksheet) = ws.PivotTables(missing) :?> PivotTables

    /// 特勤单元格的个数：范围内没有这类格子时 Excel 会抛，按 0 算
    let private specialCount (range: Range) (kind: XlCellType) =
        try
            range.SpecialCells(kind, missing).Count
        with _ ->
            0

    let private count (f: unit -> int) =
        try
            f ()
        with _ -> 0

    /// 现场清点（只读，不改任何东西）
    let factsOf (ws: Worksheet) : Facts =
        let used = ws.UsedRange

        { Name = ws.Name
          Origin = sprintf "R%dC%d" (count (fun () -> used.Row)) (count (fun () -> used.Column))
          Rows = count (fun () -> used.Rows.Count)
          Cols = count (fun () -> used.Columns.Count)
          Formulas = specialCount used XlCellType.xlCellTypeFormulas
          Shapes = count (fun () -> ws.Shapes.Count)
          Charts = count (fun () -> (chartObjects ws).Count)
          Ole = count (fun () -> (oleObjects ws).Count)
          Comments = specialCount used XlCellType.xlCellTypeComments
          CondFormats = count (fun () -> ws.Cells.FormatConditions.Count)
          Validations = specialCount used XlCellType.xlCellTypeAllValidation
          Pivots = count (fun () -> (pivotTables ws).Count)
          Lists = count (fun () -> ws.ListObjects.Count)
          Queries = count (fun () -> ws.QueryTables.Count)
          Hyperlinks = count (fun () -> ws.Hyperlinks.Count) }

    /// 倒着删集合里的每个对象（正着删会跳号）。返回 (项目, 删了几个)，-1 表示这项动不了
    let private deleteEach (label: string) (size: unit -> int) (remove: int -> unit) : string * int =
        let mutable doneCount = 0

        try
            for i in size () .. -1 .. 1 do
                try
                    remove i
                    doneCount <- doneCount + 1
                with _ ->
                    ()

            label, doneCount
        with _ ->
            label, -1

    /// 数据透视表没有 Delete：读成值再写回去，只剩静态值
    let private flattenPivots (ws: Worksheet) : string * int =
        let mutable doneCount = 0

        try
            let pivots = pivotTables ws

            for i in pivots.Count .. -1 .. 1 do
                try
                    let range = pivots.Item(i).TableRange2
                    let values = range.Value2
                    range.Clear() |> ignore

                    if not (isNull values) then
                        range.Value2 <- values

                    doneCount <- doneCount + 1
                with _ ->
                    ()

            "数据透视表→值", doneCount
        with _ ->
            "数据透视表→值", -1

    /// 智能表退回普通区域（保留公式和格式，只去掉表对象）
    let private unlistTables (ws: Worksheet) : string * int =
        let mutable doneCount = 0

        try
            let tables = ws.ListObjects

            for i in tables.Count .. -1 .. 1 do
                try
                    tables.Item(i).Unlist()
                    doneCount <- doneCount + 1
                with _ ->
                    ()

            "智能表→区域", doneCount
        with _ ->
            "智能表→区域", -1

    /// 删掉副本里不该带走的东西。返回逐项「删了多少」，0 的不列，-1 表示这项动不了。
    let purge (ws: Worksheet) : (string * int) list =
        let charts =
            deleteEach
                "图表"
                (fun () -> (chartObjects ws).Count)
                (fun i -> ((chartObjects ws).Item(i) :?> ChartObject).Delete() |> ignore)

        let ole =
            deleteEach
                "OLE对象"
                (fun () -> (oleObjects ws).Count)
                (fun i -> ((oleObjects ws).Item(i) :?> OLEObject).Delete() |> ignore)

        let shapes = deleteEach "图形图片控件" (fun () -> ws.Shapes.Count) (fun i -> ws.Shapes.Item(i).Delete() |> ignore)

        let comments =
            let n = specialCount ws.UsedRange XlCellType.xlCellTypeComments

            if n > 0 then
                try
                    ws.Cells.ClearComments()
                with _ ->
                    ()

            "批注", n

        let conds =
            let n = count (fun () -> ws.Cells.FormatConditions.Count)

            if n > 0 then
                try
                    ws.Cells.FormatConditions.Delete()
                with _ ->
                    ()

            "条件格式", n

        let validations =
            let n = specialCount ws.UsedRange XlCellType.xlCellTypeAllValidation

            if n > 0 then
                try
                    ws.Cells.Validation.Delete()
                with _ ->
                    ()

            "数据验证", n

        let hyperlinks =
            let n = count (fun () -> ws.Hyperlinks.Count)

            if n > 0 then
                try
                    ws.Hyperlinks.Delete()
                with _ ->
                    ()

            "超链接", n

        let queries = deleteEach "查询表" (fun () -> ws.QueryTables.Count) (fun i -> ws.QueryTables.Item(i).Delete() |> ignore)

        [ charts
          ole
          shapes
          comments
          conds
          validations
          hyperlinks
          queries
          unlistTables ws
          flattenPivots ws ]

    /// 只留「真动过的」项，拼成一行给报告用
    let describePurge (items: (string * int) list) =
        items
        |> List.filter (fun (_, n) -> n <> 0)
        |> List.map (fun (label, n) -> if n < 0 then label + "✗" else label + string n)
        |> String.concat " "

    /// 已经开着的那本就用它，别再 Open 一次（重复 Open 同一路径会把 Excel 拖进怪状态）
    let private findOpen (app: Application) (path: string) : Workbook =
        let mutable found: Workbook = null

        for i in 1 .. app.Workbooks.Count do
            let book = app.Workbooks.Item(i)

            if not (isNull book) && String.Equals(book.FullName, path, StringComparison.OrdinalIgnoreCase) then
                found <- book

        found

    /// 整本保存：把 .xls 拷到临时副本 → 在**副本**上清掉带不走的东西 → 整本 `SaveAs` 成 .xlsx。
    /// 比逐页抄少丢东西（图表工作表、页设置、打印区域都还在），但老文件里若有硬伤，
    /// 整本另存可能失败——失败了就回去用 `convert` 逐页抄。
    /// 原件一个字不改：只读复制出去，动的全是临时副本。
    let wholeSave (app: Application) (source: string) (target: string) : string * (string * string) list * string * string =
        let temp =
            Path.Combine(Path.GetTempPath(), "purgexls-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xls")

        let mutable book: Workbook = null
        let lines = ResizeArray<string * string>()

        let body () =
            File.Copy(source, temp, true)
            book <- BookIO.openBook app temp false

            for i in 1 .. book.Worksheets.Count do
                let ws = book.Worksheets.Item(i) :?> Worksheet
                let facts = factsOf ws
                let purged = purge ws

                let note =
                    sprintf
                        "%s 行%d 列%d 公式%d%s"
                        facts.Origin
                        facts.Rows
                        facts.Cols
                        facts.Formulas
                        (match describePurge purged with
                         | "" -> ""
                         | text -> "　清掉：" + text)

                lines.Add(ws.Name, note)

            book.SaveAs(box target, box 51) |> ignore
            book.Close(box false) |> ignore
            book <- null

        let status, detail =
            try
                body ()

                let chartSheets =
                    try
                        "转了", sprintf "%d 张工作表（整本另存）" lines.Count
                    with _ ->
                        "转了", "整本另存"

                chartSheets
            with ex ->
                "失败", OsHost.Errors.explain ex

        if not (isNull book) then BookIO.closeBook book

        (try
            if File.Exists temp then File.Delete temp
         with _ ->
             ())

        target, List.ofSeq lines, status, detail

    /// 逐页复制：每张表 `Copy` 进新簿 → 在新簿那张表上 `purge` → 全部抄完存 .xlsx。
    /// 返回 (真正落盘的目标, 每张表一行说明, 状态, 详情)
    let convert (app: Application) (source: string) (target: string) : string * (string * string) list * string * string =
        let mutable src: Workbook = null
        let mutable out: Workbook = null
        let mutable ownSrc = false
        let mutable chartSheets = 0
        let lines = ResizeArray<string * string>()

        let body () =
            let already = findOpen app source

            if isNull already then
                src <- BookIO.openBook app source true
                ownSrc <- true
            else
                src <- already

            for i in 1 .. src.Sheets.Count do
                if (src.Sheets.Item(i) :? Chart) then
                    chartSheets <- chartSheets + 1

            out <- app.Workbooks.Add()
            let placeholder = out.Worksheets.Item(1) :?> Worksheet
            placeholder.Name <- "__drop__"

            for i in 1 .. src.Worksheets.Count do
                let ws = src.Worksheets.Item(i) :?> Worksheet
                let facts = factsOf ws

                // 抄到最后一张后面
                ws.Copy(missing, box (out.Sheets.Item(out.Sheets.Count)))

                let copy = out.Sheets.Item(out.Sheets.Count) :?> Worksheet
                let purged = purge copy

                let note =
                    sprintf
                        "%s 行%d 列%d 公式%d%s"
                        facts.Origin
                        facts.Rows
                        facts.Cols
                        facts.Formulas
                        (match describePurge purged with
                         | "" -> ""
                         | text -> "　清掉：" + text)

                lines.Add(copy.Name, note)

            placeholder.Delete()
            out.SaveAs(box target, box 51) |> ignore
            out.Close(box false) |> ignore
            out <- null

        let status, detail =
            try
                body ()

                if chartSheets > 0 then
                    "转了", sprintf "%d 张工作表（另有 %d 张图表工作表没抄）" lines.Count chartSheets
                else
                    "转了", sprintf "%d 张工作表" lines.Count
            with ex ->
                "失败", OsHost.Errors.explain ex

        if not (isNull out) then BookIO.closeBook out
        if ownSrc && not (isNull src) then BookIO.closeBook src

        target, List.ofSeq lines, status, detail
