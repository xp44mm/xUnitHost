/// Excel 转换共用的 COM 细节：开/关工作簿、内容指纹、判宏、临时文件、只读读回。
/// 放在 ExcelHost 里，因为这些全是 Excel 自己的活：XlsToXlsx / XlsToXlsm 两条转换，
/// 以及测试项目的事后核对，都走这一份，不重复实现。
namespace ExcelHost

open System
open System.Globalization
open System.IO
open System.Text
open Microsoft.Office.Interop.Excel

module BookIO =

    /// Excel 的互操作参数是**按值**传 object（Word 才是 obj&）。
    /// UpdateLinks=0 是 xlUpdateLinksNever：别弹「要不要更新链接」。
    /// CorruptLoad=xlNormalLoad（第 15 个参数）：老 .xls 被判损坏时不进修复模式、不弹修复对话框。
    let private missing: obj = Type.Missing

    let openBook (app: Application) (path: string) (readOnly: bool) : Workbook =
        app.Workbooks.Open(
            path,
            box 0,
            box readOnly,
            missing,
            missing,
            missing,
            missing,
            missing,
            missing,
            missing,
            missing,
            missing,
            missing,
            missing,
            box(int XlCorruptLoad.xlNormalLoad)
        )

    let closeBook (book: Workbook) =
        try
            book.Close(box false)
        with _ ->
            ()

    /// 一本工作簿的内容指纹：每张有内容的表的**表名 + 格子的值**（行优先）。
    ///
    /// 特意**不比** `UsedRange` 的行列数：.xls 会把「只有格式的空列」也算进 UsedRange，
    /// 存成 .xlsx 之后又算不上，比尺寸会假报警。只有格式没有值的空表（.xls 常见的
    /// 多余 Sheet2/Sheet3）也不算内容，跳过不记。
    let bookPrint (book: Workbook) : string =
        let sb = StringBuilder()

        for i in 1..book.Worksheets.Count do
            let sheet = book.Worksheets.Item(i) :?> Worksheet
            let values = StringBuilder()
            let used = sheet.UsedRange
            let cell = used.Value2

            if not (isNull cell) then
                match cell with
                | :? (obj[,]) as grid ->
                    // 坑：Excel 返回的二维数组下界是 1（不是 0），按 0 起算必越界
                    for r in grid.GetLowerBound 0..grid.GetUpperBound 0 do
                        for c in grid.GetLowerBound 1..grid.GetUpperBound 1 do
                            let v = grid.[r, c]

                            if not (isNull v) then
                                values.Append(System.Convert.ToString(v, CultureInfo.InvariantCulture)).Append('\u001F')
                                |> ignore
                | single -> values.Append(System.Convert.ToString(single, CultureInfo.InvariantCulture)).Append('\u001F') |> ignore

            if values.Length > 0 then
                sb.Append(sheet.Name).Append('|').Append(values.ToString()) |> ignore

        sb.ToString()

    /// 读 VBA 工程在不在。读不出来（信任设置没开）就当「没有」，调用方自己决定怎么办。
    let hasVba (book: Workbook) : bool =
        try
            book.HasVBProject
        with _ ->
            false

    /// 已有同名新文件时先转到临时文件，跟原件比完再删；没有就直接落在目标上。
    let stageFor (exists: bool) (target: string) : string =
        if exists then
            Path.Combine(
                Path.GetTempPath(),
                "docconvert-" + Guid.NewGuid().ToString("N").Substring(0, 8) + Path.GetExtension target
            )
        else
            target

    let cleanupStage (exists: bool) (stage: string) =
        if exists && File.Exists stage then
            (try
                File.Delete stage
             with _ ->
                 ())

    /// 只读读回内容指纹：给事后核对用
    let readPrint (app: Application) (path: string) : string =
        let book = openBook app path true
        try bookPrint book finally closeBook book

    /// 分派用：只读开一下 .xls，判有没有 VBA 宏。
    /// **读不出 VBA 工程信息就报错**（原件保留），绝不猜「没宏」——猜错会把宏悄悄丢掉。
    let detectMacro (app: Application) (source: string) : Result<bool, string> =
        let mutable book: Workbook = null

        let result =
            try
                book <- openBook app source true
                Ok(book.HasVBProject)
            with ex ->
                Error(OsHost.Errors.explain ex)

        if not (isNull book) then closeBook book
        result
