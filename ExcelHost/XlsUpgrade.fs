/// .xls 升级成新格式：三个入口都在这一份文件里。
///   · `XlsToXlsx`  无宏 → .xlsx（格式码 51，格子值指纹核对）
///   · `XlsToXlsm`  有宏 → .xlsm（格式码 52，格子值 + 复核 VBA 宏还在）
///   · `XlsUpgrade` 统一入口：判有没有宏，再选上面哪条路
/// 前两条各自独立：格式码、目标后缀、核对口径、结果文案各归各的模块；
/// 共用的开/关、指纹、判宏、临时文件在 BookIO.fs。
/// 本库只回传 (目标, 状态, 详情)，组装 Outcome 是测试项目的事——依赖方向不倒挂。
namespace ExcelHost

open System
open System.IO
open Microsoft.Office.Interop.Excel

/// 转换之二：**无宏** .xls → .xlsx（xlOpenXMLWorkbook）。
/// 只接没宏的单；保宏那条路在下面的 XlsToXlsm。
module XlsToXlsx =

    /// xlOpenXMLWorkbook（无宏）
    let format: obj = box 51

    /// 这条转换认领的后缀
    let targetOf (source: string) = Path.ChangeExtension(source, ".xlsx")

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        let target = targetOf source
        let mutable book: Workbook = null
        let mutable back: Workbook = null
        let mutable exists = false
        let mutable stage = target

        let body () =
            book <- BookIO.openBook app source false

            // 守卫：带宏的 .xls 存成 .xlsx 会把 VBA 悄悄扔掉，这单不接。
            // 分派本该把它送去 XlsToXlsm，这里再挡一道，防的是漏判。
            if BookIO.hasVba book then
                (target, "失败", "这份 .xls 带 VBA 宏，不该存成 .xlsx；原件保留")
            else
                exists <- File.Exists target

                // 已有同名新文件：先转到临时文件，跟原件比完再删
                stage <- BookIO.stageFor exists target

                let before = BookIO.bookPrint book
                let sheets = book.Worksheets.Count

                book.SaveAs(box stage, format) |> ignore
                BookIO.closeBook book
                book <- null

                back <- BookIO.openBook app stage true
                let after = BookIO.bookPrint back
                BookIO.closeBook back
                back <- null

                if before <> after then
                    (target, "失败", sprintf "内容对不上（原件 %d 字符指纹 / 新文件 %d），原件保留" before.Length after.Length)
                elif exists then
                    (target, "已有·一致", sprintf "现有 %s 的内容与原件一致" (Path.GetExtension target))
                else
                    (target, "转了", sprintf "%d 张表，内容一致" sheets)

        // 同 Word：结果先算出来，收尾单独做
        let outcome =
            try
                body ()
            with ex ->
                (target, "失败", OsHost.Errors.explain ex)

        if not (isNull book) then BookIO.closeBook book
        if not (isNull back) then BookIO.closeBook back

        BookIO.cleanupStage exists stage

        outcome

/// 转换之三：**带宏** .xls → .xlsm（xlOpenXMLWorkbookMacroEnabled），VBA 原样保住。
/// 只管这一条路：格式码、.xlsm 后缀、格子值核对、宏还在的复核、结果文案都在这里。
module XlsToXlsm =

    /// xlOpenXMLWorkbookMacroEnabled（保留宏）
    let format: obj = box 52

    /// 这条转换认领的后缀
    let targetOf (source: string) = Path.ChangeExtension(source, ".xlsm")

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        let target = targetOf source
        let mutable book: Workbook = null
        let mutable back: Workbook = null
        let mutable exists = false
        let mutable stage = target

        let body () =
            book <- BookIO.openBook app source false
            let hadMacro = BookIO.hasVba book

            exists <- File.Exists target

            // 已有同名 xlsm：先转到临时文件，跟原件比完再删
            stage <- BookIO.stageFor exists target

            let before = BookIO.bookPrint book
            let sheets = book.Worksheets.Count

            book.SaveAs(box stage, format) |> ignore
            BookIO.closeBook book
            book <- null

            back <- BookIO.openBook app stage true
            let after = BookIO.bookPrint back
            // 「转成 xlsm」的意义就在这儿：宏必须原样还在
            let macroKept = (not hadMacro) || BookIO.hasVba back
            BookIO.closeBook back
            back <- null

            if before <> after then
                (target, "失败", sprintf "内容对不上（原件 %d 字符指纹 / 新文件 %d），原件保留" before.Length after.Length)
            elif not macroKept then
                (target, "失败", "新文件里 VBA 宏不见了，原件保留")
            elif exists then
                (target, "已有·一致", sprintf "现有 %s 的内容与原件一致，VBA 宏还在" (Path.GetExtension target))
            else
                (target, "转了", sprintf "%d 张表，内容一致%s" sheets (if hadMacro then "，VBA 宏原样保留" else ""))

        let outcome =
            try
                body ()
            with ex ->
                (target, "失败", OsHost.Errors.explain ex)

        if not (isNull book) then BookIO.closeBook book
        if not (isNull back) then BookIO.closeBook back

        BookIO.cleanupStage exists stage

        outcome

/// 统一入口：把一份 .xls 升级成新格式。
/// 判有没有 VBA 宏 → 有宏走 XlsToXlsm（.xlsm，保宏），没宏走 XlsToXlsx（.xlsx）；
/// 判不出来就报失败、原件保留，绝不猜「没宏」（猜错会悄悄丢宏）。
/// 判宏和选路都收在 ExcelHost 里，测试项目只负责批量和结果记录。
module XlsUpgrade =

    /// 升级走哪条路，由有没有宏唯一决定（纯规则，不碰 COM，测试可直接验）
    type Route =
        | Xlsx
        | Xlsm

    let route (hasMacro: bool) : Route = if hasMacro then Xlsm else Xlsx

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        match BookIO.detectMacro app source with
        | Ok hasMacro ->
            match route hasMacro with
            | Xlsm -> XlsToXlsm.convert app source
            | Xlsx -> XlsToXlsx.convert app source
        | Error message ->
            (XlsToXlsx.targetOf source, "失败", "判不出有没有 VBA 宏，原件保留：" + message)
