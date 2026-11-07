/// .doc 升级成新格式：三个入口都在这一份文件里（照 ExcelHost/XlsUpgrade.fs）。
///   · `DocToDocx`   无宏 → .docx（格式码 12，正文逐字核对）
///   · `DocToDocm`   有宏 → .docm（格式码 13，正文 + 复核 VBA 宏还在）
///   · `DocUpgrade`  统一入口：判有没有宏，再选上面哪条路
/// 前两条各自独立：格式码、目标后缀、核对口径、结果文案各归各的模块；
/// 共用的开/关、正文指纹、判宏、临时文件在 DocIO.fs。
/// 本库只回传 (目标, 状态, 详情)，组装 Outcome 是测试项目的事——依赖方向不倒挂。
namespace WordHost

open System
open System.IO
open Microsoft.Office.Interop.Word

/// 转换之一：**无宏** .doc → .docx（wdFormatXMLDocument）。
/// 只接没宏的单；保宏那条路在下面的 DocToDocm。
module DocToDocx =

    /// wdFormatXMLDocument（无宏）
    let format: obj = box 12

    /// 这条转换认领的后缀
    let targetOf (source: string) = Path.ChangeExtension(source, ".docx")

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        let target = targetOf source
        let mutable doc: Document = null
        let mutable back: Document = null
        let mutable exists = false
        let mutable stage = target

        let body () =
            doc <- DocIO.openDoc app source false

            // 守卫：带宏的 .doc 存成 .docx 会把 VBA 悄悄扔掉，这单不接。
            // 分派本该把它送去 DocToDocm，这里再挡一道，防的是漏判。
            if DocIO.hasVba doc then
                (target, "失败", "这份 .doc 带 VBA 宏，不该存成 .docx；原件保留")
            else
                exists <- File.Exists target

                // 已有同名新文件：先转到临时文件，跟原件比完再删
                stage <- DocIO.stageFor exists target

                let before = DocIO.bodyText doc
                DocIO.saveDocAs format doc stage
                DocIO.closeDoc doc
                doc <- null

                back <- DocIO.openDoc app stage true
                let after = DocIO.bodyText back
                DocIO.closeDoc back
                back <- null

                if before <> after then
                    (target, "失败", sprintf "正文对不上（原件 %d 字 / 新文件 %d 字），原件保留" before.Length after.Length)
                elif exists then
                    (target, "已有·一致", "现有 docx 的正文与原件一致")
                else
                    (target, "转了", sprintf "正文 %d 字，一字不差" after.Length)

        // F# 不允许 try/with/finally 写在一起：结果先算出来，收尾单独做
        let outcome =
            try
                body ()
            with ex ->
                (target, "失败", OsHost.Errors.explain ex)

        if not (isNull doc) then DocIO.closeDoc doc
        if not (isNull back) then DocIO.closeDoc back

        DocIO.cleanupStage exists stage

        outcome

/// 转换之二：**带宏** .doc → .docm（wdFormatXMLDocumentMacroEnabled），VBA 原样保住。
/// 只管这一条路：格式码、.docm 后缀、正文核对、宏还在的复核、结果文案都在这里。
module DocToDocm =

    /// wdFormatXMLDocumentMacroEnabled（保留宏）
    let format: obj = box 13

    /// 这条转换认领的后缀
    let targetOf (source: string) = Path.ChangeExtension(source, ".docm")

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        let target = targetOf source
        let mutable doc: Document = null
        let mutable back: Document = null
        let mutable exists = false
        let mutable stage = target

        let body () =
            doc <- DocIO.openDoc app source false
            let hadMacro = DocIO.hasVba doc

            exists <- File.Exists target
            stage <- DocIO.stageFor exists target

            let before = DocIO.bodyText doc
            DocIO.saveDocAs format doc stage
            DocIO.closeDoc doc
            doc <- null

            back <- DocIO.openDoc app stage true
            let after = DocIO.bodyText back
            // 「转成 docm」的意义就在这儿：宏必须原样还在
            let macroKept = (not hadMacro) || DocIO.hasVba back
            DocIO.closeDoc back
            back <- null

            if before <> after then
                (target, "失败", sprintf "正文对不上（原件 %d 字 / 新文件 %d 字），原件保留" before.Length after.Length)
            elif not macroKept then
                (target, "失败", "新文件里 VBA 宏不见了，原件保留")
            elif exists then
                (target, "已有·一致", "现有 docm 的正文与原件一致，VBA 宏还在")
            else
                (target, "转了", sprintf "正文 %d 字，一字不差%s" after.Length (if hadMacro then "，VBA 宏原样保留" else ""))

        let outcome =
            try
                body ()
            with ex ->
                (target, "失败", OsHost.Errors.explain ex)

        if not (isNull doc) then DocIO.closeDoc doc
        if not (isNull back) then DocIO.closeDoc back

        DocIO.cleanupStage exists stage

        outcome

/// 统一入口：把一份 .doc 升级成新格式。
/// 判有没有 VBA 宏 → 有宏走 DocToDocm（.docm，保宏），没宏走 DocToDocx（.docx）；
/// 判不出来就报失败、原件保留，绝不猜「没宏」（猜错会悄悄丢宏）。
/// 判宏和选路都收在 WordHost 里，测试项目只负责批量和结果记录。
module DocUpgrade =

    /// 升级走哪条路，由有没有宏唯一决定（纯规则，不碰 COM，测试可直接验）
    type Route =
        | Docx
        | Docm

    let route (hasMacro: bool) : Route = if hasMacro then Docm else Docx

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        match DocIO.detectMacro app source with
        | Ok hasMacro ->
            match route hasMacro with
            | Docm -> DocToDocm.convert app source
            | Docx -> DocToDocx.convert app source
        | Error message ->
            (DocToDocx.targetOf source, "失败", "判不出有没有 VBA 宏，原件保留：" + message)
