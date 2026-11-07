/// .ppt 升级成新格式：三个入口都在这一份文件里（照 WordHost/DocUpgrade.fs、ExcelHost/XlsUpgrade.fs）。
///   · `PptToPptx`   无宏 → .pptx（格式码 24，幻灯片文字逐字核对）
///   · `PptToPptm`   有宏 → .pptm（格式码 25，幻灯片文字 + 复核 VBA 宏还在）
///   · `PptUpgrade`  统一入口：判有没有宏，再选上面哪条路；对外函数就叫 `pptUpgrade`
/// 前两条各自独立：格式码、目标后缀、核对口径、结果文案各归各的模块；
/// 共用的开/关、指纹、判宏、临时文件在 PptIO.fs。
/// 本库只回传 (目标, 状态, 详情)，组装 Outcome 是测试项目的事——依赖方向不倒挂。
namespace PowerPointHost

open System
open System.IO
open Microsoft.Office.Interop.PowerPoint

/// 转换之一：**无宏** .ppt → .pptx（ppSaveAsOpenXMLPresentation = 24）。
/// 只接没宏的单；保宏那条路在下面的 PptToPptm。
module PptToPptx =

    /// 强类型枚举，值 24；PowerPoint 的 SaveAs 收的就是它
    let format = PpSaveAsFileType.ppSaveAsOpenXMLPresentation

    /// 这条转换认领的后缀
    let targetOf (source: string) = Path.ChangeExtension(source, ".pptx")

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        let target = targetOf source
        let mutable pres: Presentation = null
        let mutable back: Presentation = null
        let mutable exists = false
        let mutable stage = target

        let body () =
            pres <- PptIO.openPpt app source false

            // 守卫：带宏的 .ppt 存成 .pptx 会把 VBA 悄悄扔掉，这单不接。
            // 分派本该把它送去 PptToPptm，这里再挡一道，防的是漏判。
            if PptIO.hasVba pres then
                (target, "失败", "这份 .ppt 带 VBA 宏，不该存成 .pptx；原件保留")
            else
                exists <- File.Exists target

                // 已有同名新文件：先转到临时文件，跟原件比完再删
                stage <- PptIO.stageFor exists target

                let before = PptIO.slidePrint pres
                PptIO.savePptAs format pres stage
                PptIO.closePpt pres
                pres <- null

                back <- PptIO.openPpt app stage true
                let after = PptIO.slidePrint back
                PptIO.closePpt back
                back <- null

                if before <> after then
                    (target, "失败", sprintf "幻灯片文字对不上（原件 %d 字符指纹 / 新文件 %d），原件保留" before.Length after.Length)
                elif exists then
                    (target, "已有·一致", "现有 pptx 的幻灯片文字与原件一致")
                else
                    (target, "转了", sprintf "幻灯片文字 %d 字符，一字不差" after.Length)

        // F# 不允许 try/with/finally 写在一起：结果先算出来，收尾单独做
        let outcome =
            try
                body ()
            with ex ->
                (target, "失败", OsHost.Errors.explain ex)

        if not (isNull pres) then PptIO.closePpt pres
        if not (isNull back) then PptIO.closePpt back

        PptIO.cleanupStage exists stage

        outcome

/// 转换之二：**带宏** .ppt → .pptm（ppSaveAsOpenXMLPresentationMacroEnabled = 25），VBA 原样保住。
/// 只管这一条路：格式码、.pptm 后缀、文字核对、宏还在的复核、结果文案都在这里。
module PptToPptm =

    /// 强类型枚举，值 25
    let format = PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled

    /// 这条转换认领的后缀
    let targetOf (source: string) = Path.ChangeExtension(source, ".pptm")

    /// 返回 (真正落盘的目标, 状态, 详情)
    let convert (app: Application) (source: string) : string * string * string =
        let target = targetOf source
        let mutable pres: Presentation = null
        let mutable back: Presentation = null
        let mutable exists = false
        let mutable stage = target

        let body () =
            pres <- PptIO.openPpt app source false
            let hadMacro = PptIO.hasVba pres

            exists <- File.Exists target
            stage <- PptIO.stageFor exists target

            let before = PptIO.slidePrint pres
            PptIO.savePptAs format pres stage
            PptIO.closePpt pres
            pres <- null

            back <- PptIO.openPpt app stage true
            let after = PptIO.slidePrint back
            // 「转成 pptm」的意义就在这儿：宏必须原样还在
            let macroKept = (not hadMacro) || PptIO.hasVba back
            PptIO.closePpt back
            back <- null

            if before <> after then
                (target, "失败", sprintf "幻灯片文字对不上（原件 %d 字符指纹 / 新文件 %d），原件保留" before.Length after.Length)
            elif not macroKept then
                (target, "失败", "新文件里 VBA 宏不见了，原件保留")
            elif exists then
                (target, "已有·一致", "现有 pptm 的幻灯片文字与原件一致，VBA 宏还在")
            else
                (target, "转了", sprintf "幻灯片文字 %d 字符，一字不差%s" after.Length (if hadMacro then "，VBA 宏原样保留" else ""))

        let outcome =
            try
                body ()
            with ex ->
                (target, "失败", OsHost.Errors.explain ex)

        if not (isNull pres) then PptIO.closePpt pres
        if not (isNull back) then PptIO.closePpt back

        PptIO.cleanupStage exists stage

        outcome

/// 统一入口：把一份 .ppt 升级成新格式。
/// 判有没有 VBA 宏 → 有宏走 PptToPptm（.pptm，保宏），没宏走 PptToPptx（.pptx）；
/// 判不出来就报失败、原件保留，绝不猜「没宏」（猜错会悄悄丢宏）。
/// 判宏和选路都收在 PowerPointHost 里，测试项目只负责批量和结果记录。
module PptUpgrade =

    /// 升级走哪条路，由有没有宏唯一决定（纯规则，不碰 COM，测试可直接验）
    type Route =
        | Pptx
        | Pptm

    let route (hasMacro: bool) : Route = if hasMacro then Pptm else Pptx

    /// **pptUpgrade**：一份 .ppt → 新格式。返回 (真正落盘的目标, 状态, 详情)。
    /// 调用方负责给它一个 PowerPoint Application（一般用 `Host.execQuery` 起的那只）。
    let pptUpgrade (app: Application) (source: string) : string * string * string =
        match PptIO.detectMacro app source with
        | Ok hasMacro ->
            match route hasMacro with
            | Pptm -> PptToPptm.convert app source
            | Pptx -> PptToPptx.convert app source
        | Error message ->
            (PptToPptx.targetOf source, "失败", "判不出有没有 VBA 宏，原件保留：" + message)

    /// 与 `DocUpgrade.convert` / `XlsUpgrade.convert` 对齐的别名：批量分派可以直接照旧调 `convert`
    let convert = pptUpgrade
