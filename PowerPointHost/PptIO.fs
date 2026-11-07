/// PowerPoint 转换共用的 COM 细节：开/关演示文稿、另存、幻灯片文字指纹、判宏、临时文件、只读读回。
/// 放在 PowerPointHost 里，因为这些全是 PowerPoint 自己的活：PptToPptx / PptToPptm 两条转换，
/// 以及测试项目的事后核对，都走这一份，不重复实现。
/// 对应 WordHost/DocIO.fs、ExcelHost/BookIO.fs。
namespace PowerPointHost

open System
open System.IO
open System.Text
open Microsoft.Office.Core
open Microsoft.Office.Interop.PowerPoint

module PptIO =

    /// PowerPoint 的 Open / SaveAs 参数是**强类型按值**（Word 是 obj&，Excel 是 object）。
    /// WithWindow=msoFalse：不弹演示文稿窗口，自动化时不抢用户桌面。
    /// Untitled=msoFalse：按文件路径正常打开，不当副本。
    let openPpt (app: Application) (path: string) (readOnly: bool) : Presentation =
        let ro = if readOnly then MsoTriState.msoTrue else MsoTriState.msoFalse
        app.Presentations.Open(path, ro, MsoTriState.msoFalse, MsoTriState.msoFalse)

    /// 另存成新格式；格式码由各条转换自己带，这里不认得 .pptx / .pptm 谁是谁。
    let savePptAs (fmt: PpSaveAsFileType) (pres: Presentation) (target: string) =
        pres.SaveAs(target, fmt, MsoTriState.msoTriStateMixed)

    let closePpt (pres: Presentation) =
        try
            pres.Close()
        with _ ->
            ()

    /// 文本框里的换行在 PowerPoint 里是 \r / 垂直制表符，比较口径统一去掉。
    let normalize (text: string) : string =
        if isNull text then
            ""
        else
            text.Replace("\u000D", "").Replace("\u000B", "").Replace("\u0007", "").Trim()

    /// 一份演示文稿的内容指纹：每张幻灯片上有文字的形状，按「页序 / 形状序」拼起来。
    /// 特意**不比**版式、主题、形状数量 —— .ppt 与 .pptx 之间这些会变，文字才是内容。
    let slidePrint (pres: Presentation) : string =
        let sb = StringBuilder()

        for i in 1..pres.Slides.Count do
            let slide = pres.Slides.Item(i)
            sb.Append(i).Append('|') |> ignore

            for j in 1..slide.Shapes.Count do
                let shape = slide.Shapes.Item(j)
                let text = try normalize shape.TextFrame.TextRange.Text with _ -> ""

                if text <> "" then
                    sb.Append(j).Append(':').Append(text).Append('\u001F') |> ignore

            sb.Append('\u001E') |> ignore

        sb.ToString()

    /// 读 VBA 工程在不在。读不出来（信任设置没开）就当「没有」，调用方自己决定怎么办。
    let hasVba (pres: Presentation) : bool =
        try
            pres.HasVBProject
        with _ ->
            false

    /// 已有同名新文件时先转到临时文件，跟原件比完再删；没有就直接落在目标上。
    let stageFor (exists: bool) (target: string) : string =
        if exists then
            Path.Combine(
                Path.GetTempPath(),
                "pptconvert-" + Guid.NewGuid().ToString("N").Substring(0, 8) + Path.GetExtension target
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
        let pres = openPpt app path true
        try slidePrint pres finally closePpt pres

    /// 分派用：只读开一下 .ppt，判有没有 VBA 宏。
    /// **读不出 VBA 工程信息就报错**（原件保留），绝不猜「没宏」——猜错会把宏悄悄丢掉。
    let detectMacro (app: Application) (source: string) : Result<bool, string> =
        let mutable pres: Presentation = null

        let result =
            try
                pres <- openPpt app source true
                Ok(pres.HasVBProject)
            with ex ->
                Error(OsHost.Errors.explain ex)

        if not (isNull pres) then closePpt pres
        result
