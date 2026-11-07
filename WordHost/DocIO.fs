/// Word 转换共用的 COM 细节：开/关文档、另存、正文指纹、判宏、临时文件、只读读回。
/// 放在 WordHost 里，因为这些全是 Word 自己的活：DocToDocx / DocToDocm 两条转换，
/// 以及测试项目的事后核对，都走这一份，不重复实现。
/// 对应 ExcelHost/BookIO.fs。
namespace WordHost

open System
open System.IO
open Microsoft.Office.Interop.Word

module DocIO =

    let private noSave: obj = box 0 // wdDoNotSaveChanges

    /// Word 的互操作参数是 **obj&**（按引用传），Excel 才是按值传 object。
    /// ConfirmConversions=false 别弹「要不要转换格式」，AddToRecent=false 不留最近文档。
    /// **OpenNoRepairDialog**：老 .doc 一旦被判为损坏，Documents.Open 会弹「是否修复」模态框，
    /// COM 调用就一直挂到有人去点掉；这个方法是 Office 专门给自动化用的，不弹那个框。
    /// 再配合 Host.quieten 的 DisplayAlerts=wdAlertsNone / Options.ConfirmConversions=false。
    let openDoc (app: Application) (path: string) (readOnly: bool) : Document =
        let mutable fileName: obj = box path
        let mutable confirmConversions: obj = box false
        let mutable ro: obj = box readOnly
        let mutable addToRecent: obj = box false
        app.Documents.OpenNoRepairDialog(&fileName, &confirmConversions, &ro, &addToRecent)

    /// 另存成新格式；格式码由各条转换自己带，这里不认得 .docx / .docm 谁是谁。
    let saveDocAs (fmt: obj) (doc: Document) (target: string) =
        let mutable fileName: obj = box target
        let mutable format: obj = fmt
        doc.SaveAs(&fileName, &format) |> ignore

    let closeDoc (doc: Document) =
        try
            let mutable save: obj = noSave
            doc.Close(&save)
        with _ ->
            ()

    /// 正文比较口径：去掉段落标记、单元格标记和两端空白。
    /// .doc 与 .docx / .docm 的正文该完全一样。
    let normalize (text: string) : string =
        if isNull text then "" else text.Replace("\r", "").Replace("\a", "").Trim()

    /// 一份文档的正文指纹
    let bodyText (doc: Document) : string = normalize doc.Content.Text

    /// 读 VBA 工程在不在。读不出来（信任设置没开）就当「没有」，调用方自己决定怎么办。
    let hasVba (doc: Document) : bool =
        try
            doc.HasVBProject
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

    /// 只读读回正文：给事后核对用
    let readText (app: Application) (path: string) : string =
        let doc = openDoc app path true
        try bodyText doc finally closeDoc doc

    /// 分派用：只读开一下 .doc，判有没有 VBA 宏。
    /// **读不出 VBA 工程信息就报错**（原件保留），绝不猜「没宏」——猜错会把宏悄悄丢掉。
    let detectMacro (app: Application) (source: string) : Result<bool, string> =
        let mutable doc: Document = null

        let result =
            try
                doc <- openDoc app source true
                Ok(doc.HasVBProject)
            with ex ->
                Error(OsHost.Errors.explain ex)

        if not (isNull doc) then closeDoc doc
        result
