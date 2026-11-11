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

    /// 正文比较口径：去掉段落标记、单元格标记、**域控制符**和两端空白。
    /// .doc 与 .docx / .docm 的正文该完全一样。
    ///
    /// 域控制符 `\u0013` / `\u0014` / `\u0015`（field begin / separator / end）是 Word 内部的
    /// 不可见标记：老 .doc 另存成 .docx 时 Word 会把这类残留标记规范化掉，可见文字一个不少。
    /// 2026-11-11 实测：连续 5 份 .doc 都只差这一个 `\x15`（`\x01` 图片占位前后），
    /// 按原始串逐字比就一律判「正文对不上」。
    /// **域里的可见内容（含域代码文字）照常参与比较**，真丢内容仍然会被抓出来；
    /// `\u0001`（内嵌图片/对象占位）也保留 —— 图片数量少一张必须能发现。
    let normalize (text: string) : string =
        if isNull text then
            ""
        else
            text
                .Replace("\r", "")
                .Replace("\a", "")
                .Replace("\u0013", "")
                .Replace("\u0014", "")
                .Replace("\u0015", "")
                .Trim()

    /// 一份文档的正文指纹
    let bodyText (doc: Document) : string = normalize doc.Content.Text

    /// 读 VBA 工程在不在。读不出来（信任设置没开）就当「没有」，调用方自己决定怎么办。
    let hasVba (doc: Document) : bool =
        try
            doc.HasVBProject
        with _ ->
            false

    /// 迟绑定取 COM 属性：不引 Microsoft.Vbe.Interop（那是 VBE 类型库，装没装都不一定），
    /// 取不到就 None —— 调用方按「读不出来」处理，不猜。
    let private comGet (target: obj) (name: string) (args: obj[]) : obj option =
        if isNull target then
            None
        else
            try
                Some(
                    target
                        .GetType()
                        .InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, target, args)
                )
            with _ ->
                None

    /// VBA 工程里的代码总行数：
    ///   `Some 0` = 有工程但一行代码都没有（老模板留下的空壳）
    ///   `Some n` = 真有 n 行代码
    ///   `None`   = 读不到（没开「信任对 VBA 工程对象模型的访问」AccessVBOM）
    let vbaLines (doc: Document) : int option =
        match comGet (box doc) "VBProject" [||] with
        | None -> None
        | Some project ->
            match comGet project "VBComponents" [||] with
            | None -> None
            | Some comps ->
                match comGet comps "Count" [||] with
                | None -> None
                | Some count ->
                    let mutable total = 0
                    let mutable ok = true

                    for index in 1 .. Convert.ToInt32 count do
                        let codeModule =
                            comGet comps "Item" [| box index |]
                            |> Option.bind(fun item -> comGet item "CodeModule" [||])

                        match codeModule |> Option.bind(fun module' -> comGet module' "CountOfLines" [||]) with
                        | Some n -> total <- total + Convert.ToInt32 n
                        | None -> ok <- false

                    if ok then Some total else None

    /// 真的要按「带宏」处理吗？
    /// **有工程且有代码**才算带宏；空壳工程（0 行代码）一律当无宏：
    /// 2026-11-11 实证——7 份 .doc 都只有 1 个组件、0 行代码，另存 .docm 时 Word 会把空壳丢掉，
    /// 于是 `HasVBProject` 判「有宏」→ 存 .docm → 复核没宏 → 判失败，白卡 7 份；
    /// 这类文件按 .docx 转不丢任何内容（本来就没代码可丢）。
    /// 读不到代码行数（没开 AccessVBOM）时保守当「有宏」，绝不猜「没宏」把真代码丢了。
    let hasRealVba (doc: Document) : bool =
        if not (hasVba doc) then
            false
        else
            match vbaLines doc with
            | Some 0 -> false
            | _ -> true

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

    /// 分派用：只读开一下 .doc，判有没有**真**宏（空壳工程不算）。
    /// **读不出 VBA 工程信息就报错**（原件保留），绝不猜「没宏」——猜错会把宏悄悄丢掉。
    let detectMacro (app: Application) (source: string) : Result<bool, string> =
        let mutable doc: Document = null

        let result =
            try
                doc <- openDoc app source true
                Ok(hasRealVba doc)
            with ex ->
                Error(OsHost.Errors.explain ex)

        if not (isNull doc) then closeDoc doc
        result
