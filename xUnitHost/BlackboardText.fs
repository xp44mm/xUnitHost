/// 【参考实现，不是库的 API】把插件复制到剪贴板的「黑板展示文本」读成「哪一家 + 哪只进程」。
///
/// 2026-11-13 起没有「通用参数表」这层抽象了：挂回只认 **pid**，进程名、ROT 名只作报告里的佐证。
/// 所以这个类型留在测试项目里（`BoardApp` / `BoardDoc`），OsHost 不再认识任何「参数表」概念：
///   插件改了标题、键名、措辞，只改这一份参考代码，不动库。
/// 挂回入口一律按进程号：
///   Word / Excel / AutoCAD → `*Task.GetApplicationAsync pid`，PowerPoint → `PowerPointHost.Host.GetApplication pid`。
/// 文档名 / 路径不在 BoardApp 里：要就用 docOfText 另取一条 BoardDoc（只作报告数据）。
/// 文本里的 HWND 行、采集时间行都不解析：窗口与抄录时刻不属于「这是哪只进程」。
namespace XUnitHost

open System
open System.Globalization
open OsHost

/// 黑板抄下来的「哪一家 + 哪只进程」。挂回只用 Pid；Image / Rot 只贴报告。
type BoardApp =
    { /// 进程名（Proc.ids 的口径，不带 .exe）
      Image: string
      Pid: int
      /// ROT 注册名；只作佐证，不参与认人
      Rot: string }

    /// 一行摘要，贴报告 / 报错用
    member this.describe () : string =
        sprintf
            "%s pid=%d rot=%s"
            (if String.IsNullOrWhiteSpace this.Image then "(没有进程名)" else this.Image)
            this.Pid
            (if String.IsNullOrWhiteSpace this.Rot then "-" else this.Rot)

/// 文档信息：插件抄下来的「当前文档名 / 落盘路径」，**只进报告、不参与任何判断**。
type BoardDoc =
    { /// 文档名（不含路径），例如「文档1」「Book1」「Drawing1.dwg」
      Name: string option
      /// 落盘路径；新文档还没存过盘就是 None
      Path: string option }

    static member empty : BoardDoc = { Name = None; Path = None }

    /// 一行摘要，贴报告用
    member this.describe () : string =
        let show (value: string option) =
            match value with
            | Some text when not (String.IsNullOrWhiteSpace text) -> text
            | _ -> "(无)"

        sprintf "名字=%s 路径=%s" (show this.Name) (show this.Path)

module internal BlackboardText =

    let private contains (needle: string) (hay: string) : bool =
        hay.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0

    /// 取「键 : 值」里的值。行里再有冒号（例如盘符 C:\）不影响 —— 只认第一个冒号。
    let private valueOf (key: string) (lines: string array) : string option =
        lines
        |> Array.tryPick(fun line ->
            let i = line.IndexOf ':'

            if
                i > 0
                && String.Equals(line.Substring(0, i).Trim(), key, StringComparison.OrdinalIgnoreCase)
            then
                Some(line.Substring(i + 1).Trim())
            else
                None)

    /// EXE 路径那一行的键各家不同（WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE / ACAD.EXE）
    let private exeLine (lines: string array) : string =
        lines
        |> Array.tryPick(fun line ->
            let i = line.IndexOf ':'

            if i > 0 && line.Substring(0, i).Trim().EndsWith(".EXE", StringComparison.OrdinalIgnoreCase) then
                Some(line.Substring(i + 1).Trim())
            else
                None)
        |> Option.defaultValue ""

    /// 认出是哪一家 → 取那家专用 Host 上的进程名
    let private imageOfLine (lines: string array) : string option =
        match valueOf "KIND" lines with
        | Some value when contains "word" value -> Some WordHost.Host.SelfImage
        | Some value when contains "excel" value -> Some ExcelHost.Host.SelfImage
        | Some value when contains "ppt" value || contains "powerpoint" value ->
            Some PowerPointHost.Host.SelfImage
        | _ -> None

    /// 从插件的标题 / ROT 名 / EXE 路径认是哪一家
    let private imageOfText (text: string) : string option =
        if contains "WordVSTO" text || contains "WINWORD.EXE" text then
            Some WordHost.Host.SelfImage
        elif contains "ExcelVSTO" text || contains "EXCEL.EXE" text then
            Some ExcelHost.Host.SelfImage
        elif contains "PowerPointVSTO" text || contains "PPTVSTO" text || contains "POWERPNT" text then
            Some PowerPointHost.Host.SelfImage
        else
            None

    /// 「(无)」「(未保存)」「(无打开的文档)」都当没有这一项
    let private optional (raw: string option) : string option =
        match raw with
        | None -> None
        | Some value when String.IsNullOrWhiteSpace value -> None
        | Some value when value.StartsWith "(" -> None
        | Some value -> Some value

    let private docName (lines: string array) : string option =
        [ "文档"; "工作簿"; "演示文稿"; "图纸" ]
        |> List.tryPick(fun key -> valueOf key lines)
        |> optional

    let private docPath (lines: string array) : string option =
        valueOf "路径" lines |> optional |> Option.bind Paths.canonical

    /// 标题行只给人看，不进任何结构；单独留一个函数是为了写报告时能贴原文。
    let caption (text: string) : string =
        text.Replace("\r\n", "\n").Split '\n'
        |> Array.map(fun line -> line.Trim())
        |> Array.tryFind(fun line -> line.StartsWith "[")
        |> Option.defaultValue "(没有标题行)"

    /// 展示文本 → 非空行数组（每行去首尾空白）。进程信息与文档信息各取一次，都是从这儿切。
    let private linesOf (text: string) : string array =
        text.Replace("\r\n", "\n").Split '\n'
        |> Array.map(fun line -> line.Trim())
        |> Array.filter(fun line -> not (String.IsNullOrWhiteSpace line))

    /// 文档信息：与识别进程分开拿。缺键 /「(未保存)」都当没有，
    /// 认不出也不报错 —— 它只用来贴报告，不参与认人。
    let docOfText (text: string) : BoardDoc =
        if String.IsNullOrWhiteSpace text then
            BoardDoc.empty
        else
            let lines = linesOf text
            { Name = docName lines; Path = docPath lines }

    /// 参考解析：插件剪贴板原文 / 会话里贴的一段 / 存下来的 txt → BoardApp。
    /// 缺 PID / 认不出哪一家，一律 Error —— 信息不完整就不该往下走。
    /// 注意这里只是「把展示文本填成字段」，真正的现场核对在各家 Host 的 checkPid / findOnSta；
    /// 文档级交接检查按用户要求推迟（2026-11-07 13:51），暂不做。
    let ofText (text: string) : Result<BoardApp, string> =
        if String.IsNullOrWhiteSpace text then
            Error "黑板信息是空的"
        else
            let lines = linesOf text

            let image =
                match imageOfLine lines with
                | Some value -> Some value
                | None -> imageOfText text

            let pid =
                match valueOf "PID" lines with
                | Some value ->
                    match Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture) with
                    | true, number -> number
                    | _ -> 0
                | None -> 0

            if image.IsNone then
                Error "认不出是哪一家 COM：黑板信息里既没有插件名，也没有 KIND 行"
            elif pid <= 0 then
                Error "黑板信息里没有有效 PID：不猜、不新建"
            else
                Ok(
                    { Image = image.Value
                      Pid = pid
                      Rot = valueOf "ROT" lines |> Option.defaultValue "" }
                )
