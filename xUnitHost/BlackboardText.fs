/// 【参考实现，不是库的 API】把插件复制到剪贴板的「黑板展示文本」读成结构化参数表 AppInfo。
///
/// 生产库 OsHost 只认 AppInfo 那张参数表：调用方（插件 / 宿主程序）按字段填报，
/// 库不认文本、不解析排版。插件改了标题、键名、措辞，只改这一份参考代码，不动库。
/// 用不上文本源的程序请直接填参数表：
///   { (AppInfo.ofParts AppKind.Word 27232 "!WordVSTO.Instance.27232" 0n "" captured)
///       with DocName = Some "文档1" }
/// 解析出来的字段与手工填报的字段完全相同，四家 GetApplication(info) 分不出来源。
namespace XUnitHost

open System
open System.Globalization
open System.Text.RegularExpressions
open OsHost

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

    let private parseHwnd (raw: string) : nativeint =
        let text = raw.Trim()

        if String.IsNullOrWhiteSpace text then
            0n
        elif text.StartsWith("0x", StringComparison.OrdinalIgnoreCase) then
            match UInt32.TryParse(text.Substring 2, NumberStyles.HexNumber, CultureInfo.InvariantCulture) with
            | true, value -> nativeint (int64 value)
            | _ -> 0n
        else
            match Int64.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture) with
            | true, value -> nativeint value
            | _ -> 0n

    let private parseClock (text: string) : int64 option =
        let hit =
            Regex.Match(text, @"采集时间\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})")

        if not hit.Success then
            None
        else
            match
                DateTime.TryParseExact(
                    hit.Groups.[1].Value,
                    "yyyy-MM-dd HH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None
                )
            with
            | true, moment -> Some(Clock.unixMs moment)
            | _ -> None

    let private kindOfLine (lines: string array) : AppKind option =
        match valueOf "KIND" lines with
        | Some value when contains "word" value -> Some AppKind.Word
        | Some value when contains "excel" value -> Some AppKind.Excel
        | Some value when contains "ppt" value || contains "powerpoint" value -> Some AppKind.PowerPoint
        | Some value when contains "acad" value -> Some AppKind.AutoCAD
        | _ -> None

    /// 从插件的标题 / ROT 名 / EXE 路径认是哪一家
    let private kindOfText (text: string) : AppKind option =
        if contains "WordVSTO" text || contains "WINWORD.EXE" text then
            Some AppKind.Word
        elif contains "ExcelVSTO" text || contains "EXCEL.EXE" text then
            Some AppKind.Excel
        elif contains "PowerPointVSTO" text || contains "PPTVSTO" text || contains "POWERPNT" text then
            Some AppKind.PowerPoint
        elif contains "AcadVSTO" text || contains "AutoCAD" text || contains "ACAD.EXE" text then
            Some AppKind.AutoCAD
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

    /// 标题行只给人看，不进参数表；单独留一个函数是为了写报告时能贴原文。
    let caption (text: string) : string =
        text.Replace("\r\n", "\n").Split '\n'
        |> Array.map(fun line -> line.Trim())
        |> Array.tryFind(fun line -> line.StartsWith "[")
        |> Option.defaultValue "(没有标题行)"

    /// 参考解析：插件剪贴板原文 / 会话里贴的一段 / 存下来的 txt → AppInfo 参数表。
    /// 缺 PID / 采集时间 / 认不出种类，一律 Error —— 信息不完整就不该往下走。
    /// 注意这里只是「把展示文本填成参数表」，真正的校验在 OsHost.AppInfo.verify / checkDoc。
    let ofText (text: string) : Result<AppInfo, string> =
        if String.IsNullOrWhiteSpace text then
            Error "黑板信息是空的"
        else
            let lines =
                text.Replace("\r\n", "\n").Split '\n'
                |> Array.map(fun line -> line.Trim())
                |> Array.filter(fun line -> not (String.IsNullOrWhiteSpace line))

            let kind =
                match kindOfLine lines with
                | Some value -> Some value
                | None -> kindOfText text

            let pid =
                match valueOf "PID" lines with
                | Some value ->
                    match Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture) with
                    | true, number -> number
                    | _ -> 0
                | None -> 0

            let capturedMs = parseClock text |> Option.defaultValue 0L

            if kind.IsNone then
                Error "认不出是哪一家 COM：黑板信息里既没有插件名，也没有 KIND 行"
            elif pid <= 0 then
                Error "黑板信息里没有有效 PID：不猜、不新建"
            elif capturedMs <= 0L then
                Error "黑板信息里没有有效采集时间：无法判断信息是否过时"
            else
                Ok
                    { Kind = kind.Value
                      Pid = pid
                      Rot = valueOf "ROT" lines |> Option.defaultValue ""
                      Hwnd = valueOf "HWND" lines |> Option.map parseHwnd |> Option.defaultValue 0n
                      ExePath = exeLine lines
                      CapturedMs = capturedMs
                      DocName = docName lines
                      DocPath = docPath lines }
