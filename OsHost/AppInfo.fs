/// 黑板模式：插件把「当前实例」的现场抄到黑板上，消费方按这段信息精确挂回那一只实例。
///
/// 这一层只认文本、进程、窗口、时间和 ROT 名字，**不认任何 Office 类型**，
/// 也绝不创建 / 销毁实例：认不出、对不上、已过时，一律交回 Error 让调用方停手。
/// 四家 host 的 GetApplication(info) 共用这里的「认人 + 校验 + 按名取对象」。
namespace OsHost

open System
open System.Globalization
open System.Text.RegularExpressions

/// 四家 COM host 的种类（与 XUnitHost.Kind 互不相干，这儿是 OsHost 的通用种类）
[<RequireQualifiedAccess>]
type AppKind =
    | Word
    | Excel
    | PowerPoint
    | AutoCAD

/// 黑板上的一条实例信息。字段都来自插件复制的那段文本，也可以结构化构造。
type AppInfo =
    { Kind: AppKind
      Pid: int
      /// ROT 注册名，例：!WordVSTO.Instance.27232；空串表示只能按 ProgID 找
      Rot: string
      Hwnd: nativeint
      ExePath: string
      /// 采集时刻（Unix 毫秒）：所有「过没过时」的判断都拿它当基准
      CapturedMs: int64
      DocName: string option
      DocPath: string option
      Title: string }

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module AppInfo =

    let image (kind: AppKind) : string =
        match kind with
        | AppKind.Word -> "WINWORD"
        | AppKind.Excel -> "EXCEL"
        | AppKind.PowerPoint -> "POWERPNT"
        | AppKind.AutoCAD -> "acad"

    let kindName (kind: AppKind) : string =
        match kind with
        | AppKind.Word -> "Word"
        | AppKind.Excel -> "Excel"
        | AppKind.PowerPoint -> "PowerPoint"
        | AppKind.AutoCAD -> "AutoCAD"

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

    /// 解析插件复制的黑板文本（剪贴板原文、会话里贴的一段、存下来的 txt 都行）。
    /// 缺 PID / 采集时间 / 认不出种类，一律 Error —— 信息不完整就不该往下走。
    let ofText (text: string) : Result<AppInfo, string> =
        if String.IsNullOrWhiteSpace text then
            Error "黑板信息是空的"
        else
            let lines =
                text.Replace("\r\n", "\n").Split '\n'
                |> Array.map(fun line -> line.Trim())
                |> Array.filter(fun line -> not (String.IsNullOrWhiteSpace line))

            let title =
                lines
                |> Array.tryFind(fun line -> line.StartsWith "[")
                |> Option.defaultValue ""

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
                      DocPath = docPath lines
                      Title = title }

    /// 结构化构造（别的程序直接给字段时用），同样要求采集时间。
    let ofParts
        (kind: AppKind)
        (pid: int)
        (rot: string)
        (hwnd: nativeint)
        (exePath: string)
        (capturedMs: int64)
        : AppInfo =
        { Kind = kind
          Pid = pid
          Rot = rot
          Hwnd = hwnd
          ExePath = exePath
          CapturedMs = capturedMs
          DocName = None
          DocPath = None
          Title = "(结构化构造)" }

    let describe (info: AppInfo) : string =
        sprintf
            "%s pid=%d rot=%s hwnd=%A 采集=%s"
            (kindName info.Kind)
            info.Pid
            (if String.IsNullOrWhiteSpace info.Rot then "-" else info.Rot)
            info.Hwnd
            (Clock.show info.CapturedMs)

    let private firstError (checks: (bool * string) list) : string option =
        checks |> List.tryPick(fun (isBad, message) -> if isBad then Some message else None)

    /// 附着之前的「不过时」核对：进程还在、没被复用、窗口还属于它、EXE 路径对得上、
    /// ROT 名里的进程号与 PID 行一致。任何一条不通过就拒接，绝不改用 GetActiveObject 猜。
    let verify (info: AppInfo) : Result<unit, string> =
        let alive = Proc.ids (image info.Kind) |> Set.contains info.Pid
        let started = Proc.startedMs info.Pid
        let exe = Proc.exePath info.Pid

        let hwndPid =
            if info.Hwnd = 0n then
                None
            else
                Win32.pidOfWindow(Win32.root info.Hwnd)

        let rotPid =
            let dot = info.Rot.LastIndexOf '.'

            if dot < 0 then
                None
            else
                match
                    Int32.TryParse(
                        info.Rot.Substring(dot + 1),
                        NumberStyles.Integer,
                        CultureInfo.InvariantCulture
                    )
                with
                | true, number -> Some number
                | _ -> None

        let checks =
            [ (info.Pid <= 0, "黑板信息缺少 PID：不猜、不新建")
              (info.CapturedMs <= 0L, "黑板信息缺少采集时间：无法判断信息是否过时")
              ( not alive,
                sprintf
                    "PID %d 的 %s 已经不在运行：黑板信息已过时，绝不另起一个"
                    info.Pid
                    (image info.Kind) )
              ( (match started with
                 | Some ms -> ms > info.CapturedMs + 2000L
                 | None -> false),
                sprintf "PID %d 在采集之后才启动（进程号被复用）：黑板信息已过时" info.Pid )
              ( (match hwndPid with
                 | Some number -> number <> info.Pid
                 | None -> info.Hwnd <> 0n),
                sprintf "HWND %A 已不属于 PID %d（窗口关过或换了）：黑板信息已过时" info.Hwnd info.Pid )
              ( not (String.IsNullOrWhiteSpace info.ExePath)
                && (match exe with
                    | Some path -> not (String.Equals(path, info.ExePath, StringComparison.OrdinalIgnoreCase))
                    | None -> false),
                sprintf "EXE 路径对不上：黑板上是 %s" info.ExePath )
              ( (match rotPid with
                 | Some number -> number <> info.Pid
                 | None -> false),
                sprintf "ROT 名字里的进程号与 PID 行不一致：%s" info.Rot ) ]

        match firstError checks with
        | Some message -> Error message
        | None -> Ok()

    /// 各家的 GetApplication 第一道门：别把别家的实例认成自己。
    let ensureKind (kind: AppKind) (info: AppInfo) : Result<unit, string> =
        if info.Kind <> kind then
            Error(
                sprintf
                    "黑板信息是 %s，%s 的 GetApplication 不接：别把别家的实例认成自己"
                    (kindName info.Kind)
                    (kindName kind)
            )
        else
            Ok()

    /// 挂上之后再核一次「黑板上那份文档还是不是现在这份」：
    /// 名字 / 路径不符，或文件在采集之后被保存过，都说明中途有别的东西动过 —— 拒接。
    /// 黑板信息里没抄文档行就跳过这一层（老版插件只抄了五行）。
    let checkDoc (info: AppInfo) (currentName: string option) (currentPath: string option) : Result<unit, string> =
        let same (left: string) (right: string) =
            String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

        let checks =
            [ ( (match info.DocName, currentName with
                 | Some expected, Some actual -> not (same expected actual)
                 | Some _, None -> true
                 | _ -> false),
                sprintf "当前文档与黑板信息不一致：黑板上是 %A，现在是 %A" info.DocName currentName )
              ( (match info.DocPath, currentPath with
                 | Some expected, Some actual -> not (same expected actual)
                 | Some _, None -> true
                 | _ -> false),
                sprintf "当前文档路径与黑板信息不一致：黑板上是 %A，现在是 %A" info.DocPath currentPath )
              ( (match info.DocPath with
                 | Some path -> Paths.savedMs(Some path) > info.CapturedMs
                 | None -> false),
                sprintf "文档在采集时间 %s 之后被保存过：黑板信息已过时" (Clock.show info.CapturedMs) ) ]

        match firstError checks with
        | Some message -> Error message
        | None -> Ok()

    /// 四家 GetApplication 共用的一段：种类 → 现场校验 → 按 ROT 名取现成对象 → 交给各家强类型检查。
    /// 全程在**固定那条长活 STA 线程**上（Sta.run）：ROT 拿到的 RCW 和后面的强类型调用必须同公寓，
    /// 临时 STA 线程一退出，交出去的 Application 就没人给它泵消息了（见 gotchas）。
    /// 调用方后续对 Application 的读写也请包在 `Sta.run` 里；只想一步做完就直接调 attach 传 work。
    /// 名字不在 ROT 里就 Error，**不退回** GetActiveObject —— 多实例机器上那等于认错人。
    let attach
        (info: AppInfo)
        (kind: AppKind)
        (progIds: string list)
        (work: obj -> Result<'T, string>)
        : Result<'T, string> =
        ensureKind kind info
        |> Result.bind(fun () -> verify info)
        |> Result.bind(fun () ->
            Sta.run(fun () ->
                let found =
                    if String.IsNullOrWhiteSpace info.Rot then
                        Com.running progIds
                    else
                        Com.runningAt info.Rot

                match found with
                | Some raw -> work raw
                | None ->
                    let target =
                        if String.IsNullOrWhiteSpace info.Rot then
                            String.Join(" / ", progIds)
                        else
                            info.Rot

                    Error(sprintf "ROT 里没有 %s：插件没登记或实例已退出，绝不新建" target)))
