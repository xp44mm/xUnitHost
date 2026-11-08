/// 黑板模式：一张与种类完全解耦的通用参数表 GeneralAppInfo（现在只此一条路）。
///
/// 旧路 `AppInfo` 把「种类」钉在封闭枚举 `AppKind` 上，由枚举派生进程名与显示名；
/// 2026-11-07 14:11 按用户决定把 `OsHost/AppInfo.fs` 整份删除，四家各自的进程名 / 显示名 / 标记
/// 搬进专用 Host 的 `Host.SelfIdentity`，四家 `GetApplication` 统一改收 `GeneralAppInfo`。
/// 这一份把身份交给调用方用**参数**填：
///   · GeneralIdentity：Image（进程名）/ Name（显示名）/ Tag（调用方自己的种类标记）；
///   · 「怎么拿到现成对象」「拿到 obj 之后怎么转强类型、核身份」用**获取函数**占位：
///     attachWith 的 acquire / work 两个参数，库一行 Office 类型都不认，也绝不新建实例。
/// 于是第五家、第十家不用改库。
/// 现场校验口径（进程 / 启动时间 / HWND / EXE / ROT 后缀）不变；文档级交接检查按用户要求推迟。
namespace OsHost

open System
open System.Globalization

/// 身份参数：三个值都由调用方填（四家专用 Host 各有一份 SelfIdentity）。库只比字符串，不解释 Tag 的含义。
type GeneralIdentity =
    { /// 进程名（Proc.ids 用的 image，不带 .exe）
      Image: string
      /// 报错 / 报告里给人看的名字
      Name: string
      /// 调用方自己的种类标记，例如 "word" / "acad" / "et"
      Tag: string }

    /// 参数占位：三个值都由调用方填，不是从某个枚举推出来的。
    static member create (image: string, name: string, tag: string) : GeneralIdentity =
        { Image = image
          Name = name
          Tag = tag }

    member this.describe () : string =
        sprintf "%s(tag=%s, image=%s)" this.Name this.Tag this.Image

/// 通用参数表：**只装认进程用的东西**。没有枚举、没有展示文本、没有 COM 类型、没有文档。
/// 文档名 / 路径不参与认人（文档级交接检查已取消），另立 GeneralDocInfo，只当报告数据。
type GeneralAppInfo =
    { Identity: GeneralIdentity
      Pid: int
      /// ROT 注册名；空串时按调用方给的 ProgID 列表找
      Rot: string
      /// 窗口句柄；0n = 借用进程，不校验窗口
      Hwnd: nativeint
      ExePath: string
      /// 采集时刻（Unix 毫秒）：所有「过没过时」的判断都拿它当基准
      CapturedMs: int64 }

    /// 静态构造成员：识别进程的最小必填集。
    static member create
        (
            identity: GeneralIdentity,
            pid: int,
            rot: string,
            hwnd: nativeint,
            exePath: string,
            capturedMs: int64
        )
        : GeneralAppInfo =
        { Identity = identity
          Pid = pid
          Rot = rot
          Hwnd = hwnd
          ExePath = exePath
          CapturedMs = capturedMs }

    /// 一行摘要，贴报告 / 报错用
    member this.describe () : string =
        sprintf
            "%s pid=%d rot=%s hwnd=%A 采集=%s"
            this.Identity.Name
            this.Pid
            (if String.IsNullOrWhiteSpace this.Rot then "-" else this.Rot)
            this.Hwnd
            (Clock.show this.CapturedMs)

/// 文档信息：插件抄下来的「当前文档名 / 落盘路径」。
/// 与 GeneralAppInfo 分开的道理：那张表只用来认进程，文档换了、关了都不该影响认人；
/// 这两个值只进报告 / 日志，不参与任何判断。
type GeneralDocInfo =
    { /// 文档名（不含路径），例如「文档1」「Book1」「Drawing1.dwg」
      Name: string option
      /// 落盘路径；新文档还没存过盘就是 None
      Path: string option }

    static member create (name: string option, path: string option) : GeneralDocInfo =
        { Name = name; Path = path }

    static member empty : GeneralDocInfo = { Name = None; Path = None }

    /// 一行摘要，贴报告用
    member this.describe () : string =
        let show (value: string option) =
            match value with
            | Some text when not (String.IsNullOrWhiteSpace text) -> text
            | _ -> "(无)"

        sprintf "名字=%s 路径=%s" (show this.Name) (show this.Path)

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module GeneralAppInfo =

    let private firstError (checks: (bool * string) list) : string option =
        checks |> List.tryPick(fun (isBad, message) -> if isBad then Some message else None)

    /// 借用进程：只认进程身份，不认别人文档窗口的句柄。
    /// HWND 是插件抄下来的「调用方当前打开的文档窗口」，借用方打开 / 关闭自己的文件后
    /// 它必然被换掉（甚至那个窗口整个关掉），拿它当交接凭据会让下一轮凭空断链。
    /// 这里把 Hwnd 置 0n —— verify 本来就把它定义为「没填这一项，就不校验窗口」；
    /// PID 存活 / 采集时间 / 启动时刻 / EXE / ROT 名后缀一条不少。
    /// 调用方仍可保留原表里的 HWND 数值作数据（报告、日志用），只是不交给入口判断。
    let borrowProcess (info: GeneralAppInfo) : GeneralAppInfo =
        { info with Hwnd = 0n }

    /// 附着之前的「不过时」核对：就是原先旧路那一套口径，只是进程名 / 显示名从参数来。
    /// HWND 是**可选**凭据：Hwnd = 0n 表示借用进程、不校验窗口（见 borrowProcess）。
    /// 任何一条不通过就拒接，绝不改用 GetActiveObject 猜。
    let verify (info: GeneralAppInfo) : Result<unit, string> =
        let image = info.Identity.Image
        let alive = Proc.ids image |> Set.contains info.Pid
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
            [ (info.Pid <= 0, "通用参数表缺少 PID：不猜、不新建")
              (info.CapturedMs <= 0L, "通用参数表缺少采集时间：无法判断信息是否过时")
              (String.IsNullOrWhiteSpace image, "通用参数表缺少 Image（进程名）：不猜、不新建")
              ( not alive,
                sprintf
                    "PID %d 的 %s 已经不在运行：黑板信息已过时，绝不另起一个"
                    info.Pid
                    image )
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

    /// 第一道门：别把别家的实例认成自己。这里比的是调用方自己填的 Tag，库不认任何封闭种类。
    let ensureTag (tag: string) (info: GeneralAppInfo) : Result<unit, string> =
        if not (String.Equals(tag, info.Identity.Tag, StringComparison.OrdinalIgnoreCase)) then
            Error(
                sprintf
                    "通用参数表是 %s(tag=%s)，tag=%s 的入口不接：别把别家的实例认成自己"
                    info.Identity.Name
                    info.Identity.Tag
                    tag
            )
        else
            Ok()

    /// 文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求**推迟**（2026-11-07 13:51）：
    /// 新路同样不做，`checkDoc` 已撤掉；要找回旧实现看 git 历史（`6be1725`）。
    /// 2026-11-08 起文档信息整块搬出本表（GeneralDocInfo）：认进程的表里不再有文档字段。

    /// 默认获取函数：按 ROT 名精确取；没填 ROT 名就按调用方给的 ProgID 列表找。
    /// 必须在 Sta.run 那条固定 STA 线程上被调用（attachWith 已保证）。
    let defaultAcquire (progIds: string list) (info: GeneralAppInfo) : obj option =
        if String.IsNullOrWhiteSpace info.Rot then
            Com.running progIds
        else
            Com.runningAt info.Rot

    /// 获取函数占位：acquire 决定「怎么取到现成对象」，work 决定「拿到 obj 之后怎么转强类型 / 核身份」。
    /// 两者都由调用方注入，因此库不需要知道 Word / Excel / PowerPoint / AutoCAD 里任何一家。
    /// 全程在**固定那条长活 STA 线程**上（Sta.run）：ROT 拿到的 RCW 与后续强类型调用必须同公寓。
    /// 取不到就 Error，绝不退回 GetActiveObject 猜、绝不新建。
    let attachWith
        (info: GeneralAppInfo)
        (tag: string)
        (acquire: GeneralAppInfo -> obj option)
        (work: obj -> Result<'T, string>)
        : Result<'T, string> =
        ensureTag tag info
        |> Result.bind(fun () -> verify info)
        |> Result.bind(fun () ->
            Sta.run(fun () ->
                match acquire info with
                | Some raw -> work raw
                | None ->
                    Error(
                        sprintf
                            "按注入的获取函数没拿到 %s（rot=%s）的现成实例：绝不新建"
                            info.Identity.Name
                            (if String.IsNullOrWhiteSpace info.Rot then "-" else info.Rot)
                    )))

    /// 常规用法：取对象走默认的 ROT / ProgID，强类型转换与身份核对由 work 注入。
    let attach
        (info: GeneralAppInfo)
        (tag: string)
        (progIds: string list)
        (work: obj -> Result<'T, string>)
        : Result<'T, string> =
        attachWith info tag (defaultAcquire progIds) work
