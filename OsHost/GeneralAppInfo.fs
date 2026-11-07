/// 黑板模式的另一条路：一张与 AppKind 完全解耦的通用参数表 GeneralAppInfo。
///
/// 原 AppInfo 把「种类」钉在封闭的 AppKind 上（Word / Excel / PowerPoint / AutoCAD 四选一），
/// 由枚举派生进程名与显示名。这一份把那两块拿出来交给调用方：
///   · GeneralIdentity 用**参数**占位：Image（进程名）/ Name（显示名）/ Tag（调用方自己的种类标记）；
///   · 「怎么拿到现成对象」「拿到 obj 之后怎么转强类型、核身份」用**获取函数**占位：
///     attachWith 的 acquire / work 两个参数，库一行 Office 类型都不认，也绝不新建实例。
/// 于是第五家、第十家不用改库；四家 host 要不要搬过来，由调用方自己决定。
/// 校验口径（进程 / 启动时间 / HWND / EXE / ROT 后缀 / 文档没被换或存过）与 AppInfo 一致。
///
/// 原文件 OsHost/AppInfo.fs 与四家 host 的 GetApplication(info) 一行未动，两条路并存。
namespace OsHost

open System
open System.Globalization

/// 身份参数：取代 AppKind 上那两个 match。库只做字符串比较，不解释 Tag 的含义。
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

/// 通用参数表：一条实例信息的全部结构化值。没有枚举、没有展示文本、没有 COM 类型。
type GeneralAppInfo =
    { Identity: GeneralIdentity
      Pid: int
      /// ROT 注册名；空串时按调用方给的 ProgID 列表找
      Rot: string
      Hwnd: nativeint
      ExePath: string
      /// 采集时刻（Unix 毫秒）：所有「过没过时」的判断都拿它当基准
      CapturedMs: int64
      DocName: string option
      DocPath: string option }

    /// 静态构造成员：最小必填集；文档名 / 路径按记录更新补上。
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
          CapturedMs = capturedMs
          DocName = None
          DocPath = None }

    /// 一行摘要，贴报告 / 报错用
    member this.describe () : string =
        sprintf
            "%s pid=%d rot=%s hwnd=%A 采集=%s"
            this.Identity.Name
            this.Pid
            (if String.IsNullOrWhiteSpace this.Rot then "-" else this.Rot)
            this.Hwnd
            (Clock.show this.CapturedMs)

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module GeneralAppInfo =

    let private firstError (checks: (bool * string) list) : string option =
        checks |> List.tryPick(fun (isBad, message) -> if isBad then Some message else None)

    /// 附着之前的「不过时」核对：与 AppInfo.verify 同一套口径，只是进程名 / 显示名从参数来。
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

    /// 第一道门：别把别家的实例认成自己。这里比的是调用方自己填的 Tag，不是枚举。
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

    /// 挂上之后再核一次「黑板上那份文档还是不是现在这份」：
    /// 名字 / 路径不符，或文件在采集之后被保存过，都说明中途有别的东西动过 —— 拒接。
    /// 参数表里没填文档行就跳过这一层。
    let checkDoc
        (info: GeneralAppInfo)
        (currentName: string option)
        (currentPath: string option)
        : Result<unit, string> =
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
