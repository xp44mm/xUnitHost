/// 黑板模式：一张只装「怎么找到那只进程」的通用参数表 GeneralAppInfo（现在只此一条路）。
///
/// 旧路 `AppInfo` 把「种类」钉在封闭枚举 `AppKind` 上，由枚举派生进程名与显示名；
/// 2026-11-07 14:11 按用户决定把 `OsHost/AppInfo.fs` 整份删除。
/// 2026-11-08 又按用户决定删掉 `GeneralIdentity`：认种类、显示名、种类标记一律不要，
/// 表里只留现场值 —— Image（进程名）+ PID + ROT 名 + EXE + 采集时刻，
/// 够把那只进程认出来就行，别再往上叠概念。
/// 2026-11-09 删掉 HWND：插件抄的是「它当时打开的文档窗口」，借用进程后必然换，
/// 拿它当凭据只会断链；`borrowProcess` 于是成了空操作，一并撤掉。
/// 「怎么拿到现成对象」「拿到 obj 之后怎么转强类型、核身份」用**获取函数**占位：
/// attachWith 的 acquire / work 两个参数，库一行 Office 类型都不认，也绝不新建实例。
/// 于是第五家、第十家不用改库。
/// 现场校验口径（进程 / 启动时间 / EXE / ROT 后缀）；文档级交接检查按用户要求推迟。
namespace OsHost

open System
open System.Globalization

/// 通用参数表：**只装找到那只进程所需的值**。没有种类枚举、没有显示名、没有种类标记、
/// 没有窗口句柄、没有 COM 类型、没有文档 —— 文档信息另立 GeneralDocInfo，只当报告数据。
type GeneralAppInfo =
    { /// 进程名（Proc.ids 用的 image，不带 .exe），例如 "WINWORD" / "EXCEL" / "acad"
      Image: string
      Pid: int
      /// ROT 注册名；空串时按调用方给的 ProgID 列表找
      Rot: string
      ExePath: string
      /// 采集时刻（Unix 毫秒）：所有「过没过时」的判断都拿它当基准
      CapturedMs: int64 }

    /// 静态构造成员：找到一只进程的最小必填集。
    static member create
        (
            image: string,
            pid: int,
            rot: string,
            exePath: string,
            capturedMs: int64
        )
        : GeneralAppInfo =
        { Image = image
          Pid = pid
          Rot = rot
          ExePath = exePath
          CapturedMs = capturedMs }

    /// 一行摘要，贴报告 / 报错用
    member this.describe () : string =
        sprintf
            "%s pid=%d rot=%s exe=%s 采集=%s"
            (if String.IsNullOrWhiteSpace this.Image then "(没有进程名)" else this.Image)
            this.Pid
            (if String.IsNullOrWhiteSpace this.Rot then "-" else this.Rot)
            (if String.IsNullOrWhiteSpace this.ExePath then "-" else this.ExePath)
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

    /// 附着之前的「不过时」核对：进程名 / PID 存活 / 启动时刻 / EXE / ROT 后缀，一样不落地比。
    /// 不核窗口：窗口句柄属于「谁当时打开了什么文档」，不属于「这是哪只进程」。
    /// 任何一条不通过就拒接，绝不改用 GetActiveObject 猜。
    let verify (info: GeneralAppInfo) : Result<unit, string> =
        let image = info.Image
        let alive = Proc.ids image |> Set.contains info.Pid
        let started = Proc.startedMs info.Pid
        let exe = Proc.exePath info.Pid

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

    /// 文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求**推迟**（2026-11-07 13:51）：
    /// 新路同样不做，`checkDoc` 已撤掉；要找回旧实现看 git 历史（`6be1725`）。
    /// 2026-11-08 起文档信息整块搬出本表（GeneralDocInfo）：认进程的表里不再有文档字段。
    /// 同一天按用户决定撤掉 `ensureTag`：表里没有种类标记，进错入口的代价是 work 里强转失败
    /// 或 pid 核对不符 —— 一样是明确 Error，不必为此多留一个字段。
    /// 2026-11-09 撤掉 `borrowProcess` 与 HWND 校验：字段没了，清零也就没意义。

    /// 默认获取函数：按 ROT 名精确取；没填 ROT 名就按调用方给的 ProgID 列表找。
    /// 必须在 Sta.run 那条固定 STA 线程上被调用（attachWith 已保证）。
    let defaultAcquire (progIds: string list) (info: GeneralAppInfo) : obj option =
        if String.IsNullOrWhiteSpace info.Rot then
            Com.running progIds
        else
            Com.runningAt info.Rot

    /// 获取函数占位：acquire 决定「怎么取到现成对象」，work 决定「拿到 obj 之后怎么转强类型 / 核身份」。
    /// 两者都由调用方注入，因此库不需要知道 Word / Excel / PowerPoint / AutoCAD 里任何一家。
    /// 全部只认「找到那只进程」的现场值，没有种类标记可对：进错入口会在 work 里强转失败或 pid 不符。
    /// 全程在**固定那条长活 STA 线程**上（Sta.run）：ROT 拿到的 RCW 与后续强类型调用必须同公寓。
    /// 取不到就 Error，绝不退回 GetActiveObject 猜、绝不新建。
    let attachWith
        (info: GeneralAppInfo)
        (acquire: GeneralAppInfo -> obj option)
        (work: obj -> Result<'T, string>)
        : Result<'T, string> =
        verify info
        |> Result.bind(fun () ->
            Sta.run(fun () ->
                match acquire info with
                | Some raw -> work raw
                | None ->
                    Error(
                        sprintf
                            "按注入的获取函数没拿到 %s（rot=%s）的现成实例：绝不新建"
                            info.Image
                            (if String.IsNullOrWhiteSpace info.Rot then "-" else info.Rot)
                    )))

    /// 常规用法：取对象走默认的 ROT / ProgID，强类型转换与身份核对由 work 注入。
    let attach
        (info: GeneralAppInfo)
        (progIds: string list)
        (work: obj -> Result<'T, string>)
        : Result<'T, string> =
        attachWith info (defaultAcquire progIds) work
