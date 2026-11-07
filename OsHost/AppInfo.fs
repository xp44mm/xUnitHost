/// 黑板模式：插件把「当前实例」的现场抄到黑板上，消费方按这段信息精确挂回那一只实例。
///
/// 这一层只认**结构化参数表 AppInfo**：调用方（插件 / 宿主程序）按字段填报，
/// 库不认任何展示文本、不解析剪贴板原文 —— 文本排版一改只该改调用方，不该改库。
/// 把插件文本转成 AppInfo 的参考写法放在测试项目 xUnitHost/BlackboardText.fs，**不进生产库**。
/// 这一层不认任何 Office 类型，也绝不创建 / 销毁实例：认不出、对不上、已过时，一律交回 Error。
/// 四家 host 的 GetApplication(info) 共用这里的「认人 + 校验 + 按名取对象」。
namespace OsHost

open System
open System.Globalization

/// 四家 COM host 的种类（与 XUnitHost.Kind 互不相干，这儿是 OsHost 的通用种类）
[<RequireQualifiedAccess>]
type AppKind =
    | Word
    | Excel
    | PowerPoint
    | AutoCAD

    /// 进程名（Proc.ids 用的 image，不带 .exe）
    member this.image () : string =
        match this with
        | AppKind.Word -> "WINWORD"
        | AppKind.Excel -> "EXCEL"
        | AppKind.PowerPoint -> "POWERPNT"
        | AppKind.AutoCAD -> "acad"

    /// 报错 / 报告里给人看的名字
    member this.kindName () : string =
        match this with
        | AppKind.Word -> "Word"
        | AppKind.Excel -> "Excel"
        | AppKind.PowerPoint -> "PowerPoint"
        | AppKind.AutoCAD -> "AutoCAD"

/// 调用方要填的那张参数表：一条实例信息的全部字段，都是结构化值，没有一行展示文本。
/// 直接构造记录即可（`{ Kind = ...; Pid = ...; ... }`），也可以用静态构造成员
/// `AppInfo.create(...)` 起个最小集再 with 补齐。
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
      DocPath: string option }

    /// 静态构造成员：调用方按最小必填集填报结构化数据。文档名 / 路径按记录更新补上，例如：
    ///   { (AppInfo.create (AppKind.Word, 27232, "!WordVSTO.Instance.27232", 0n, "", captured))
    ///       with DocName = Some "文档1"; DocPath = None }
    /// 采集时间由填报方给出（Unix 毫秒），库不替你猜，也不从任何文本里读。
    static member create
        (kind: AppKind, pid: int, rot: string, hwnd: nativeint, exePath: string, capturedMs: int64)
        : AppInfo =
        { Kind = kind
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
            (this.Kind.kindName())
            this.Pid
            (if String.IsNullOrWhiteSpace this.Rot then "-" else this.Rot)
            this.Hwnd
            (Clock.show this.CapturedMs)

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module AppInfo =

    let private firstError (checks: (bool * string) list) : string option =
        checks |> List.tryPick(fun (isBad, message) -> if isBad then Some message else None)

    /// 附着之前的「不过时」核对：进程还在、没被复用、窗口还属于它、EXE 路径对得上、
    /// ROT 名里的进程号与 PID 行一致。任何一条不通过就拒接，绝不改用 GetActiveObject 猜。
    let verify (info: AppInfo) : Result<unit, string> =
        let alive = Proc.ids (info.Kind.image()) |> Set.contains info.Pid
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
                    (info.Kind.image()) )
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
                    (info.Kind.kindName())
                    (kind.kindName())
            )
        else
            Ok()

    /// 挂上之后再核一次「黑板上那份文档还是不是现在这份」：
    /// 名字 / 路径不符，或文件在采集之后被保存过，都说明中途有别的东西动过 —— 拒接。
    /// 参数表里没填文档行就跳过这一层（老版插件只抄了五行）。
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
