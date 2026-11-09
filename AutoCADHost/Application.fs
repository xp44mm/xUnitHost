/// AutoCAD 的 Application 模块：给一个系统里正活着的 acad 进程号，挂回它的 COM Application，
/// 并把从 Application 直接伸出来的根对象一次取齐。**不新建、不 Quit、不改应用状态**。
///
/// 关于「根对象」的两层边界（本模块最重要的权衡，别混为一谈）：
///   · COM 层（跨进程拿得到）：AcadApplication 是唯一真正的根，它直属的
///     Documents / ActiveDocument / Preferences / MenuGroups / GetAcadState() 是第一批子根。
///     这些是进程外服务器在 ROT 里暴露的自动化对象，只要按 PID 认准那一只就能挂回去。
///   · .NET 进程内层（跨进程拿不到）：HostApplicationServices.Current、
///     Autodesk.AutoCAD.ApplicationServices.Core.Application / Application、DocumentManager
///     这些是 AutoCAD 进程**内部**的 CLR 静态单例，COM 只是跨进程代理，看不见它们。
///     想读只能让代码跑在同一个进程里：项目里的 AcadMgdHost 插件就是这条路，
///     它的 XH_DB 命令已经在读 HostApplicationServices.WorkingDatabase。
///     本模块不假装能取到，也不为此拉起进程或注入 DLL；边界见 Application.boundary。
///
/// 认人靠硬判据：AcadApplication.HWND → Win32.pidOfWindow → 与给定 PID 比对。
/// 先用版本化 ProgID 直取（不枚举 ROT，照项目一贯口径）；PID 对不上才枚举 ROT 全部条目
/// 逐条核身份 —— 多实例机器上也绝不「抓到第一只就当是它」。
namespace AutoCADHost

open System
open System.Threading.Tasks
open Autodesk.AutoCAD.Interop
open OsHost

/// 一只活着的 AutoCAD 实例：COM 根对象 + 现场读数。
/// COM 对象请始终在 OsHost.Sta.run 那条固定 STA 线程上使用（与挂回时同公寓）；
/// 要一份能随便带走的文字，用 Application.describe / Application.describeByPid。
type ApplicationRoots =
    { /// 唯一真正的 COM 根对象
      Application: AcadApplication
      /// 打开的图纸集合
      Documents: AcadDocuments option
      /// 当前图纸；一张都没开时 None
      ActiveDocument: AcadDocument option
      /// 选项 / 系统配置入口
      Preferences: AcadPreferences option
      /// 菜单与自定义组（CUI）
      MenuGroups: AcadMenuGroups option
      /// 空闲状态对象（GetAcadState）
      State: AcadState option
      /// 以下全是纯数据，不碰 COM，跨线程断言 / 贴报告都安全
      Pid: int
      Hwnd: nativeint
      HwndPid: int option
      Version: string
      Caption: string
      DocumentCount: int
      MenuGroupCount: int
      Quiescent: bool option
      ActiveDocumentName: string
      ActiveDocumentFullName: string
      /// 命中来源：ProgID 直取 / ROT 登记名
      RotName: string }

[<RequireQualifiedAccess>]
module Application =

    /// 与编译时的类型库同代（R23 = AutoCAD 2020），再退一步认通用 ProgID
    let private progIds = [ "AutoCAD.Application.23"; "AutoCAD.Application" ]

    /// 边界说明：哪些根对象这一侧能取、哪些必须靠进程内插件。贴报告 / 写文档都用它。
    let boundary: string list =
        [ "COM 根：AcadApplication → Documents / ActiveDocument / Preferences / MenuGroups / AcadState"
          "进程内 .NET 根：HostApplicationServices.Current、Core.Application、DocumentManager"
          "进程内根跨进程取不到：COM 只是代理；读它们走 AcadMgdHost 的 XH_DB（已读 WorkingDatabase）" ]

    /// 取 COM 属性时忙 / 没就绪都可能抛，逐个兜住，绝不因为一个读数就整条挂回失败
    let private safe (get: unit -> 'T) (fallback: 'T) : 'T =
        try
            get ()
        with _ ->
            fallback

    let private hwndOf (app: AcadApplication) : nativeint =
        safe (fun () -> nativeint app.HWND) 0n

    /// 这只应用对象到底属于哪个进程：AutoCAD 的 HWND 就是主框架窗口
    let private pidOfApp (app: AcadApplication) : int option =
        Win32.pidOfWindow(hwndOf app)

    /// obj → AcadApplication，并用 HWND → PID 硬核对。不是 AutoCAD / pid 对不上一律 None。
    let private matchApp (pid: int) (raw: obj) : AcadApplication option =
        try
            let app = raw :?> AcadApplication

            if obj.ReferenceEquals(app, null) then
                None
            elif pidOfApp app = Some pid then
                Some app
            else
                None
        with _ ->
            None

    /// 只认现成对象：绝不 CoCreateInstance，绝不 GetOrCreateApplication，绝不 Quit。
    /// 必须在 Sta.run 那条固定 STA 线程上调用（由公开入口包好）。
    let private findOnSta (pid: int) : Result<AcadApplication * string, string> =
        let direct =
            match Com.running progIds with
            | Some raw ->
                matchApp pid raw
                |> Option.map(fun app -> app, "ProgID 直取 " + String.Join(" / ", progIds))
            | None -> None

        match direct with
        | Some hit -> Ok hit
        | None ->
            let entries = Com.runningEntries ()

            let hit =
                entries
                |> List.tryPick(fun (name, raw) ->
                    matchApp pid raw |> Option.map(fun app -> app, "ROT " + name))

            match hit with
            | Some value -> Ok value
            | None ->
                Error(
                    sprintf
                        "ROT 里没有 PID %d 的 AutoCAD 自动化对象（ProgID 直取没对上，另扫了 %d 条登记项）：可能是 accoreconsole 无 COM / 正在冷启动 / 尚未注册；绝不新建"
                        pid
                        entries.Length
                )

    /// 前置闸门：PID 有效，且确实是正在运行的 acad 进程。任何一条不满足都拒绝挂回。
    let private checkPid (pid: int) : Result<unit, string> =
        if pid <= 0 then
            Error "PID 必须大于 0：不猜、不新建"
        elif not (Proc.ids Host.SelfImage |> Set.contains pid) then
            Error(sprintf "PID %d 不是正在运行的 %s 进程：拒绝挂回，绝不新建" pid Host.SelfImage)
        else
            Ok()

    /// 从已挂回的 Application 取根对象（只读，不开图、不改设置、不 Quit）。
    /// COM 调用排在固定 STA 线程上，与 byPid 那条是同一条。
    let private rootsOnSta (app: AcadApplication, rot: string) : Result<ApplicationRoots, string> =
        try
            let hwnd = hwndOf app
            let docs = safe (fun () -> Some app.Documents) None

            let active =
                safe (fun () -> match app.ActiveDocument with null -> None | doc -> Some doc) None

            let prefs = safe (fun () -> Some app.Preferences) None
            let groups = safe (fun () -> Some app.MenuGroups) None
            let state = safe (fun () -> Some(app.GetAcadState())) None

            let docName =
                active
                |> Option.map(fun doc -> safe (fun () -> doc.Name) "")
                |> Option.defaultValue ""

            let docFull =
                active
                |> Option.map(fun doc -> safe (fun () -> doc.FullName) "")
                |> Option.defaultValue ""

            Ok
                { Application = app
                  Documents = docs
                  ActiveDocument = active
                  Preferences = prefs
                  MenuGroups = groups
                  State = state
                  Pid = Win32.pidOfWindow hwnd |> Option.defaultValue 0
                  Hwnd = hwnd
                  HwndPid = Win32.pidOfWindow hwnd
                  Version = safe (fun () -> app.Version) ""
                  Caption = safe (fun () -> app.Caption) ""
                  DocumentCount = safe (fun () -> app.Documents.Count) 0
                  MenuGroupCount = safe (fun () -> app.MenuGroups.Count) 0
                  Quiescent = state |> Option.bind(fun s -> safe (fun () -> Some s.IsQuiescent) None)
                  ActiveDocumentName = docName
                  ActiveDocumentFullName = docFull
                  RotName = rot }
        with ex ->
            Error(Errors.explain ex)

    /// 主函数之一：给一个活着的 acad PID，挂回它的 COM Application。
    /// 不新建、不 Quit、不改任何状态；返回的 RCW 请在 OsHost.Sta.run 上用。
    let byPid (pid: int) : Result<AcadApplication, string> =
        match checkPid pid with
        | Error message -> Error message
        | Ok() -> Sta.run(fun () -> findOnSta pid |> Result.map fst)

    /// 异步版用 Task 自己的错误处理：挂不回就是 faulted Task，await 时抛
    /// InvalidOperationException，Message 与同步 byPid 的 Error 文案一致；外层不再包 Result。
    let private unwrap (result: Result<'T, string>) : 'T =
        match result with
        | Ok value -> value
        | Error message -> raise (InvalidOperationException(message) :> exn)

    /// byPid 的异步版（按 .NET 命名约定加 Async 后缀）：功能相同，但不阻塞调用方。
    /// checkPid 不合格直接交回一个已完成且 faulted 的 Task（连 STA 队列都不用排）；合格才把
    /// findOnSta 排进 OsHost.Sta.runAsync 那条**固定 STA 线程**，COM 对象仍只在那里创建；
    /// 队列里若报 Error，unwrap 抛的异常由 runAsync 的 TaskCompletionSource 置成 faulted。
    /// await 之后要用这个 RCW，请包一层 OsHost.Sta.run；要纯数据用下面的
    /// GetApplicationWithRootsAsync，它的现场字段跨线程直接读是安全的。
    let GetApplicationAsync (pid: int) : Task<AcadApplication> =
        match checkPid pid with
        | Error message -> Task.FromException<AcadApplication>(InvalidOperationException(message) :> exn)
        | Ok() -> Sta.runAsync(fun () -> findOnSta pid |> Result.map fst |> unwrap)

    /// 「把动作送进 AutoCAD」的入口：pid + 一段动作 → 一个 Task。
    /// 与 GetApplicationAsync 不同：这个 Task 要等「挂回 → 订阅 ready 通知 → 每一动作都等到
    /// AutoCAD 的完成信号 → 退订」全做完才完成，调用方 await 到才算结束。
    ///
    /// 动作通过 gate.SubmitAsync 逐步提交：每一步是一个独立 Task，AutoCAD 通知 ready 之前
    /// 不会发下一步 —— 正是「打开一个 dwg 后 task.completed() 才能继续」的形状。
    /// 也可以整段只 SubmitAsync 一次，那就是「打开命令和后续动作打包成一个动作、算一个 task」。
    /// 订阅挂在 AutoCAD 的 COM 事件（EndOpen / NewDrawing / EndCommand / EndModal）上，
    /// 放行还看 AcadState.IsQuiescent 连中 + 静置；不用 Busy.retry 去撞 RPC_E_CALL_REJECTED。
    /// 失败一律是 faulted Task（await 抛异常），不包 Result。
    let executeAsync (pid: int) (action: Ready.Gate -> AcadApplication -> Task<'T>) : Task<'T> =
        task {
            let! app = GetApplicationAsync pid
            use gate = Ready.subscribe app
            return! action gate app
        }

    /// 只提交一次的简写：整段动作写在一个同步 lambda 里，对外仍是一个 Task，
    /// 内部同样先等 ready、跑完再等 ready；完成时刻由 AutoCAD 的通知决定，不是 COM 返回。
    let executeOnceAsync (pid: int) (action: AcadApplication -> 'T) : Task<'T> =
        executeAsync pid (fun gate _ -> gate.SubmitAsync action)

    /// 配套的一步到位异步入口：Application + 全部 COM 根对象 + 现场读数。
    /// 挂回与读值都在同一条 STA 线程上做完，免得异步链被同步的 rootsOf 截断。
    /// 失败同样是 faulted Task，await 时抛异常。
    let GetApplicationWithRootsAsync (pid: int) : Task<ApplicationRoots> =
        match checkPid pid with
        | Error message -> Task.FromException<ApplicationRoots>(InvalidOperationException(message) :> exn)
        | Ok() -> Sta.runAsync(fun () -> findOnSta pid |> Result.bind rootsOnSta |> unwrap)

    /// 主函数之二：一步拿到 Application + 全部 COM 根对象 + 现场读数。
    /// 这是「给 PID，取 Application 和其他根对象」的常用入口。
    let byPidWithRoots (pid: int) : Result<ApplicationRoots, string> =
        match checkPid pid with
        | Error message -> Error message
        | Ok() -> Sta.run(fun () -> findOnSta pid |> Result.bind rootsOnSta)

    /// 已经拿着 Application 时，单独把根对象补齐。
    let rootsOf (app: AcadApplication) : Result<ApplicationRoots, string> =
        Sta.run(fun () -> rootsOnSta (app, "(调用方已有的 Application)"))

    /// 把 roots 读成 (标签, 值) 的行：全是纯数据，不碰 COM，任何线程都能调。
    let describe (roots: ApplicationRoots) : (string * string) list =
        let yes = function
            | Some _ -> "取到"
            | None -> "取不到"

        let show (value: bool option) =
            match value with
            | Some flag -> string flag
            | None -> "取不到"

        [ "PID", string roots.Pid
          "HWND", sprintf "%d → PID %A" (int64 roots.Hwnd) roots.HwndPid
          "命中", (if String.IsNullOrWhiteSpace roots.RotName then "-" else roots.RotName)
          "版本", roots.Version
          "标题", roots.Caption
          "App", "AcadApplication（COM 唯一根对象）"
          "Docs", sprintf "AcadDocuments %s，%d 份图纸" (yes roots.Documents) roots.DocumentCount
          "ActiveDoc",
          (match roots.ActiveDocument with
           | Some _ -> sprintf "「%s」 %s" roots.ActiveDocumentName roots.ActiveDocumentFullName
           | None -> "（没有活动图纸）")
          "Prefs", sprintf "AcadPreferences %s" (yes roots.Preferences)
          "Groups", sprintf "AcadMenuGroups %s，%d 组" (yes roots.MenuGroups) roots.MenuGroupCount
          "State", sprintf "AcadState %s，IsQuiescent=%s" (yes roots.State) (show roots.Quiescent) ]

    /// 一步取回可贴进报告的文字（挂回 + 取根对象 + 读值都在同一条 STA 线程上做完）。
    let describeByPid (pid: int) : Result<(string * string) list, string> =
        match checkPid pid with
        | Error message -> Error message
        | Ok() ->
            Sta.run(fun () ->
                findOnSta pid
                |> Result.bind rootsOnSta
                |> Result.map describe)
