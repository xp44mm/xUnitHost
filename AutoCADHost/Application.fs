/// 挂回层：给一个活着的 acad 进程号，挂回它的 COM Application。**绝不新建、绝不 Quit、绝不改应用状态。**
///
/// 搬过来的判据很窄 —— 只搬挂回真正用到的东西：
///   · GetApplicationAsync
///   · unwrap / checkPid / findOnSta / matchApp / pidOfApp / hwndOf / safe / progIds  ← 它自己走的那条链
///   · 「acad」映像名：原文件用的是 Host.SelfImage，Host.fs 其余内容与这条链无关，只取这一个常量
/// 同一份原文件里的 rootsOnSta / describe / rootsOf / executeCoreAsync / Ready.Gate 都不在这条链上，**没搬**。
///
/// 依赖落点：Sta / Com / Proc / Win32 属于 OsHost（三家 host 共用的那一份），所以引 OsHost 项目。
///
/// 前置校验 checkPid 是**同步**做的 —— 不合格立刻交回一个已完成且 faulted 的 Task，
/// 连固定 STA 线程的队列都不排；合格才把 findOnSta 排进 Sta.runAsync，
/// Result.Error 由 unwrap 抛成异常，被 runAsync 的 TaskCompletionSource 置成 faulted。
/// await 之后要用这个 RCW，仍得包一层 OsHost.Sta.run（COM 对象只在 STA 上用）。
namespace AutoCADHost

open System
open System.Threading.Tasks
open Autodesk.AutoCAD.Interop
open OsHost

/// 挂回活着的 acad：只认现成对象，绝不新建、绝不 Quit、绝不改应用状态。
[<RequireQualifiedAccess>]
module Application =

    /// 原 Host.SelfImage 的最小搬法：这段只用到「acad 这个映像名」
    let private selfImage = "acad"

    /// 与编译时的类型库同代（R23 = AutoCAD 2020），再退一步认通用 ProgID
    let private progIds = [ "AutoCAD.Application.23"; "AutoCAD.Application" ]

    /// 取 COM 属性时忙 / 没就绪都可能抛，逐个兜住，绝不因为一个读数就整条挂回失败
    let private safe (get: unit -> 'T) (fallback: 'T) : 'T =
        try
            get ()
        with _ ->
            fallback

    let private hwndOf (app: IAcadApplication) : nativeint =
        safe (fun () -> nativeint app.HWND) 0n

    /// 这只应用对象到底属于哪个进程：AutoCAD 的 HWND 就是主框架窗口
    let private pidOfApp (app: IAcadApplication) : int option =
        Win32.pidOfWindow(hwndOf app)

    /// obj → IAcadApplication，并用 HWND → PID 硬核对。不是 AutoCAD / pid 对不上一律 None。
    let private matchApp (pid: int) (raw: obj) : IAcadApplication option =
        try
            let app = raw :?> IAcadApplication

            if obj.ReferenceEquals(app, null) then
                None
            elif pidOfApp app = Some pid then
                Some app
            else
                None
        with _ ->
            None

    /// 只认现成对象：绝不 CoCreateInstance，绝不 GetOrCreateApplication，绝不 Quit。
    /// 必须在 Sta.run / Sta.runAsync 那条固定 STA 线程上调用（由公开入口包好）。
    let findOnSta (pid: int) : Result<IAcadApplication * string, string> =
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
    /// 这是纯查进程表，不碰 COM —— 所以自检能离线跑到这一层。
    let checkPid (pid: int) : Result<unit, string> =
        if pid <= 0 then
            Error "PID 必须大于 0：不猜、不新建"
        elif not (Process.ids selfImage |> Set.contains pid) then
            Error(sprintf "PID %d 不是正在运行的 %s 进程：拒绝挂回，绝不新建" pid selfImage)
        else
            Ok()

    /// 异步版用 Task 自己的错误处理：挂不回就是 faulted Task，await 时抛
    /// InvalidOperationException，Message 与同步版的 Error 文案一致；外层不再包 Result。
    let unwrap (result: Result<'T, string>) : 'T =
        match result with
        | Ok value -> value
        | Error message -> raise (InvalidOperationException(message) :> exn)

