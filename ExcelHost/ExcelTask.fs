/// 线程规约（与 AutoCADTask.fs 同口径，下面全部适用）：COM 对象只在挂着 RCW 的那条 STA 线程上
/// 创建 / 使用 —— start 也在调用线程上同步执行；Async 版不替你切 STA，只把结果转成 Task。
/// 要排进固定 STA，套 `OsHost.Sta.runAsync(fun () -> ExcelHost.ExcelTask.xxx ...)`。
///
/// 这一份放**成品入口**：判定链（checkPid / findOnSta / unwrap）、取实例、认人、收尾
/// 全留在 Host.fs 里，这里只做「前置闸门同步判 + 排进固定 STA + Result 抛成 faulted」这一层，
/// 保持一眼能读完；唯一的例外是「起或取」—— 那段取/建逻辑就在本文件（`getOrCreateOnSta`），
/// 因为它的两个调用形态（同步排不进 STA 的、Task 的）都在这儿，放在 Host 只是多一跳。
module ExcelHost.ExcelTask

open System
open System.Threading.Tasks
open Microsoft.Office.Interop.Excel
open OsHost

/// 挂回活着的 Excel 的 COM Application，按 .NET 命名约定加 Async 后缀。
///
/// 只认现成对象：绝不 CoCreateInstance、绝不 GetOrCreateApplication、绝不 Quit ——
/// 认人靠 `Application.Hwnd` → PID 硬核对，对不上就是错误，不猜、不顺手新建。
///
/// checkPid 不合格直接交回一个已完成且 faulted 的 Task（连 STA 队列都不用排）；
/// 合格才把 Host.findOnSta 排进 OsHost.Sta.runAsync 那条**固定 STA 线程**，COM 对象仍只在那里创建；
/// 队列里若报 Error，Host.unwrap 抛的异常由 runAsync 的 TaskCompletionSource 置成 faulted，
/// await 时抛 InvalidOperationException，文案与同步版的 Error 一致。
///
/// await 之后要用这个 RCW，请包一层 OsHost.Sta.run。
let GetApplicationAsync (pid: int) : Task<Application> =
    match Host.checkPid pid with
    | Error message -> Task.FromException<Application>(InvalidOperationException(message) :> exn)
    | Ok() -> STA.runAsync(fun () -> Host.findOnSta pid |> Host.unwrap)

/// 有实例就交回那个实例，没有就新建一个（并补一张空白工作簿：一本都没有的自动化实例会自己退出）。
///
/// 做法长在这儿（2026-11-13 从 `Host.GetOrCreateApplication` 内联过来）：ROT 直取，拿不到才
/// `CoCreateInstance` —— 新进程由 COM 拉起，调用线程拿到的是代理。
/// 这是**对外签约**入口 —— 真会起进程。按地址回来（harness 内部交接）别用它，用 `GetApplicationAsync`
/// （那条绝不新建，人不在就报错）。
///
/// **必须已经在那条固定 STA 线程上**才调（`Session.openSession` 就在 `Sta.run` 里直接调它）；
/// 不在 STA 的调用方用下面的 `GetOrCreateApplicationAsync`。
let getOrCreateOnSta () : Application =
    match Com.running Host.progIds with
    | Some running -> running :?> Application
    | None ->
        let app = Com.create Host.progIds :?> Application

        (try
            if app.Workbooks.Count = 0 then
                app.Workbooks.Add() |> ignore
         with _ ->
             ())

        app

/// 同一件事的 Task 版：活儿排进 OsHost.Sta.runAsync 那条**固定 STA 线程**，
/// COM 对象仍只在那里创建；异常由 TaskCompletionSource 置成 faulted，await 时抛出。
/// await 之后要用这个 RCW，请包一层 OsHost.Sta.run。
let GetOrCreateApplicationAsync () : Task<Application> = STA.runAsync getOrCreateOnSta
