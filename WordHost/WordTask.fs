/// 线程规约（与 ExcelTask.fs / AutoCADTask.fs 同口径，下面全部适用）：COM 对象只在挂着 RCW 的那条
/// STA 线程上创建 / 使用 —— Async 版不替你切 STA，只把结果转成 Task。
/// 要排进固定 STA，套 `OsHost.Sta.runAsync(fun () -> WordHost.WordTask.xxx ...)`。
///
/// 这一份只放**成品 Task 入口**：判定链（GeneralAppInfo.verify / Host.matchApp / findOnSta / unwrap）、
/// 取实例、认人、收尾全留在 Host.fs 里，这里只做「前置闸门同步判 + 排进固定 STA + Result 抛成 faulted」
/// 这一层，保持一眼能读完。
module WordHost.WordTask

open System
open System.Threading.Tasks
open Microsoft.Office.Interop.Word
open OsHost

/// 按黑板参数表（GeneralAppInfo）挂回活着的 Word 的 COM Application，按 .NET 命名约定加 Async 后缀。
///
/// 只认现成对象：绝不 CoCreateInstance、绝不 GetOrCreateApplication、绝不 Quit、不改应用状态 ——
/// 认人靠 ROT 名精确附着 + `ActiveWindow` 上溯 `OpusApp` 得到的 PID 硬核对，对不上就是错误，
/// 不猜、不顺手新建。
///
/// 前置闸门 GeneralAppInfo.verify 是**同步**做的、纯查进程表（PID>0、Image 非空、PID 存活、EXE、ROT 名后缀），
/// 不合格直接交回一个已完成且 faulted 的 Task，连 STA 队列都不用排；
/// 合格才把 Host.findOnSta 排进 OsHost.Sta.runAsync 那条**固定 STA 线程**，COM 对象仍只在那里创建；
/// 队列里若报 Error，Host.unwrap 抛的异常由 runAsync 的 TaskCompletionSource 置成 faulted，
/// await 时抛 InvalidOperationException，文案与同步版的 Error 一致。
///
/// 注意 Word 与 Excel 的差别：Word 的 Application 没有 Hwnd，PID 只能从文档窗口上溯 ——
/// 一份文档都没开时 pidOf 返回 None，此时**跳过 pid 核对**（认不出就不反驳），不是挂回失败。
///
/// await 之后要用这个 RCW，请包一层 OsHost.Sta.run。
let GetApplicationAsync (info: GeneralAppInfo) : Task<Application> =
    match GeneralAppInfo.verify info with
    | Error message -> Task.FromException<Application>(InvalidOperationException(message) :> exn)
    | Ok() -> Sta.runAsync(fun () -> Host.findOnSta info |> Host.unwrap)
