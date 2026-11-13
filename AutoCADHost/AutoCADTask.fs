/// 线程规约（与 SendCommand 一致，下面全部适用）：必须在挂着 RCW 的那条 STA 线程上调用，
/// start 也在调用线程上同步执行。Async 版不改变这条 —— 它只把结果转成 Async，不替你切 STA；
/// 要排进固定 STA，套 `OsHost.Sta.runAsync(fun () -> AutoCADHost.AutoCADTask.xxx ...)`。
module AutoCADHost.AutoCADTask
open System
open System.Threading.Tasks
open Autodesk.AutoCAD.Interop
open OsHost


/// 挂回活着的 AutoCAD 的 COM Application，按 .NET 命名约定加 Async 后缀。
/// checkPid 不合格直接交回一个已完成且 faulted 的 Task（连 STA 队列都不用排）；合格才把
/// findOnSta 排进 OsHost.Sta.runAsync 那条**固定 STA 线程**，COM 对象仍只在那里创建；
/// 队列里若报 Error，unwrap 抛的异常由 runAsync 的 TaskCompletionSource 置成 faulted。
/// await 之后要用这个 RCW，请包一层 OsHost.Sta.run。
let GetApplicationAsync (pid: int) : Task<IAcadApplication> =
    match Application.checkPid pid with
    | Error message -> Task.FromException<IAcadApplication>(InvalidOperationException(message) :> exn)
    | Ok() -> STA.runAsync(fun () -> Application.findOnSta pid |> Result.map fst |> Application.unwrap)

/// 任意启动动作，等 EndCommand；结果就是那条命令的名字。
let commandAsync
    (options: EventPair.Options)
    (app: IAcadApplication)
    (start: unit -> unit)
    : Task<string> =
    AcadBeginEnd.runWith options (AcadBeginEnd.command app) start


/// 一步到位：发一条命令，等它真的结束。
///
/// **SendCommand 返回 ≠ 命令跑完**。SendCommand 是同步返回的（命令已交给 AutoCAD），
/// EndCommand 要等命令真正结束后才到；这个 Task 只有收到 EndCommand 才完成，
/// 调用方才好接着发下一条。
///
/// 命令行串必须自己带换行收尾，否则最后一个词留在命令行上等输入，整只 acad 会一直「忙」
/// （后续调用 RPC_E_CALL_REJECTED）。
let sendCommandAsync
    (options: EventPair.Options)
    (doc: IAcadDocument)
    (commandLine: string)
    : Task<string> =
    commandAsync options doc.Application (fun () -> doc.SendCommand commandLine)

// ── modal：BeginModal → EndModal ───────────────────────────────

/// 等一个模态对话框关掉。modal 没有「启动动作」——开始由开对话框的一方触发，
/// 这里只等 EndModal（对话框关掉后 AutoCAD 才真正空下来）。
let modalAsync (options: EventPair.Options) (app: IAcadApplication) : Task<unit> =
    AcadBeginEnd.runWith options (AcadBeginEnd.modal app) ignore

// ── open：BeginOpen → EndOpen ─────────────────────────────────

/// start 里发 Documents.Open（或命令行 OPEN），等 EndOpen 才算真的开完；结果是图纸全名。
/// 比「撞忙就重试」可靠：Open 期间发什么都会被拒。
let openDrawingAsync
    (options: EventPair.Options)
    (app: IAcadApplication)
    (start: unit -> unit)
    : Task<string> =
    AcadBeginEnd.runWith options (AcadBeginEnd.openDrawing app) start


// ── 特殊输入：LISP ────────────────────────────────────────────
//
// LISP 走的是**应用级 / 文档级**事件：BeginLisp / EndLisp / LispCancelled（见 AcadBeginEnd.lisp），
// 没有 BeginCommand/EndCommand 那条命令事件 —— 除非代码里自己 (command "_.LINE" ...) 调起命令。
// 下面这一对分工：
//   · sendLispAsync      —— 只交出去，不等（调用方自己想办法确认，比如让 LISP 写结果文件）；
//   · runLispAsync 等    —— 等 EndLisp：整段求值真的结束了才返回。

/// 发一行 LISP 表达式，**不等它跑完**。
///
/// 返回只代表「已交给 AutoCAD」。要「跑完了再发下一条」，用下面的 runLispAsync：
/// 它订阅 EndLisp，等整段求值结束才返回。
let sendLispAsync (doc: IAcadDocument) (expression: string) : Task<unit> =
    try
        doc.SendCommand(expression.TrimEnd() + "\n")
        Task.FromResult(())
    with ex ->
        Task.FromException<unit>(ex)


/// 载入一个 .lsp 文件（不等它跑完）：路径给绝对路径（(load) 不认相对路径），反斜杠换成斜杠。
///
/// 两个坑：① AutoCAD 2020 的 AutoLISP 按本地 ANSI(GBK) 读 .lsp，文件里有中文就可能读坏；
/// ② SECURELOAD 非 0 时会拦路径 —— 纯 ASCII 且落在受信目录最省事。
let loadLispAsync (doc: IAcadDocument) (lispPath: string) : Task<unit> =
    sendLispAsync doc (sprintf "(load \"%s\")" (lispPath.Replace('\\', '/')))


// ── LISP：等整段求值结束（EndLisp） ────────────────────────────

/// 发一行 LISP，等**整段求值结束**才返回。
///
/// 判据是应用级的 EndLisp，不是某条命令的 EndCommand：SendCommand 送进去的整段表达式
/// 算一个整体 —— 里面 (command "_.LINE" ...) 的 Begin/EndCommand 先各自走完，最后才来 EndLisp。
/// 用户按 Esc 时代替 EndLisp 来的是 LispCancelled，这里交回 Cancelled，不冒充成功。
///
/// 只代表「求值结束」，不代表「成功」：AutoLISP 的求值错误不会回报给调用方
/// （命令行打印一行错误就过去了），要拿结论就让代码自己写结果文件。
///
/// 线程：SendCommand 与订阅都必须在挂着 RCW 的 STA 上（与 sendCommandAsync 一致）。
///
/// 默认策略认 BeginLisp（RequireBegin=true）：万一本机某版不派 BeginLisp（只在手敲时派），
/// 会一直等不到开始 → 超时；这时传 `{ EventPair.defaults with RequireBegin = false }` 即可。
let sendLispAndWaitAsync
    (options: EventPair.Options)
    (doc: IAcadDocument)
    (code: string)
    : Task<LispOutcome> =
    AcadBeginEnd.runWith
        options
        (AcadBeginEnd.lisp doc.Application)
        (fun () -> doc.SendCommand(code.TrimEnd() + "\n"))


/// sendLispAndWaitAsync 的默认策略版：认 BeginLisp、30 秒时限、不取消
let runLispAsync (doc: IAcadDocument) (code: string) : Task<LispOutcome> =
    sendLispAndWaitAsync EventPair.defaults doc code


/// (load "...") 的等待版：文件里的代码全部跑完（EndLisp）才算完
let runLispFileAsync (doc: IAcadDocument) (lispPath: string) : Task<LispOutcome> =
    runLispAsync doc (sprintf "(load \"%s\")" (lispPath.Replace('\\', '/')))
