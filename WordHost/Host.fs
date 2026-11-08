/// Word 主机库：拿到一个 Word Application，并且能干净地起一个、用一次、关掉。
/// 系统里已有 Word 实例就用现成的（运行对象表里注册的那个），没有就新建一个。
///
/// 只负责 Word 自己的事：取/建实例、认自己是谁、干净地关掉自己起的实例。
/// 线程、进程账本、窗口 → 进程号、COM 底层这些三家通用的活儿都在 `OsHost` 里。
namespace WordHost

open System
open System.Runtime.InteropServices
open Microsoft.Office.Core
open Microsoft.Office.Interop.Word
open OsHost

[<RequireQualifiedAccess>]
module Host =

    /// Word 的 ProgID 不带版本后缀：2016/2019/365 都注册成同一个名字（类型库都是 16）
    [<Literal>]
    let private ProgId = "Word.Application"

    let private progIds = [ ProgId ]

    /// 有实例就返回那个实例，没有就新建一个并返回新建的。
    /// 新建的实例是隐藏的（Visible 由调用方按需打开），不自动退出。
    ///
    /// 这是**对外签约**用的：可以真的起一个进程。按地址回来（harness 内部交接）别用它 —— 用 `TryGetRunningApplication`。
    let GetOrCreateApplication () : Application =
        match Com.running progIds with
        | Some running -> running :?> Application
        | None -> Com.create progIds :?> Application

    /// 只接现有的：ROT 里有就交出来，没有就是 None —— **绝不新建**。
    ///
    /// 「凭地址回到那台 host」走这条：地址说的人不在，就该老实报错，
    /// 不能因为它不在就顺手变一个出来（那等于把别人的会话顶掉，还认错了人）。
    let TryGetRunningApplication () : Application option =
        Com.running progIds |> Option.map(fun value -> value :?> Application)

    // ---------------- 黑板模式：按 info 挂回现成的那一只 ----------------

    /// 这只 Word 主窗口的进程号。Word 的 Application 没有 Hwnd，只能从文档窗口往上找 OpusApp；
    /// 一份文档都没有时压根没有文档窗口 → None（调用方会退到别的办法）。
    let pidOf (app: Application) : int option =
        try
            Win32.pidOfWindow(Win32.root(nativeint app.ActiveWindow.Hwnd))
        with _ ->
            None

    /// 这家 host 自己的身份参数：原来 AppKind.image() / kindName() 里 Word 那一支搬到专用 Host。
    /// 调用方填 GeneralAppInfo 的 Identity 就用它（进程名 / 显示名 / 种类标记三件套）。
    let SelfIdentity : GeneralIdentity =
        GeneralIdentity.create ("WINWORD", "Word", "word")

    /// 按插件抄下来的 GeneralAppInfo 精确挂回那一只 Word 实例。
    /// 只附着、不新建、不 Quit、不改应用状态（连 DisplayAlerts 都不碰）—— 黑板的生命周期不归这儿管。
    /// 返回的 Application 请在 OsHost.Sta.run 那条固定 STA 线程上用（与附着同公寓）。
    /// 种类标记不符 / 信息过时 / ROT 里没有这个名字 / 认错进程，一律 Error。
    /// 文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求推迟，暂不做；
    /// 文档名 / 路径本来就不在 info 里（GeneralDocInfo 单列），认人只看进程身份。
    let GetApplication (info: GeneralAppInfo) : Result<Application, string> =
        GeneralAppInfo.attach info SelfIdentity.Tag progIds (fun raw ->
            try
                let app = raw :?> Application

                match pidOf app with
                | Some number when number <> info.Pid ->
                    Error(sprintf "获取到的 PID %d 与给定的 PID %d 不符" number info.Pid)
                | _ -> Ok app
            with ex ->
                Error(Errors.explain ex))

    // ---------------- 认自己是谁 ----------------
    // 地址里的指纹就是从这儿来的：哪个进程、锚定哪份文档、那份文档最后什么时候存的。

    /// 现场：进程号 + 锚定的文档（活动文档）+ 那份文档在磁盘上的最后写入时间 + 有没有未保存改动。
    /// 新文档（还没存过盘）路径是 None、保存时间是 0。
    let identity (app: Application) : Identity =
        let pid = pidOf app |> Option.defaultValue 0

        let path, dirty =
            try
                let doc = app.ActiveDocument
                let saved = try doc.Saved with _ -> true
                Paths.canonical doc.FullName, not saved
            with _ ->
                // 一份文档都没有（或还没就绪）：没有锚点，也就不谈保存时间
                None, false

        { Pid = pid
          StartedMs = Proc.startedMs pid |> Option.defaultValue 0L
          DocPath = path
          SavedMs = Paths.savedMs path
          Dirty = dirty }

    // ---------------- 收尾 ----------------

    /// 关文档 + 退出：Quit 带上「不保存」，Word 就不会弹「是否保存」。
    let QuitWithoutSaving (app: Application) : unit =
        let mutable saveChanges: obj = box WdSaveOptions.wdDoNotSaveChanges
        let mutable format: obj = Type.Missing
        let mutable route: obj = Type.Missing
        app.Quit(&saveChanges, &format, &route)

    // ---------------- 方式一：一次命令（exec / execQuery） ----------------
    // 语义照参考项目 WordHost.Host.exec：**一次进程、一个实例、一段活**。
    // 改了一处：参考项目碰到「已有 Word 在跑」是弹框请人先退出，库不该弹框，直接返回 Error。

    /// 转换专用：不启用宏。带 AutoOpen / Workbook_Open 的老文件会弹宏提示，
    /// 一样把 COM 调用挂住；ForceDisable 让宏一律不跑，判宏仍只看 HasVBProject。
    /// 返回原值，收尾 restore 回去 —— 即便挂回用户实例也不留痕。
    module Security =
        let apply (app: Application) : MsoAutomationSecurity option =
            let previous =
                try
                    Some app.AutomationSecurity
                with _ ->
                    None

            (try
                app.AutomationSecurity <- MsoAutomationSecurity.msoAutomationSecurityForceDisable
             with _ ->
                 ())

            previous

        let restore (app: Application) (previous: MsoAutomationSecurity option) : unit =
            match previous with
            | Some value ->
                (try
                    app.AutomationSecurity <- value
                 with _ ->
                     ())
            | None -> ()

    /// 后台跑，不给用户弹窗口；一切确认框都不弹，直接按默认走。
    /// 返回 AutomationSecurity 原值，runOnce 在活干完后恢复。
    let private quieten (app: Application) : MsoAutomationSecurity option =
        (try
            app.Visible <- false
         with _ ->
             ())

        (try
            app.DisplayAlerts <- WdAlertLevel.wdAlertsNone
         with _ ->
             ())

        // 格式确认框：老 .doc 常问「要不要转换格式」，没人点就挂住
        (try
            app.Options.ConfirmConversions <- false
         with _ ->
             ())

        Security.apply app

    /// 放掉 RCW：进程外服务器要等客户端松开引用才肯真的退出。
    /// 顺带把终结器撵一遍，不然要等下一次 GC 才松手。
    let private release (target: obj) : unit =
        (try
            Marshal.FinalReleaseComObject target |> ignore
         with _ ->
             ())

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    /// 起一个新的 Word → 跑 work → 关掉退出。已有实例在跑就返回 Error（绝不碰用户的实例）。
    let private runOnce (work: Application -> 'T) : Result<'T, string> =
        Sta.here (fun () ->
            match Com.running progIds with
            | Some _ -> Error "本机已有 Word 在运行：一次性命令要求它先退干净，否则会碰你手上的文档"
            | None ->
                let before = Proc.ids "WINWORD"

                match
                    (try
                        Ok(Com.create progIds :?> Application)
                     with ex ->
                         Error(Errors.explain ex))
                with
                | Error message -> Error message
                | Ok app ->
                    let pids = Set.difference (Proc.ids "WINWORD") before
                    let previousSecurity = quieten app

                    let outcome =
                        try
                            Ok(work app)
                        with ex ->
                            Error(Errors.explain ex)

                    // 另存完把宏安全级别恢复成起实例时的值，不留痕
                    Security.restore app previousSecurity

                    (try
                        QuitWithoutSaving app
                     with _ ->
                         ())

                    release app
                    Proc.kill (Proc.waitGone "WINWORD" pids 10000)
                    outcome)

    /// 一次性执行，只关心成败（异常原因以 Error 返回，不向上抛）。
    let exec (action: Application -> unit) : Result<unit, string> = runOnce action

    /// 一次性执行并取回一个值。
    let execQuery (query: Application -> 'T) : Result<'T, string> = runOnce query
