/// AutoCAD 主机库：拿到一个 AutoCAD Application，并且能干净地起一个、用一次、关掉。
/// 系统里已有 AutoCAD 实例就用现成的（运行对象表里注册的那个），没有就新建一个。
///
/// 只负责 AutoCAD 自己的事：取/建实例、认自己是谁、等它冷启动就绪、干净地关掉自己起的实例。
/// 线程、进程账本、窗口 → 进程号、COM 底层这些三家通用的活儿都在 `OsHost` 里。
/// 注意：AutoCAD 冷启动是分钟级，GetOrCreateApplication() 会一直阻塞到自动化对象可用。
namespace AutoCADHost

open System
open System.Diagnostics
open System.Runtime.InteropServices
open System.Threading
open Autodesk.AutoCAD.Interop
open OsHost

[<RequireQualifiedAccess>]
module Host =

    /// 先认钉死的版本化 ProgID（R23 = AutoCAD 2020，与编译时的类型库同代），再退一步认通用 ProgID
    let private progIds = [ "AutoCAD.Application.23"; "AutoCAD.Application" ]

    /// 有实例就返回那个实例，没有就新建一个并返回新建的。
    /// 新建的实例窗口可见（AutoCAD 的 COM 服务器不带隐藏参数），退出由调用方负责。
    ///
    /// 这是**对外签约**用的：可以真的起一个进程（冷启动分钟级）。按地址回来（harness 内部交接）
    /// 别用它 —— 用 `TryGetRunningApplication`。
    let GetOrCreateApplication () : AcadApplication =
        match Com.running progIds with
        | Some running -> running :?> AcadApplication
        | None -> Com.create progIds :?> AcadApplication

    /// 只接现有的：ROT 里有就交出来，没有就是 None —— **绝不新建**。
    ///
    /// 「凭地址回到那台 host」走这条：地址说的人不在，就该老实报错，
    /// 不能因为它不在就顺手起一个（AutoCAD 冷启动还要分钟级等它就绪，更不能乱来）。
    let TryGetRunningApplication () : AcadApplication option =
        Com.running progIds |> Option.map(fun value -> value :?> AcadApplication)

    // ---------------- 黑板模式：按 info 挂回现成的那一只 ----------------

    /// 这只 AutoCAD 的进程号（AcadApplication 自己有 HWND，一问就到）
    let pidOf (app: AcadApplication) : int option =
        try
            Win32.pidOfWindow(nativeint app.HWND)
        with _ ->
            None

    /// 按插件抄下来的 info 精确挂回那一只 AutoCAD 实例。
    /// 只附着、不新建、不 Quit、不改应用状态 —— 冷启动分钟级，这里连碰都不会碰新建那条路。
    /// 返回的 AcadApplication 请在 OsHost.Sta.run 那条固定 STA 线程上用（与附着同公寓）。
    /// 文档级交接检查（换了图纸 / 采集后存过盘等）按用户要求推迟，暂不做；
    /// info 里的 DocName / DocPath 只是收下的数据，不参与判断。
    let GetApplication (info: AppInfo) : Result<AcadApplication, string> =
        AppInfo.attach info AppKind.AutoCAD progIds (fun raw ->
            try
                let app = raw :?> AcadApplication

                match pidOf app with
                | Some number when number <> info.Pid ->
                    Error(sprintf "ROT 名字指向的是 PID %d，不是黑板上的 %d" number info.Pid)
                | _ -> Ok app
            with ex ->
                Error(Errors.explain ex))

    // ---------------- 认自己是谁 ----------------
    // 地址里的指纹就是从这儿来的：哪个进程、锚定哪张图、那张图最后什么时候存的。

    /// 图纸有没有未保存改动（部分版本/接口上不一定有，取不到就当没有）
    let private docSaved (doc: obj) : bool option =
        try
            let prop = doc.GetType().GetProperty "Saved"

            if isNull prop then
                None
            else
                match prop.GetValue doc with
                | :? bool as value -> Some value
                | _ -> None
        with _ ->
            None

    /// 现场：进程号 + 锚定的图纸（当前图纸）+ 那个文件在磁盘上的最后写入时间 + 有没有未保存改动。
    /// 新图（还没存过盘）路径是 None、保存时间是 0。
    let identity (app: AcadApplication) : Identity =
        let pid = pidOf app |> Option.defaultValue 0

        let path, dirty =
            try
                let doc = app.ActiveDocument

                let dirty =
                    match docSaved doc with
                    | Some saved -> not saved
                    | None -> false

                Paths.canonical doc.FullName, dirty
            with _ ->
                // 一张图都没有（或还没就绪）：没有锚点，也就不谈保存时间
                None, false

        { Pid = pid
          StartedMs = Proc.startedMs pid |> Option.defaultValue 0L
          DocPath = path
          SavedMs = Paths.savedMs path
          Dirty = dirty }

    // ---------------- 收尾 ----------------
    // 为什么要有这几个：AutoCAD 只要还有「改过没存」的图纸，Quit() 就弹「是否保存 Drawing1」的模态框，
    // 调用方就卡死在那儿等人点按钮（自动化场景没人点）。所以退出前必须先把图纸关掉、改动丢掉。

    /// 关掉当前活动图纸并丢弃改动（Close(SaveChanges=false)），不弹保存对话框。
    ///
    /// 一定要用 Application.ActiveDocument 拿到的那个对象来关：
    /// Documents.Item(n) 返回的对象调 Close 会直接 E_FAIL(0x80004005)（AutoCAD 2020 实测），
    /// 关不掉 → Quit 时照样弹保存框。
    let CloseActiveDocumentWithoutSaving (app: AcadApplication) : bool =
        try
            app.ActiveDocument.Close(box false, Type.Missing)
            true
        with _ ->
            false

    /// 把实例里所有图纸关掉（改动全丢弃），返回关掉几张。
    /// 有关不掉的（比如用户正存盘、图纸被锁）就停下，不硬来；
    /// 关到最后一张后如果 AutoCAD 又自动开一张新的（有的配置会这样），也当停手，别死循环。
    let CloseAllDocumentsWithoutSaving (app: AcadApplication) : int =
        let mutable closed = 0
        let mutable keepGoing = true
        let mutable lastCount = Int32.MaxValue

        while keepGoing && closed < 32 do
            let remaining =
                try
                    app.Documents.Count
                with _ ->
                    0

            if remaining <= 0 || remaining >= lastCount then
                keepGoing <- false
            else
                lastCount <- remaining

                if CloseActiveDocumentWithoutSaving app then
                    closed <- closed + 1
                else
                    keepGoing <- false

        closed

    /// 先关掉所有图纸（丢弃改动），再 Quit —— 这样不会弹「是否保存」，调用方不会卡在模态框上。
    /// 只对「本程序自己起的实例」用；用户手上正开着图的实例别碰。
    let QuitWithoutSaving (app: AcadApplication) : unit =
        CloseAllDocumentsWithoutSaving app |> ignore
        app.Quit()

    /// 等刚起来的 AutoCAD 真的闲下来：GetAcadState().IsQuiescent 要连中 5 次（约 1.5 秒）再静置一会儿。
    /// 刚起来时它会短暂地报 true，紧接着动它就是 RPC_E_SERVERFAULT，甚至进程直接崩。
    let private waitReady (app: AcadApplication) : bool =
        let clock = Stopwatch.StartNew()
        let mutable streak = 0

        while streak < 5 && clock.ElapsedMilliseconds < 240000L do
            Thread.Sleep 300

            if
                (try
                    app.GetAcadState().IsQuiescent
                 with _ ->
                     false)
            then
                streak <- streak + 1
            else
                streak <- 0

        if streak >= 5 then
            Thread.Sleep 1500
            true
        else
            false

    // ---------------- 方式一：一次命令（exec / execQuery） ----------------
    // 语义照参考项目 AutoCADHost.Host.exec：**一次进程、一个实例、一段活**。
    // 改了一处：碰到「已有 AutoCAD 在跑」不弹框求人退出，直接返回 Error。

    /// 后台跑，不给用户弹窗口
    let private quieten (app: AcadApplication) : unit =
        (try
            app.Visible <- false
         with _ ->
             ())

    /// 放掉 RCW：进程外服务器要等客户端松开引用才肯真的退出
    let private release (target: obj) : unit =
        (try
            Marshal.FinalReleaseComObject target |> ignore
         with _ ->
             ())

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    /// 起一个新的 AutoCAD → 等它就绪 → 跑 work → 关掉退出。已有实例在跑就返回 Error。
    let private runOnce (work: AcadApplication -> 'T) : Result<'T, string> =
        Sta.here (fun () ->
            match Com.running progIds with
            | Some _ -> Error "本机已有 AutoCAD 在运行：一次性命令要求它先退干净，否则会碰你手上的图"
            | None ->
                let before = Proc.ids "acad"

                match
                    (try
                        Ok(Com.create progIds :?> AcadApplication)
                     with ex ->
                         Error(Errors.explain ex))
                with
                | Error message -> Error message
                | Ok app ->
                    let pids = Set.difference (Proc.ids "acad") before
                    quieten app

                    let outcome =
                        if not (waitReady app) then
                            Error "AutoCAD 240 秒内仍未就绪"
                        else
                            try
                                Ok(work app)
                            with ex ->
                                Error(Errors.explain ex)

                    (try
                        QuitWithoutSaving app
                     with _ ->
                         ())

                    release app
                    Proc.kill (Proc.waitGone "acad" pids 15000)
                    outcome)

    /// 一次性执行，只关心成败（异常原因以 Error 返回，不向上抛）。
    let exec (action: AcadApplication -> unit) : Result<unit, string> = runOnce action

    /// 一次性执行并取回一个值。
    let execQuery (query: AcadApplication -> 'T) : Result<'T, string> = runOnce query
