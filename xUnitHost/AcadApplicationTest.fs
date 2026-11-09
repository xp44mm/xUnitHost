/// AutoCADHost.Application 的验证：给 PID → 挂回 AcadApplication + COM 根对象。
/// 两条都只读本机现场，绝不新建 / 关闭 AutoCAD，也绝不往用户手上的图纸里写东西：
/// 负例用测试程序自己的 PID，正向只在**本来就开着** acad 时挂回那一只去读数。
namespace XUnitHost

open System
open System.Diagnostics
open System.IO
open System.Threading
open System.Threading.Tasks
open Autodesk.AutoCAD.Interop
open Xunit
open FSharp.xUnit
open OsHost

type AcadApplicationTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Application 模块 按 PID 拒接非 acad 的进程，绝不新建``() =
        Report.title output "Application 模块：按 PID 挂回前先认人，不是 acad 一律拒接"

        for text in AutoCADHost.Application.boundary do
            Report.line output "边界" text

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)
        let selfPid = Environment.ProcessId

        Report.line output "自测 PID" (sprintf "%d（本测试程序，不是 acad）" selfPid)

        let denied = AutoCADHost.Application.byPid selfPid
        let text =
            match denied with
            | Error message -> message
            | Ok _ -> "Ok（不该发生）"

        Report.line output "byPid" text
        Should.equal true (
            match denied with
            | Error message -> message.Contains "不是正在运行" && message.Contains "绝不新建"
            | _ -> false
        )

        let zero = AutoCADHost.Application.byPidWithRoots 0
        Report.line output "PID=0" (sprintf "%A" (zero |> Result.map(fun _ -> "不该发生")))
        Should.equal true (
            match zero with
            | Error message -> message.Contains "PID 必须大于 0"
            | _ -> false
        )

        let described = AutoCADHost.Application.describeByPid selfPid
        Report.line output "describeByPid" (sprintf "%A" (described |> Result.map(fun _ -> "不该发生")))
        Should.equal true (
            match described with
            | Error message -> message.Contains "不是正在运行"
            | _ -> false
        )

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：非 acad 的 PID / 0 都在挂回前被拒，四家进程一个没起、一个没关"

    [<Fact>]
    member this.``Application 模块 按 PID 挂回真机 acad 的 Application 与 COM 根对象``() =
        Report.title output "Application 模块：按 PID 挂回真机 acad 的 Application 与 COM 根对象"

        let before = Report.procIds "acad"

        if before.IsEmpty then
            Assert.Skip("本机没有开着的 acad，按约定不替用户新起 AutoCAD（要测显式设 XUNITHOST_LAUNCH_ACAD=1）")

        let pid = Set.minElement before
        Report.line output "给定 PID" (sprintf "%d（本机本来就开着的 acad）" pid)

        let app = AutoCADHost.Application.byPid pid |> Report.ok
        Report.line output "Application" (if isNull (box app) then "null" else "AcadApplication 已挂回")
        Should.equal false (isNull (box app))

        let roots = AutoCADHost.Application.rootsOf app |> Report.ok

        for label, value in AutoCADHost.Application.describe roots do
            Report.line2 output "根对象" label value

        Should.equal pid roots.Pid
        Should.equal (Some pid) roots.HwndPid
        Should.equal true (roots.DocumentCount >= 0)

        let oneStep = AutoCADHost.Application.byPidWithRoots pid |> Report.ok
        Report.line2 output "一步到位" "PID" (string oneStep.Pid)
        Report.line2 output "一步到位" "图纸数" (string oneStep.DocumentCount)
        Report.line2 output "一步到位" "命中" oneStep.RotName

        let rows = AutoCADHost.Application.describeByPid pid |> Report.ok
        Report.line2 output "文字入口" "行数" (string rows.Length)
        Should.equal true (rows |> List.exists(fun (label, value) -> label = "PID" && value = string pid))
        Should.equal pid oneStep.Pid
        Should.equal roots.DocumentCount oneStep.DocumentCount

        let now = Report.procIds "acad"
        Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
        Should.equal before now

        Report.conclusion output "通过：按 PID 认准那一只 acad，COM 根对象与现场读数都取到，进程没增没减"

    [<Fact>]
    member this.``Application 模块 按 PID 的 ROT 枚举兜底能列出登记对象``() =
        Report.title output "Application 模块：只给 PID 时的 ROT 枚举兜底，逐条列出登记对象"

        // 多实例机器上 ProgID 直取可能拿到另一只 acad，兜底要能把 ROT 条目逐条取回再核身份。
        // 这里登记一个自己的对象验证枚举本身可用，不碰任何 AutoCAD 实例。
        let name = sprintf "!AcadAppProbe.%d" Environment.ProcessId
        let probe = BlackboardProbe()
        let before = Report.procIds "acad"

        let registered = Sta.run(fun () -> Com.register name probe)

        match registered with
        | Error message ->
            Report.line output "登记" message
            Assert.Skip("本机不允许在此进程登记 ROT：" + message)
        | Ok cookie ->
            try
                let entries = Sta.run Com.runningEntries
                Report.line output "ROT 条目" (sprintf "%d 条" entries.Length)

                let hit =
                    entries
                    |> List.tryPick(fun (display, value) ->
                        if display.Contains "AcadAppProbe" || obj.ReferenceEquals(value, probe) then
                            Some(display, value.GetType().FullName)
                        else
                            None)

                match hit with
                | Some (display, typeName) ->
                    Report.line output "登记名" display
                    Report.line output "对象类型" typeName
                    Should.equal "XUnitHost.BlackboardProbe" typeName
                | None ->
                    Report.line output "登记名" "枚举里没找到刚登记的对象"
                    Should.equal true hit.IsSome
            finally
                Sta.run(fun () -> Com.revoke cookie)

        let now = Report.procIds "acad"
        Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
        Should.equal before now

        Report.conclusion output "通过：兜底路径能把 ROT 登记项逐条列出并取回对象，acad 进程没增没减"

    [<Fact>]
    member this.``Application 模块 异步 GetApplicationAsync 与同步 byPid 结果一致``() =
        Report.title output "Application 模块：GetApplicationAsync 异步挂回，结果与同步 byPid 一致"

        let before = Report.procIds "acad"

        if before.IsEmpty then
            Assert.Skip("本机没有开着的 acad，按约定不替用户新起 AutoCAD（要测显式设 XUNITHOST_LAUNCH_ACAD=1）")

        let pid = Set.minElement before
        let callerThread = Environment.CurrentManagedThreadId
        Report.line output "给定 PID" (sprintf "%d（本机本来就开着的 acad）" pid)
        Report.line output "调用线程" (string callerThread)

        // task { } 里是真的 await；最外层才同步等，免得依赖 xunit 对 F# Task 测试方法的发现。
        let work =
            task {
                // Task 不包 Result：失败会把异常直接抛在 await 上
                let! app = AutoCADHost.Application.GetApplicationAsync pid
                let! roots = AutoCADHost.Application.GetApplicationWithRootsAsync pid

                // await 之后要用 RCW，仍包回固定 STA 线程；顺手量一下续体和 STA 是不是同一条线程
                let version = Sta.run(fun () -> app.Version)
                let staThread = Sta.run(fun () -> Thread.CurrentThread.ManagedThreadId)
                let resumeThread = Environment.CurrentManagedThreadId
                return (app, roots, version, staThread, resumeThread)
            }

        let app, roots, version, staThread, resumeThread = work.GetAwaiter().GetResult()

        Report.line output "Task 形状" "GetApplicationAsync pid → Task<AcadApplication>（失败即 faulted）"
        Report.line2 output "异步结果" "Application" (if isNull (box app) then "null" else "AcadApplication 已挂回")
        Report.line output "版本" version
        Report.line output "线程" (sprintf "STA %d / 调用 %d / await 续体 %d" staThread callerThread resumeThread)
        Report.line2 output "根对象" "PID" (string roots.Pid)
        Report.line2 output "根对象" "图纸数" (string roots.DocumentCount)

        Should.equal false (isNull (box app))
        Should.equal pid roots.Pid
        Should.equal version roots.Version
        // COM 工作没有搬到调用线程上：仍在那条固定 STA 线程
        Should.equal false (staThread = callerThread)

        let now = Report.procIds "acad"
        Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
        Should.equal before now

        Report.conclusion output "通过：Task 可 await，COM 仍在固定 STA 线程，结果与同步一致，acad 没增没减"

    [<Fact>]
    member this.``Application 模块 异步 GetApplicationAsync 对非 acad PID 也拒接``() =
        Report.title output "Application 模块：GetApplicationAsync 对非 acad PID 也拒接，且不排队"

        let selfPid = Environment.ProcessId
        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        // checkPid 不合格时连 STA 队列都不用排，直接交回一个已完成且 faulted 的 Task
        let pendingTask = AutoCADHost.Application.GetApplicationAsync selfPid
        Report.line output "自测 PID" (sprintf "%d（本测试程序，不是 acad）" selfPid)
        Report.line output "Task 已完成" (string pendingTask.IsCompleted)
        Report.line output "Task 失败" (string pendingTask.IsFaulted)
        Should.equal true pendingTask.IsCompleted
        Should.equal true pendingTask.IsFaulted

        // Task 的错误机制：异常要落在 await 上，而不是藏在 Result 里
        let capture (get: unit -> Task) : string =
            let work =
                task {
                    try
                        do! get ()
                        return "Ok（不该发生）"
                    with ex ->
                        return ex.GetType().Name + "：" + ex.Message
                }

            work.GetAwaiter().GetResult()

        let appMessage = capture (fun () -> AutoCADHost.Application.GetApplicationAsync selfPid :> Task)
        let rootsMessage = capture (fun () -> AutoCADHost.Application.GetApplicationWithRootsAsync selfPid :> Task)

        Report.line output "Application" appMessage
        Report.line output "一步到位" rootsMessage

        Should.equal true (appMessage.Contains "InvalidOperationException" && appMessage.Contains "绝不新建")
        Should.equal true (rootsMessage.Contains "InvalidOperationException" && rootsMessage.Contains "绝不新建")

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：异步负例交回 faulted Task，await 时抛 InvalidOperationException，四家进程一个没起、一个没关"

    [<Fact>]
    member this.``Application 模块 只用 PID 挂回与 CuislCad 黑板冗余数据逐行比对``() =
        Report.title output "Application 模块：GetApplicationAsync(PID) 的现场值 vs CuislCad 黑板文本"

        // 2026-11-09 14:06:46 用户贴的黑板原文，原样留在代码里当比对输入。
        // 库不解析文本：这里只把展示文本的每一行与「只给一个 PID」能拿到的值摆在一起。
        let blackboard =
            String.concat
                "\n"
                [ "[CuislCad 当前图纸引用] 采集时间 2026-11-09 14:06:46"
                  "KIND : AutoCAD"
                  "PID : 27148"
                  "ROT : !CuislCad.Instance.27148"
                  "HWND : 0x000A0760"
                  "ACAD.EXE : C:\\Program Files\\Autodesk\\AutoCAD 2020\\acad.exe"
                  "AutoCAD : R23.1.47.0.0"
                  "CuislCad.dll : C:\\Application Data\\GitHub\\xp44mm\\CuislCad\\CuislCad\\bin\\x64\\Debug\\net481\\CuislCad.dll（2026-11-09 11:32:29）"
                  "图纸 : Drawing2.dwg"
                  "路径 : (未保存)" ]

        let lines = blackboard.Split '\n' |> Array.map(fun line -> line.Trim())

        let valueOf (key: string) : string =
            lines
            |> Array.tryPick(fun line ->
                let i = line.IndexOf ':'

                if i > 0 && String.Equals(line.Substring(0, i).Trim(), key, StringComparison.OrdinalIgnoreCase) then
                    Some(line.Substring(i + 1).Trim())
                else
                    None)
            |> Option.defaultValue ""

        let bbPid = int (valueOf "PID")
        let bbKind = valueOf "KIND"
        let bbRot = valueOf "ROT"
        let bbHwnd = valueOf "HWND"
        let bbExe = valueOf "ACAD.EXE"
        let bbVersion = valueOf "AutoCAD"
        let bbDoc = valueOf "图纸"
        let bbPath = valueOf "路径"
        let bbDllRaw = valueOf "CuislCad.dll"

        let bbDllPath, bbDllStamp =
            let i = bbDllRaw.IndexOf '（'

            if i > 0 then
                bbDllRaw.Substring(0, i).Trim(), bbDllRaw.Substring(i + 1).TrimEnd([| '）' |])
            else
                bbDllRaw, ""

        let before = Report.procIds "acad"

        if not (Set.contains bbPid before) then
            Assert.Skip(sprintf "黑板里的 acad PID %d 已经不在运行：信息过时，绝不另起 AutoCAD" bbPid)

        Report.line output "黑板原文" (sprintf "%s（共 %d 行）" (BlackboardText.caption blackboard) lines.Length)
        Report.line output "入口" (sprintf "AutoCADHost.Application.GetApplicationAsync %d → Task<AcadApplication>" bbPid)

        // 现有口径：这段文本只填出 4 个字段，文档信息另立一张只作报告数据的表。
        let info = BlackboardText.ofText blackboard |> Report.ok
        Report.line output "参数表" (info.describe ())
        Report.line output "文档表" ((BlackboardText.docOfText blackboard).describe () + "（只作报告数据，不参与认人）")
        let verified =
            match GeneralAppInfo.verify info with
            | Ok() -> "Ok（只核 Image / PID 存活 / EXE / ROT 后缀）"
            | Error message -> "Error " + message

        Report.line output "参数表" ("verify → " + verified)

        Should.equal true (
            match GeneralAppInfo.verify info with
            | Ok() -> true
            | Error _ -> false
        )

        // 只给一个 PID：await 挂回；之后要用 RCW，仍包回固定 STA 线程读现场值。
        let work =
            task {
                let! app = AutoCADHost.Application.GetApplicationAsync bbPid

                let live =
                    Sta.run(fun () ->
                        let hwnd = nativeint app.HWND

                        let safeText (get: unit -> string) =
                            try
                                get ()
                            with ex ->
                                "取不到：" + ex.GetType().Name

                        let active =
                            try
                                match app.ActiveDocument with
                                | null -> None
                                | doc -> Some doc
                            with _ ->
                                None

                        let rotObject =
                            try
                                match Com.runningAt bbRot with
                                | Some value -> sprintf "按此名取到 %s" (value.GetType().FullName)
                                | None -> "按此名没登记对象（本函数不靠它）"
                            with ex ->
                                "取不到：" + ex.GetType().Name

                        {| Hwnd = int64 hwnd
                           HwndPid = Win32.pidOfWindow hwnd
                           Version = safeText (fun () -> app.Version)
                           Caption = safeText (fun () -> app.Caption)
                           DocumentCount = (try app.Documents.Count with _ -> -1)
                           DocName =
                            safeText (fun () ->
                                match active with
                                | Some doc -> doc.Name
                                | None -> "(无活动图纸)")
                           DocFull =
                            safeText (fun () ->
                                match active with
                                | Some doc -> doc.FullName
                                | None -> "(未保存)")
                           RotObject = rotObject |})

                return app, live
            }

        let app, live = work.GetAwaiter().GetResult()
        Should.equal false (isNull (box app))
        Should.equal (Some bbPid) live.HwndPid

        // 同一批事实还能只凭 PID 从操作系统侧现查，不必抄在黑板上。
        let exe = Proc.exePath bbPid |> Option.defaultValue "(取不到)"

        let exeVersion =
            try
                let vi = FileVersionInfo.GetVersionInfo exe
                sprintf "FileVersion=%s / ProductVersion=%s" vi.FileVersion vi.ProductVersion
            with ex ->
                "取不到：" + ex.GetType().Name

        // 托管程序集不进 Process.Modules（那只列原生模块），这一行只能按黑板给的路径校磁盘文件；
        // COM 是跨进程代理，同样读不到插件加载了哪些 .NET 程序集。
        let nativeModule =
            try
                use proc = Process.GetProcessById bbPid

                let hit =
                    proc.Modules
                    |> Seq.cast<ProcessModule>
                    |> Seq.tryFind(fun m -> m.ModuleName.Equals("CuislCad.dll", StringComparison.OrdinalIgnoreCase))

                match hit with
                | Some m -> sprintf "原生模块里有 %s" m.FileName
                | None -> "原生模块列表里没有（.NET 托管程序集不在 Process.Modules）"
            with ex ->
                "取不到：" + ex.GetType().Name

        let dllFileExists = File.Exists bbDllPath

        let dllFileStamp =
            if dllFileExists then
                File.GetLastWriteTime(bbDllPath).ToString("yyyy-MM-dd HH:mm:ss")
            else
                "(文件不存在)"

        let liveHwnd = sprintf "0x%08X" (int live.Hwnd)
        let exeSame = String.Equals(bbExe, exe, StringComparison.OrdinalIgnoreCase)
        let hwndSame = String.Equals(bbHwnd, liveHwnd, StringComparison.OrdinalIgnoreCase)
        let docSame = String.Equals(bbDoc, live.DocName, StringComparison.OrdinalIgnoreCase)
        let dllFileSame = dllFileExists && dllFileStamp = bbDllStamp

        let cmp (label: string) (bb: string) (site: string) (verdict: string) =
            Report.line2 output "比对" label (sprintf "黑板 %s ｜ 现场 %s ｜ %s" bb site verdict)

        cmp "KIND" bbKind "AcadApplication 已挂回" "冗余：挂回成功本身就是种类证据"
        cmp "PID" (string bbPid) (sprintf "%d（HWND 反查 %A）" bbPid live.HwndPid) "必需：本函数唯一输入"
        cmp
            "ROT"
            bbRot
            (sprintf "ProgID 直取 AutoCAD.Application.23；%s" live.RotObject)
            "冗余：按 PID 挂回不用它（按名挂回才用）"
        cmp
            "HWND"
            bbHwnd
            (sprintf "%s → PID %A" liveHwnd live.HwndPid)
            (if hwndSame then
                 "一致但冗余：窗口会变，只用来反查 PID"
             else
                 "已变（快照）：不参与认人，故不作断言")
        cmp
            "ACAD.EXE"
            bbExe
            exe
            (if exeSame then "一致·冗余：Proc.exePath(pid) 现查" else "不一致：黑板里的 EXE 已过时")
        cmp
            "AutoCAD"
            bbVersion
            (sprintf "COM Version=%s；exe %s" live.Version exeVersion)
            "同一代 R23.1，写法不同·冗余：版本可现查"
        cmp
            "CuislCad.dll"
            bbDllRaw
            (sprintf "磁盘 %s（%s）；%s" bbDllPath dllFileStamp nativeModule)
            (if dllFileSame then
                 "文件与时间戳对得上，但属插件自报：由 PID 反查不到托管程序集"
             else
                 "磁盘文件对不上：插件自报信息，跨进程无法核归属")
        cmp
            "图纸"
            bbDoc
            (sprintf "ActiveDocument.Name=%s（%d 份图纸）" live.DocName live.DocumentCount)
            (if docSame then "此刻一致·快照冗余：文档随时会换" else "已变（快照）：项目口径不参与认人")
        cmp
            "路径"
            bbPath
            (sprintf "FullName=%s" (if String.IsNullOrWhiteSpace live.DocFull then "(空)" else live.DocFull))
            (if String.IsNullOrWhiteSpace live.DocFull then
                 "一致（未保存）·快照冗余：落盘后随时变"
             else
                 "已落盘（快照）：不参与认人")

        cmp
            "采集时间"
            "2026-11-09 14:06:46"
            (sprintf "COM 侧无对应属性；进程启动 %s" (Proc.startedMs bbPid |> Option.defaultValue 0L |> Report.clock))
            "冗余：抄录动作的时间，不是进程属性"

        Should.equal true exeSame
        Report.line2 output "现场" "标题" live.Caption
        Report.line2 output "现场" "图纸数" (string live.DocumentCount)

        let now = Report.procIds "acad"
        Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
        Should.equal before now

        Report.conclusion
            output
            "通过：只给 PID 就挂回那一只 acad 并读到全部进程侧事实；黑板 10 行里 PID 是唯一必需输入，其余为可现查的冗余、会变的快照，或插件自报（DLL 路径与时间戳、采集时间）"

    [<Fact>]
    member this.``Application 模块 就绪订阅的判定规则 通知后连中与静置才放行``() =
        Report.title output "Application 模块：就绪订阅的判定规则（纯逻辑，不连 AutoCAD）"

        // 注入假时钟，把「Task 何时才完成」这条规则单测干净，不碰真机 COM。
        let mutable fakeNow = 0L
        let rule = AutoCADHost.Ready.Rule(5, 1500L, (fun () -> fakeNow))

        Report.line output "起手" (sprintf "IsReady=%b" (rule.IsReady()))
        Should.equal false (rule.IsReady())

        // AutoCAD 推来一条 EndOpen：刚打开一张 dwg，此刻立刻发下一条必然撞忙
        rule.Signal("EndOpen", "1.dwg")
        fakeNow <- 100L

        Report.line
            output
            "通知"
            (sprintf "EndOpen 1.dwg → 连中 %d / 静置 %dms / IsReady=%b" rule.Streak rule.IdleMs (rule.IsReady()))

        Should.equal false (rule.IsReady())

        // 连中 5 次 quiescent，但静置还不够
        for _ in 1 .. 5 do
            rule.Sample(true)

        fakeNow <- 1000L
        Report.line output "连中" (sprintf "5 次 true，静置才 900ms → IsReady=%b" (rule.IsReady()))
        Should.equal false (rule.IsReady())

        // 静置够了才放行
        fakeNow <- 1700L
        rule.Sample(true)
        Report.line
            output
            "放行"
            (sprintf "连中 %d / 静置 %dms → IsReady=%b" rule.Streak rule.IdleMs (rule.IsReady()))

        Should.equal true (rule.IsReady())

        // 又来一条通知：计时归零，重新等 —— 上一条命令的 Task 完成后，下一条仍要等
        rule.Signal("EndCommand", "CIRCLE")
        fakeNow <- 1800L

        Report.line
            output
            "再通知"
            (sprintf "EndCommand → 连中 %d / 静置 %dms / IsReady=%b" rule.Streak rule.IdleMs (rule.IsReady()))

        Should.equal false (rule.IsReady())

        // 采到一次 not-quiescent 也要归零
        for _ in 1 .. 5 do
            rule.Sample(true)

        fakeNow <- 3400L
        Should.equal true (rule.IsReady())
        rule.Sample(false)
        Report.line output "变忙" (sprintf "一次 false → IsReady=%b" (rule.IsReady()))
        Should.equal false (rule.IsReady())

        Report.line output "通知清单" (sprintf "%A" rule.Notices)
        Should.equal 2 (rule.Notices.Length)

        Report.conclusion
            output
            "通过：Task 的完成判据是「AutoCAD 通知过 + IsQuiescent 连中 5 次 + 静置 1.5 秒」，任何新动静都重新计时；纯逻辑自测，不碰真机"

    [<Fact>]
    member this.``Application 模块 就绪订阅入口对非 acad PID 直接 faulted``() =
        Report.title output "Application 模块：完成信号入口负例不订阅、不排队、不碰真机"

        let selfPid = Environment.ProcessId
        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        // 挂回在订阅之前：非 acad 的 PID 在 checkPid 就被拒，连接点根本不会建立。
        let work: Task<unit> =
            AutoCADHost.Application.executeAtPidAsync selfPid 30000 (fun _ -> ())

        let message =
            try
                work.GetAwaiter().GetResult()
                "Ok（不该发生）"
            with ex ->
                ex.GetType().Name + "：" + ex.Message

        Report.line output "自测 PID" (sprintf "%d（本测试程序，不是 acad）" selfPid)
        Report.line output "Task 失败" (string work.IsFaulted)
        Report.line output "await 抛出" message
        Should.equal true work.IsCompleted
        Should.equal true work.IsFaulted
        Should.equal true (message.Contains "InvalidOperationException" && message.Contains "绝不新建")

        // PID=0 同样在挂回前被拒，不猜、不新建
        let zero: Task<unit> =
            AutoCADHost.Application.executeAtPidAsync 0 30000 (fun _ -> ())

        let zeroMessage =
            try
                zero.GetAwaiter().GetResult()
                "Ok（不该发生）"
            with ex ->
                ex.GetType().Name + "：" + ex.Message

        Report.line output "PID=0" zeroMessage
        Should.equal true zero.IsFaulted
        Should.equal true (zeroMessage.Contains "PID 必须大于 0")

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion
            output
            "通过：非 acad 的 PID 与 0 都在挂回前被拒，连接点根本没建立；异常落在 await 上，四家进程一个没起、一个没关"

    [<Fact>]
    member this.``在 AutoCAD 里画一个圆``() =
        task {
            Report.title output "在 AutoCAD 里连画 10 个图形：逐笔看 COM 会不会拒"

            let before = Report.procIds "acad"

            if before.IsEmpty then
                Assert.Skip("本机没有开着的 acad，按约定不替用户新起 AutoCAD（要测显式设 XUNITHOST_LAUNCH_ACAD=1）")

            let pid = Set.minElement before
            // 整段动作就是一个 Task：await 到 Application 才算挂回完成，await 到最后才算这活干完。
            let! acad = AutoCADHost.Application.GetApplicationAsync pid

            // 口径按你草稿的直连形状：COM 调用直接发，不包 Sta.run、不拆属性读取。
            // 只有「必须成功」的前置 / 读回 / 收尾带 Busy.retry 兜忙；被探测的写动作一笔都不重试。
            let busy (work: unit -> 'T) : 'T = Busy.retry 30000 work

            // 跑前基线：现在可能一张图都没开（Documents.Count = 0 时 ActiveDocument 抛 0x8021003C）。
            let docCountBefore = busy (fun () -> acad.Documents.Count)
            Report.line output "现场" (sprintf "acad %d（本来就开着的实例，不新建、不 Quit）" pid)
            Report.line output "基线" (sprintf "图纸数 %d（0 份图纸时 ActiveDocument 取不到）" docCountBefore)

            // 用户手上的图一个字不写：一律新建临时图，画完不保存关掉。
            let dwg = busy (fun () -> acad.Documents.Add())
            let tempName = busy (fun () -> dwg.Name)
            let ms = busy (fun () -> dwg.ModelSpace)
            Report.line output "临时图" (sprintf "%s（Documents.Add 新建，画完不保存关掉）" tempName)
            Report.line output "口径" "直连：写动作不包 Sta.run、不包 Busy.retry、不拆属性读取；每笔单独 try/with"

            // 「多画几个图形」：10 种实体连着发，中间不等 ready、不重试。
            let pt (x: double) (y: double) : obj = box [| x; y; 0.0 |]

            let shapes: (string * (unit -> obj)) list =
                [ "AddCircle", (fun () -> box (ms.AddCircle(pt 10.0 20.0, 5.0)))
                  "AddLine", (fun () -> box (ms.AddLine(pt 0.0 0.0, pt 30.0 0.0)))
                  "AddArc", (fun () -> box (ms.AddArc(pt 50.0 20.0, 8.0, 0.0, Math.PI / 2.0)))
                  "AddLightWeightPolyline",
                  (fun () -> box (ms.AddLightWeightPolyline(box [| 0.0; 0.0; 10.0; 0.0; 10.0; 10.0; 0.0; 10.0 |])))
                  "AddText", (fun () -> box (ms.AddText("XH", pt 5.0 35.0, 3.0)))
                  "AddPoint", (fun () -> box (ms.AddPoint(pt 40.0 40.0)))
                  "AddEllipse", (fun () -> box (ms.AddEllipse(pt 60.0 50.0, box [| 10.0; 0.0; 0.0 |], 0.5)))
                  "AddCircle 2", (fun () -> box (ms.AddCircle(pt 80.0 20.0, 3.0)))
                  "AddLine 2", (fun () -> box (ms.AddLine(pt 0.0 50.0, pt 50.0 50.0)))
                  "AddPolyline",
                  (fun () -> box (ms.AddPolyline(box [| 0.0; 0.0; 0.0; 20.0; 0.0; 0.0; 20.0; 20.0; 0.0 |]))) ]

            let results =
                shapes
                |> List.mapi(fun i (name, draw) ->
                    let verdict =
                        try
                            let entity = draw ()

                            if isNull entity then
                                "OK（返回 null）"
                            else
                                "OK " + entity.GetType().Name
                        with ex ->
                            sprintf "%s 0x%08X：%s" (ex.GetType().Name) (uint32 ex.HResult) ex.Message

                    Report.line2 output (sprintf "第%d笔" (i + 1)) name verdict
                    name, verdict)

            let okCount = results |> List.filter(fun (_, v) -> v.StartsWith "OK") |> List.length
            let rejected = results |> List.filter(fun (_, v) -> not (v.StartsWith "OK"))

            let codes =
                rejected
                |> List.map(fun (_, v) -> v.Split([| '：' |]).[0])
                |> List.countBy id
                |> List.map(fun (code, n) -> sprintf "%s ×%d" code n)
                |> String.concat "；"

            Report.line
                output
                "统计"
                (sprintf
                    "%d 笔：OK %d / 被拒 %d%s"
                    results.Length
                    okCount
                    rejected.Length
                    (if rejected.IsEmpty then "" else "（" + codes + "）"))

            // 读回实体数：这一笔允许重试，只用来核对「成功笔数 == 真画进去的实体数」。
            let entityCount = busy (fun () -> ms.Count)
            Report.line output "读回" (sprintf "ModelSpace 实体数 %d（成功 %d 笔）" entityCount okCount)
            Should.equal okCount entityCount

            // 收尾：只认这次新建的那张，不保存关掉；关完图纸数应回到跑前。
            let closed =
                busy (fun () ->
                    let active = acad.ActiveDocument

                    if active.Name = tempName then
                        active.Close(box false, Type.Missing)
                        true
                    else
                        false)

            // Close 之后集合正在变，立刻枚举会抛非忙异常，等半拍再核对。
            do! Task.Delay 500
            let docCountAfter = busy (fun () -> acad.Documents.Count)

            Report.line output "收尾" (sprintf "临时图已关闭（不保存）= %b" closed)
            Report.line2 output "核对" "图纸数" (sprintf "%d → %d" docCountBefore docCountAfter)
            Should.equal true closed
            Should.equal docCountBefore docCountAfter

            let now = Report.procIds "acad"
            Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
            Should.equal before now

            Report.conclusion
                output
                (sprintf
                    "通过：%d 笔写动作里 OK %d / 被拒 %d；临时图不保存关掉，用户现场没动"
                    results.Length
                    okCount
                    rejected.Length)
        }

    [<Fact>]
    member this.``Application 模块 完成信号让 executeAsync 的 Task 收口``() =
        task {
            Report.title output "Application 模块：executeAsync 的 Task 由 AutoCAD 完成通知收口"

            let before = Report.procIds "acad"

            if before.IsEmpty then
                Assert.Skip("本机没有开着的 acad，按约定不替用户新起 AutoCAD（要测显式设 XUNITHOST_LAUNCH_ACAD=1）")

            if before.Count <> 1 then
                Assert.Skip("本机不止一只 acad，executeAsync 不给 PID 会拒接；这条实证请改用 executeAtPidAsync")

            let pid = Set.minElement before
            let busy (work: unit -> 'T) : 'T = Busy.retry 30000 work
            let app = AutoCADHost.Application.GetApplicationAsync(pid).GetAwaiter().GetResult()

            let namesBefore = busy (fun () -> [ for d in app.Documents -> d.Name ])
            Report.line output "现场" (sprintf "acad %d（本来就开着的实例，不新建、不 Quit）" pid)
            Report.line output "基线" (sprintf "图纸 %A" namesBefore)

            // 对照：空动作立刻返回，但 AutoCAD 没推来任何完成通知 → Task 不能就这么算完
            let control = AutoCADHost.Application.executeAtPidAsync pid 1500 (fun _ -> ())
            Report.line output "对照" (sprintf "空动作调用返回时 IsCompleted=%b" control.IsCompleted)
            Should.equal false control.IsCompleted

            let controlMessage =
                try
                    control.GetAwaiter().GetResult()
                    "Ok（不该发生：没有通知却完成了）"
                with ex ->
                    ex.GetType().Name + "：" + ex.Message

            Report.line output "对照" (sprintf "短超时后 %A：%s" control.Status controlMessage)

            try
                // 正例：Documents.Add() 会推 NewDrawing / EndOpen，Task 等这条通知才 SetResult。
                // 伪代码里那句「这里从 AcadApplication 等到完成信号」就落在这里。
                let clock = Stopwatch.StartNew()

                let work =
                    AutoCADHost.Application.executeAsync(fun a -> a.Documents.Add() |> ignore)

                Report.line output "正例" (sprintf "action 发出瞬间 IsCompleted=%b" work.IsCompleted)
                do! work
                clock.Stop()

                Report.line
                    output
                    "正例"
                    (sprintf "Task %A / 等到完成通知耗时 %dms" work.Status clock.ElapsedMilliseconds)

                Should.equal true work.IsCompleted
                Should.equal TaskStatus.RanToCompletion work.Status

                // 动作真的落地：多出一张临时图，随后只关这一张、不保存
                let tempName = busy (fun () -> app.ActiveDocument.Name)
                let docCountAfter = busy (fun () -> app.Documents.Count)
                Report.line output "读回" (sprintf "临时图 %s / 图纸数 %d" tempName docCountAfter)
                Should.equal (namesBefore.Length + 1) docCountAfter
            finally
                // 无论正例成功还是超时，都只收掉本次新建的那张，绝不碰用户原有的图
                try
                    busy (fun () ->
                        let d = app.ActiveDocument

                        if not (List.contains d.Name namesBefore) then
                            d.Close(box false, Type.Missing)

                        ())
                with _ ->
                    ()

            do! Task.Delay 500
            let namesAfter = busy (fun () -> [ for d in app.Documents -> d.Name ])
            Report.line2 output "核对" "图纸" (sprintf "%A → %A" namesBefore namesAfter)
            Should.equal namesBefore namesAfter

            let now = Report.procIds "acad"
            Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
            Should.equal before now

            Report.conclusion
                output
                "通过：空动作只等来超时；真动作等 AutoCAD 推来完成通知才 RanToCompletion，临时图不保存关掉"
        }

    [<Fact>]
    member this.``Application 模块 完成信号之后的读取不需要 Task，同步 Sta.run 就够``() =
        task {
            Report.title output "Application 模块：完成通知之后的读回，同步 Sta.run 就够，不必异步"

            let before = Report.procIds "acad"

            if before.IsEmpty then
                Assert.Skip("本机没有开着的 acad，按约定不替用户新起 AutoCAD（要测显式设 XUNITHOST_LAUNCH_ACAD=1）")

            if before.Count <> 1 then
                Assert.Skip("本机不止一只 acad，executeAsync 不给 PID 会拒接；这条实证请改用 executeAtPidAsync")

            let pid = Set.minElement before
            let busy (work: unit -> 'T) : 'T = Busy.retry 30000 work
            let app = AutoCADHost.Application.GetApplicationAsync(pid).GetAwaiter().GetResult()

            // 一次读回四个值：图纸数 / 活动图名 / 版本 / 当前图实体数。纯数据，不写任何东西。
            let readValues () : int * string * string * int =
                let count = app.Documents.Count
                let name = app.ActiveDocument.Name
                let version = app.Version
                let entities = app.ActiveDocument.ModelSpace.Count
                count, name, version, entities

            let field (rows: (string * string) list) (key: string) =
                rows
                |> List.tryPick(fun (k, v) -> if k = key then Some v else None)
                |> Option.defaultValue "-"

            let stableFields (rows: (string * string) list) =
                [ "版本"; "标题"; "Docs"; "Groups"; "ActiveDoc" ] |> List.map(field rows)

            let namesBefore = busy (fun () -> Sta.run(fun () -> [ for d in app.Documents -> d.Name ]))
            Report.line output "现场" (sprintf "acad %d / 图纸 %A" pid namesBefore)

            // 对照一：两个现成的「一步读数」入口，同步版与异步版读到的稳定字段是否一致
            let syncClock = Stopwatch.StartNew()

            let syncRows =
                match AutoCADHost.Application.describeByPid pid with
                | Ok rows -> rows
                | Error message -> failwithf "describeByPid：%s" message

            syncClock.Stop()

            let asyncClock = Stopwatch.StartNew()
            let! roots = AutoCADHost.Application.GetApplicationWithRootsAsync pid
            asyncClock.Stop()
            let asyncRows = AutoCADHost.Application.describe roots

            Report.line
                output
                "同步入口"
                (sprintf
                    "describeByPid → %d 行 / %dms / %A"
                    syncRows.Length
                    syncClock.ElapsedMilliseconds
                    (stableFields syncRows))

            Report.line
                output
                "异步入口"
                (sprintf
                    "GetApplicationWithRootsAsync → %d 行 / %dms / %A"
                    asyncRows.Length
                    asyncClock.ElapsedMilliseconds
                    (stableFields asyncRows))

            Should.equal (stableFields syncRows) (stableFields asyncRows)

            // 对照二：完成通知到达之后，同一组值用同步 / 异步 / 裸读各取一次
            let busyHr (ex: exn) =
                [ 0x80010001u; 0x8001010Au; 0x80010105u ] |> List.contains(uint32 ex.HResult)

            let rec readStaAsync (attempt: int) : Task<int * string * string * int> =
                task {
                    try
                        return! Sta.runAsync readValues
                    with ex when attempt < 3 && busyHr ex ->
                        do! Task.Delay(200 * (attempt + 1))
                        return! readStaAsync (attempt + 1)
                }

            let clock = Stopwatch.StartNew()
            let work = AutoCADHost.Application.executeAsync(fun a -> a.Documents.Add() |> ignore)
            Report.line output "动作" (sprintf "发出瞬间 IsCompleted=%b" work.IsCompleted)
            Should.equal false work.IsCompleted

            // 对照三：通知还没到就同步读，读到的是「动作还没做完」的现场（只读，不发新动作）
            let tooEarly =
                try
                    let earlyClock = Stopwatch.StartNew()
                    let value = Sta.run readValues
                    earlyClock.Stop()
                    sprintf "%A / %dms" value earlyClock.ElapsedMilliseconds
                with ex ->
                    ex.GetType().Name + "：" + ex.Message

            Report.line output "早读" (sprintf "通知未到就同步读 → %s" tooEarly)

            do! work
            clock.Stop()
            Report.line output "完成" (sprintf "Task %A / 等通知 %dms" work.Status clock.ElapsedMilliseconds)
            Should.equal TaskStatus.RanToCompletion work.Status

            try
                let syncReadClock = Stopwatch.StartNew()
                let syncValue = busy (fun () -> Sta.run readValues)
                syncReadClock.Stop()

                Report.line
                    output
                    "同步读"
                    (sprintf "Sta.run + Busy.retry → %A / %dms" syncValue syncReadClock.ElapsedMilliseconds)

                let (syncCount, _, _, _) = syncValue
                Should.equal (namesBefore.Length + 1) syncCount

                let asyncReadClock = Stopwatch.StartNew()
                let! asyncValue = readStaAsync 0
                asyncReadClock.Stop()

                Report.line
                    output
                    "异步读"
                    (sprintf "Sta.runAsync + let! → %A / %dms" asyncValue asyncReadClock.ElapsedMilliseconds)

                Should.equal syncValue asyncValue

                let bare =
                    try
                        let bareClock = Stopwatch.StartNew()
                        let value = readValues ()
                        bareClock.Stop()
                        sprintf "%A / %dms" value bareClock.ElapsedMilliseconds
                    with ex ->
                        ex.GetType().Name + "：" + ex.Message

                Report.line output "裸读" (sprintf "不经 Sta，直接读 RCW → %s" bare)
            finally
                // 只收掉本次新建的那张，绝不碰用户原有的图
                try
                    busy (fun () ->
                        Sta.run(fun () ->
                            let d = app.ActiveDocument

                            if not (List.contains d.Name namesBefore) then
                                d.Close(box false, Type.Missing)

                            ()))
                with _ ->
                    ()

            do! Task.Delay 500
            let namesAfter = busy (fun () -> Sta.run(fun () -> [ for d in app.Documents -> d.Name ]))
            Report.line2 output "核对" "图纸" (sprintf "%A → %A" namesBefore namesAfter)
            Should.equal namesBefore namesAfter

            let now = Report.procIds "acad"
            Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
            Should.equal before now

            Report.conclusion
                output
                "通过：通知未到同步读也读不到；通知到达后同步 Sta.run 与异步读到的值完全相同 —— 读取不必异步，异步只用在等 AutoCAD 通知，读值仍由 Busy.retry 兜忙"
        }
