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
        Report.title output "Application 模块：executeAsync 负例不订阅、不排队、不碰真机"

        let selfPid = Environment.ProcessId
        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        // 挂回在订阅之前：非 acad 的 PID 在 GetApplicationAsync 就被拒，
        // 连 Ready.subscribe 都不会被调用，更不会往用户实例上挂事件。
        let work: Task<int> =
            AutoCADHost.Application.executeAsync selfPid (fun _ _ -> task { return 1 })

        let message =
            try
                work.GetAwaiter().GetResult() |> ignore
                "Ok（不该发生）"
            with ex ->
                ex.GetType().Name + "：" + ex.Message

        Report.line output "自测 PID" (sprintf "%d（本测试程序，不是 acad）" selfPid)
        Report.line output "Task 失败" (string work.IsFaulted)
        Report.line output "await 抛出" message
        Should.equal true work.IsCompleted
        Should.equal true work.IsFaulted
        Should.equal true (message.Contains "InvalidOperationException" && message.Contains "绝不新建")

        // 简写入口同样如此
        let once: Task<int> = AutoCADHost.Application.executeOnceAsync selfPid (fun _ -> 1)

        let onceMessage =
            try
                once.GetAwaiter().GetResult() |> ignore
                "Ok（不该发生）"
            with ex ->
                ex.GetType().Name + "：" + ex.Message

        Report.line output "一次提交" onceMessage
        Should.equal true once.IsFaulted
        Should.equal true (onceMessage.Contains "绝不新建")

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion
            output
            "通过：非 acad 的 PID 在挂回前就被拒，订阅根本没建立；异常落在 await 上，四家进程一个没起、一个没关"

    [<Fact>]
    member this.``Application 模块用 Task 包裹的动作在 AutoCAD 里画一个圆``() =
        task {
            Report.title output "Application 模块：一个 Task 里的「挂回 → 画圆 → 读回 → 收尾」"

            let before = Report.procIds "acad"

            if before.IsEmpty then
                Assert.Skip("本机没有开着的 acad，按约定不替用户新起 AutoCAD")

            let pid = Set.minElement before
            Report.line output "现场" (sprintf "acad %d（本来就开着的实例，不新建、不 Quit）" pid)

            // 整段动作就是一个 Task：await 到 Application 才算挂回完成，await 到最后才算这活干完。
            let! acad = AutoCADHost.Application.GetApplicationAsync pid

            // COM 只在库里那条固定 STA 线程上用；忙时由 Busy.retry 等一下再来。
            // 重试放在 Sta.run 外面，别把 STA 线程睡死（它还要泵 COM 消息）。
            let retries = ref 0

            let onSta (work: unit -> 'T) : 'T =
                let mutable first = true

                Busy.retry 30000 (fun () ->
                    if not first then
                        incr retries

                    first <- false
                    Sta.run work)

            // 你手上那张图：只认人，一个字不写。
            let dwg = onSta(fun () -> acad.ActiveDocument)

            let user =
                onSta(fun () ->
                    {| Name = dwg.Name
                       Entities = dwg.ModelSpace.Count
                       Documents = acad.Documents.Count |})

            Report.line
                output
                "认人"
                (sprintf "「%s」实体数 %d / 图纸数 %d（只读）" user.Name user.Entities user.Documents)

            // 画一个圆：只往这次自己新建的空图里画。
            // Add() 单独一步，不和读属性连写 —— 忙时重试只读那一步，不会重复建图。
            let temp = onSta(fun () -> acad.Documents.Add())
            let tempName = onSta(fun () -> temp.Name)
            Report.line output "临时图" (sprintf "%s（Documents.Add 新建）" tempName)

            // 画圆本身只做一次；读回是纯只读，可以安全重试。
            let circle =
                onSta(fun () -> acad.ActiveDocument.ModelSpace.AddCircle(box [| 10.0; 20.0; 0.0 |], 5.0))

            let drawn =
                onSta(fun () ->
                    let center =
                        match circle.Center with
                        | :? Array as values ->
                            values
                            |> Seq.cast<obj>
                            |> Seq.map(fun v -> Convert.ToDouble v)
                            |> Seq.toArray
                        | _ -> [||]

                    {| Center = center
                       Radius = circle.Radius
                       Entities = acad.ActiveDocument.ModelSpace.Count |})

            Report.line2 output "读回" "圆心" (sprintf "%A" drawn.Center)
            Report.line2 output "读回" "半径" (sprintf "%.1f" drawn.Radius)
            Report.line2 output "读回" "实体数" (string drawn.Entities)

            // 不保存关掉，只认这次新建的那张。
            // 不在 Close 之后立刻枚举 Documents：集合正在变，枚举本身会抛非忙异常；
            // 「关没关掉」交给下面 after 的图纸数 / 活动图名去证。
            let closeNote =
                onSta(fun () ->
                    let active = acad.ActiveDocument

                    if active.Name = tempName then
                        active.Close(box false, Type.Missing)
                        sprintf "已关闭 %s（不保存）" tempName
                    else
                        sprintf "活动图已是 %s，不用再关" active.Name)

            Report.line output "收尾" closeNote

            let after =
                onSta(fun () ->
                    let d = acad.ActiveDocument
                    {| Name = d.Name
                       Entities = d.ModelSpace.Count
                       Documents = acad.Documents.Count |})

            Report.line2 output "核对" "原图实体数" (sprintf "%d → %d" user.Entities after.Entities)
            Report.line2 output "核对" "图纸数" (sprintf "%d → %d" user.Documents after.Documents)
            Report.line2 output "核对" "活动图" (sprintf "%s → %s" user.Name after.Name)

            Should.equal [| 10.0; 20.0; 0.0 |] drawn.Center
            Should.equal 5.0 drawn.Radius
            Should.equal 1 drawn.Entities
            Should.equal false (after.Name = tempName)
            Should.equal user.Entities after.Entities
            Should.equal user.Documents after.Documents
            Should.equal user.Name after.Name

            let now = Report.procIds "acad"
            Report.line2 output "进程" "acad" (sprintf "%A → %A" (Set.toList before) (Set.toList now))
            Should.equal before now
            Report.line output "忙重试" (sprintf "%d 次（Busy.retry 接住的 RPC_E_CALL_REJECTED）" retries.Value)

            Report.conclusion
                output
                "通过：一个 Task 包住整段动作，await 到才结束；圆画在临时图里并读回核对，用户图纸一个字没动"

        }
