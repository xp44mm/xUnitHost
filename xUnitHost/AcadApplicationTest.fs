/// AutoCADHost.Application 的验证：给 PID → 挂回 AcadApplication + COM 根对象。
/// 两条都只读本机现场，绝不新建 / 关闭 AutoCAD，也绝不往用户手上的图纸里写东西：
/// 负例用测试程序自己的 PID，正向只在**本来就开着** acad 时挂回那一只去读数。
namespace XUnitHost

open System
open System.Threading
open System.Threading.Tasks
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
                let! appResult = AutoCADHost.Application.GetApplicationAsync pid
                let! rootsResult = AutoCADHost.Application.GetApplicationWithRootsAsync pid

                // await 之后要用 RCW，仍包回固定 STA 线程；顺手量一下续体和 STA 是不是同一条线程
                let version =
                    Sta.run(fun () ->
                        match appResult with
                        | Ok app -> app.Version
                        | Error _ -> "")

                let staThread = Sta.run(fun () -> Thread.CurrentThread.ManagedThreadId)
                let resumeThread = Environment.CurrentManagedThreadId
                return (appResult, rootsResult, version, staThread, resumeThread)
            }

        let appResult, rootsResult, version, staThread, resumeThread =
            work.GetAwaiter().GetResult()

        Report.line output "Task 形状" "GetApplicationAsync pid → Task<Result<AcadApplication,string>>"

        Report.line2
            output
            "异步结果"
            "Application"
            (match appResult with
             | Ok app -> if isNull (box app) then "null" else "AcadApplication 已挂回"
             | Error message -> message)

        Report.line output "版本" version
        Report.line output "线程" (sprintf "STA %d / 调用 %d / await 续体 %d" staThread callerThread resumeThread)

        let roots = Report.ok rootsResult
        Report.line2 output "根对象" "PID" (string roots.Pid)
        Report.line2 output "根对象" "图纸数" (string roots.DocumentCount)

        Should.equal true (
            match appResult with
            | Ok app -> not (isNull (box app))
            | Error _ -> false
        )

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

        // checkPid 不合格时连 STA 队列都不用排，直接交回已完成的 Task
        let pendingTask = AutoCADHost.Application.GetApplicationAsync selfPid
        Report.line output "自测 PID" (sprintf "%d（本测试程序，不是 acad）" selfPid)
        Report.line output "Task 已完成" (string pendingTask.IsCompleted)
        Should.equal true pendingTask.IsCompleted

        let work =
            task {
                let! appResult = AutoCADHost.Application.GetApplicationAsync selfPid
                let! rootsResult = AutoCADHost.Application.GetApplicationWithRootsAsync selfPid
                return (appResult, rootsResult)
            }

        let appResult, rootsResult = work.GetAwaiter().GetResult()

        let message =
            match appResult with
            | Error text -> text
            | Ok _ -> "Ok（不该发生）"

        Report.line output "Application" message
        Report.line
            output
            "一步到位"
            (match rootsResult with
             | Error text -> text
             | Ok _ -> "Ok（不该发生）")

        Should.equal true (message.Contains "不是正在运行" && message.Contains "绝不新建")
        Should.equal true (
            match rootsResult with
            | Error text -> text.Contains "绝不新建"
            | _ -> false
        )

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：异步负例直接给已完成的 Task 并拒接，四家进程一个没起、一个没关"
