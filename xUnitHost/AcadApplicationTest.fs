/// AutoCADHost.Application 的验证：给 PID → 挂回 AcadApplication + COM 根对象。
/// 两条都只读本机现场，绝不新建 / 关闭 AutoCAD，也绝不往用户手上的图纸里写东西：
/// 负例用测试程序自己的 PID，正向只在**本来就开着** acad 时挂回那一只去读数。
namespace XUnitHost

open System
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
