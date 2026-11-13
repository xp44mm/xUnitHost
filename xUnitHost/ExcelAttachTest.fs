/// ExcelHost.ExcelTask.GetApplicationAsync(pid) 的验证：按进程号挂回那一只 Excel，绝不新建、绝不 Quit。
///
/// 离线两条只查进程表（不碰 COM，随时可跑）：
///   xUnitHost.exe -method "*Excel 挂回*" -noColor -showLiveOutput
///
/// 真机那条要先看本机 EXCEL 的进程号：
///   XUNITHOST_EXCEL_PID=12345 xUnitHost.exe -method "*Excel 挂回真机*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.Diagnostics
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open OsHost

/// 期待 faulted：把异常原样交回，没抛就是这条 Fact 写错了
module internal ExcelCheck =

    let fault (work: Task) : exn =
        let mutable caught: exn option = None

        try
            work.GetAwaiter().GetResult()
        with ex ->
            caught <- Some ex

        match caught with
        | Some ex -> ex
        | None -> failwith "居然没抛异常"

type ExcelAttachTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Excel 挂回 PID 不合法时同步交回已完成 faulted Task``() =
        Report.title output "Excel 挂回　PID 闸门（不碰 COM）"
        Report.line output "进程名" ExcelHost.Host.SelfImage

        for pid in [ 0; -7 ] do
            let work = ExcelHost.ExcelTask.GetApplicationAsync pid

            Report.line
                output
                (sprintf "PID %d" pid)
                (sprintf "IsCompleted=%b IsFaulted=%b" work.IsCompleted work.IsFaulted)

            // 校验不合格必须**同步**收口：说明它没把活儿排进 STA 队列
            Should.equal true work.IsCompleted

            let ex = ExcelCheck.fault work

            Report.line output "异常" (sprintf "%s：%s" (ex.GetType().Name) ex.Message)
            Should.equal true (ex :? InvalidOperationException)
            Should.equal true (ex.Message.Contains "PID 必须大于 0")
            Should.equal true work.IsFaulted

        Report.conclusion output "通过：PID ≤ 0 就地拒绝，不排 STA 队列，绝不新建"

    [<Fact>]
    member this.``Excel 挂回 本进程 PID 拒绝挂回不新建``() =
        Report.title output "Excel 挂回　非 EXCEL 进程号"

        let me = Process.GetCurrentProcess().Id
        let work = ExcelHost.ExcelTask.GetApplicationAsync me

        Report.line output "PID" (sprintf "%d（本进程）" me)
        Should.equal true work.IsCompleted

        let ex = ExcelCheck.fault work

        Report.line output "异常" ex.Message
        Should.equal true (ex :? InvalidOperationException)
        Should.equal true (ex.Message.Contains "不是正在运行的 EXCEL 进程")
        Should.equal true work.IsFaulted
        Report.conclusion output "通过：进程表里不是 EXCEL 就拒绝，绝不替人起一只"

    [<Fact>]
    member this.``Excel 挂回 过时的黑板信息按 PID 拒接不新建``() =
        Report.title output "Excel 挂回　过时的黑板信息：能用的只有 Pid，进程已退就拒接"

        // 插件抄下来的原文，PID 早就退了；挂回只认这一条 PID，其余字段一概不猜
        let text =
            "[ExcelVSTO] 实例 999999\r\nKIND: Excel\r\nPID: 999999\r\nROT: !ExcelVSTO.Instance.999999\r\nEXCEL.EXE: C:\\nowhere\\EXCEL.EXE\r\n"

        let info = BlackboardText.ofText text |> Report.ok
        let before = Report.procIds ExcelHost.Host.SelfImage

        Report.line output "黑板" (info.describe())
        Report.line output "取法" "黑板 → Pid → GetApplicationAsync（ROT 名不参与认人）"

        let work = ExcelHost.ExcelTask.GetApplicationAsync info.Pid
        Should.equal true work.IsCompleted

        let ex = ExcelCheck.fault work

        Report.line output "异常" ex.Message
        Should.equal true (ex :? InvalidOperationException)
        Should.equal true (ex.Message.Contains "不是正在运行的 EXCEL 进程")
        Report.line output "现场" (sprintf "EXCEL 进程 %A" (Set.toList before))
        Should.equal before (Report.procIds ExcelHost.Host.SelfImage)
        Report.conclusion output "通过：过时黑板信息按 PID 拒接，四家进程一个没起、一个没关"

    [<Fact>]
    member this.``Excel 挂回真机 按 PID 取到同一只实例``() =
        Report.title output "Excel 挂回真机　按 PID 取回同一只"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_EXCEL_PID：这条要挂在活着的 Excel 上，先看进程号再设环境变量")

        let pid = int pidText
        Report.line output "目标 PID" (string pid)

        let work = ExcelHost.ExcelTask.GetApplicationAsync pid
        let app = work.GetAwaiter().GetResult()

        // 挂回来的必须就是那一只：Hwnd → PID 反查一次
        let back = STA.run(fun () -> ExcelHost.Host.pidOf app)
        Report.line output "反查 PID" (sprintf "%A" back)
        Should.equal (Some pid) back

        let info =
            STA.run(fun () ->
                let version = try app.Version with _ -> "?"

                let book =
                    try
                        app.ActiveWorkbook.Name
                    with _ ->
                        "（没有打开的工作簿）"

                sprintf "版本 %s｜活动工作簿 %s｜HWND %d" version book app.Hwnd)

        Report.line output "现场" info
        Report.conclusion output "通过：按进程号挂回，反查 PID 一致，未新建、未 Quit"
