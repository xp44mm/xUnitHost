/// Word / PowerPoint 的挂回闸门：PID 不合格就地拒绝，绝不退化成「起一个给你」。
///
/// 离线跑，不碰 COM、不启任何 Office 进程：
///   xUnitHost.exe -method "*PID 闸门*" -noColor -showLiveOutput
///
/// 2026-11-13 删掉 GeneralAppInfo 后，两家的挂回都改成「只认 pid」（与 Excel / AutoCAD 同形）：
/// 前置闸门同步判、不合格立刻交回已完成且 faulted 的 Task / Error，连 STA 队列都不排。
namespace XUnitHost

open System
open System.Diagnostics
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open OsHost

/// 期待 faulted：把异常原样交回，没抛就是这条 Fact 写错了
module internal GateCheck =

    let fault (work: Task) : exn =
        let mutable caught: exn option = None

        try
            work.GetAwaiter().GetResult()
        with ex ->
            caught <- Some ex

        match caught with
        | Some ex -> ex
        | None -> failwith "居然没抛异常"

type AttachGateTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Word 与 PowerPoint 挂回 PID 不合格就地拒绝``() =
        Report.title output "Word / PowerPoint 挂回　PID 闸门（不碰 COM）"

        let me = Process.GetCurrentProcess().Id
        Report.line output "本进程 PID" (sprintf "%d（既不是 WINWORD 也不是 POWERPNT）" me)

        // Word：Task 版，闸门不合格 → 已完成且 faulted，消息与同步版一致
        for pid in [ 0; me ] do
            let work = WordHost.WordTask.GetApplicationAsync pid

            Report.line
                output
                (sprintf "Word PID %d" pid)
                (sprintf "IsCompleted=%b IsFaulted=%b" work.IsCompleted work.IsFaulted)

            Should.equal true work.IsCompleted
            Should.equal true work.IsFaulted

            let ex = GateCheck.fault work
            Report.line2 output "Word" "异常" (sprintf "%s：%s" (ex.GetType().Name) ex.Message)
            Should.equal true (ex :? InvalidOperationException)

        // PowerPoint：同步 Result
        for pid in [ 0; me ] do
            match PowerPointHost.Host.GetApplication pid with
            | Ok _ -> failwithf "PID %d 居然挂上了" pid
            | Error message -> Report.line2 output "PowerPoint" (sprintf "PID %d" pid) message

        Report.conclusion output "通过：两家的闸门都就地拒绝，不排 STA 队列、绝不新建"
