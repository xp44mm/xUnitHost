/// Rx 版的两段派发：一条事件流串起 start / close，回来的信息逐条打印。
///
/// 与同步版 CommandTest.fs 共用同一份 Excel 策略（Dispatch.start / Dispatch.close），
/// 差别只在编排：这里不自己一步步调命令，而是订阅
/// `Observable.Create(start) → SelectMany(Started ⇒ close) → Timeout(30s)`，
/// 让进程退出的事件把整条流收尾。报告就是 output 里那份。
namespace XUnitHost

open System
open System.Collections.Generic
open System.Threading
open Xunit
open FSharp.xUnit
open OsHost

type RxCommandTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Rx 两条命令：start 的事件流接上 close，回来逐条打印信息``() =
        Report.title output "Rx 派发两条命令：start → 打印回信 → close → 打印退出事件"
        let before = Proc.ids "EXCEL"
        let events = ResizeArray<ProcessEvent>()
        let failure = ref (null: exn)
        let ownPid = ref 0
        use doneSignal = new ManualResetEventSlim(false)

        // 两次派发接成一条流：start 是 Observable.Create，Started 一到就按地址派发 close
        let flow =
            ProcRx.flow
                (TimeSpan.FromSeconds 30.)
                (TimeSpan.FromMilliseconds 300.)
                (ProcRx.start Dispatch.start)
                (fun a -> ProcRx.close a Dispatch.close 1000)

        // 事件可能来自 Exited 事件线程 / Interval 线程池：先收好，回到主线程再打印
        let observer =
            { new IObserver<ProcessEvent> with
                member _.OnNext ev =
                    lock events (fun () -> events.Add ev)
                    match ev with
                    | Started l -> ownPid.Value <- l.Pid
                    | _ -> ()
                member _.OnError ex =
                    failure.Value <- ex
                    doneSignal.Set()
                member _.OnCompleted() = doneSignal.Set() }

        use subscription = flow.Subscribe(observer)

        try
            let finished = doneSignal.Wait(TimeSpan.FromSeconds 35.)
            let snapshot = lock events (fun () -> events.ToArray())

            Report.line
                output
                "管道"
                "Observable.Create(start) → SelectMany(Started ⇒ Timer(300ms) ⇒ close) → Timeout(30s) → Subscribe"

            // 回来打印信息：按事件到达顺序，一条事件一行，贴的都是命令带回来的值
            for ev in snapshot do
                match ev with
                | Started l ->
                    Report.line2
                        output
                        "回信①"
                        "Started"
                        (sprintf
                            "%s　pid=%d　启动 %s"
                            (Address.format l.Address)
                            l.Pid
                            (Report.clock l.StartedMs))

                    Report.line2
                        output
                        "回信①"
                        "现场"
                        (sprintf "起之前 EXCEL = %A；这次新起 %A" l.Others (Set.singleton l.Pid))
                | Closed c ->
                    Report.line2
                        output
                        "回信②"
                        "Closed"
                        (sprintf "%s　Quit=%b Gone=%b Forced=%b" c.Address c.Quit c.Gone c.Forced)

                    Report.line2
                        output
                        "回信②"
                        "认人"
                        (if c.RecognizedPid > 0 then
                             sprintf "ROT 交回 pid %d，四项指纹全对" c.RecognizedPid
                         else
                             "ROT 没认到，按地址里的 pid 结束")
                | Exited(pid, code, via) ->
                    Report.line2
                        output
                        "回信③"
                        "Exited"
                        (sprintf
                            "pid=%d　ExitCode=%s　来源 %s　流 OnCompleted"
                            pid
                            (match code with
                             | Some c -> string c
                             | None -> "取不到（外部进程）")
                            via)
                | Refused reason -> Report.line2 output "回信" "Refused" reason

            if not (isNull failure.Value) then
                Report.line
                    output
                    "错误"
                    (sprintf "%s：%s" (failure.Value.GetType().Name) failure.Value.Message)

            Assert.True(finished, "事件流 35 秒内没走完（Timeout 没兜住）")
            Assert.True(isNull failure.Value, sprintf "事件流报错：%A" failure.Value)

            let launched = snapshot |> Array.tryPick (function Started l -> Some l | _ -> None)
            let closed = snapshot |> Array.tryPick (function Closed c -> Some c | _ -> None)
            let exited = snapshot |> Array.tryPick (function Exited(pid, code, via) -> Some(pid, code, via) | _ -> None)

            Assert.True(launched.IsSome, "没收到 Started：启动命令没派发成功")
            Assert.True(closed.IsSome, "没收到 Closed：关闭命令没派发成功")
            Assert.True(exited.IsSome, "没收到 Exited：进程退出事件丢了")

            Should.equal (Address.format launched.Value.Address) closed.Value.Address
            Should.equal launched.Value.Pid closed.Value.Pid
            Should.equal true closed.Value.Gone
            let exitedPid, _, _ = exited.Value
            Should.equal exitedPid ownPid.Value

            // 流走完后再复核现场（Quit 异步，进程表可能有半拍延迟）
            let stubborn = Proc.waitGone "EXCEL" (Set.singleton ownPid.Value) 5000

            if not stubborn.IsEmpty then
                Proc.kill stubborn
                Proc.waitGone "EXCEL" stubborn 3000 |> ignore
                Report.line output "复核" (sprintf "进程表里还挂着 %A，已按 pid 结束" stubborn)
            else
                Report.line output "复核" "进程表里也查不到这一只了"

            let after = Proc.ids "EXCEL"
            Should.equal Set.empty (Set.difference after before)
            Report.line output "现场" (sprintf "关之后 EXCEL = %A（起之前 %A），没有残留" after before)
            Report.conclusion output "通过：Rx 一条流串起 start / close，两次回信和退出事件都打印出来，现场干净"
        finally
            // 流中途出错也别把这次新起的留在用户机器上：只认自己起的 pid
            if ownPid.Value > 0 && (Proc.ids "EXCEL").Contains ownPid.Value then
                Proc.kill (Set.singleton ownPid.Value)
                Report.line output "兜底" (sprintf "流没关干净，按本次新起的 pid %d 结束" ownPid.Value)
