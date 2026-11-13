/// OsHost.Sta：进程级单例那点语义，离线验，不起任何 host。
///
///   xUnitHost.exe -method "*STA*" -noColor -showLiveOutput
///
/// 关停那条是一次性的（关了就毒掉同进程后面所有 Fact），所以默认跳过，
/// 要验就单开一个进程、只跑它：
///
///   XUNITHOST_STA_SHUTDOWN=1 xUnitHost.exe -method "*STA 关停*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.Threading
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open OsHost

type StaThreadTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``STA 单例：run / runAsync 都排在同一条线程上，here 不换 STA 线程``() : Task =
        task {
            Report.title output "STA　一条固定线程 + 一个队列"

            let me = Thread.CurrentThread.ManagedThreadId
            let first = STA.run (fun () -> Thread.CurrentThread.ManagedThreadId)
            let second = STA.run (fun () -> Thread.CurrentThread.ManagedThreadId)

            Report.line2 output "run" "两次" (sprintf "线程 %d / %d（调用方是 %d）" first second me)
            Should.equal true (first <> me) // 长活不许落在调用线程上
            Should.equal first second // 两次必须是同一条固定线程

            // 排队：后入队的活儿不会超到前面去，run 等它前面全部做完
            let order = ResizeArray<string>()
            let a = STA.runAsync (fun () -> order.Add "一")
            let b = STA.runAsync (fun () -> order.Add "二")
            STA.run (fun () -> order.Add "三") |> ignore
            Report.await a
            Report.await b

            let seam = String.Join("", order)
            Report.line output "顺序" seam
            Should.equal "一二三" seam

            // runAsync 的活儿也在那条线程上，等回来的结果就是它的线程号
            let! afterAwait = STA.runAsync (fun () -> Thread.CurrentThread.ManagedThreadId)
            Report.line2 output "runAsync" "await 回来" (sprintf "线程 %d" afterAwait)
            Should.equal first afterAwait

            // here：人在固定 STA 线程上就就地跑，不再换线程
            let inlineId = STA.run (fun () -> STA.here (fun () -> Thread.CurrentThread.ManagedThreadId))
            Report.line2 output "here" "在固定 STA 线程里" (sprintf "线程 %d" inlineId)
            Should.equal first inlineId

            // here：调用方不是 STA 就临时起一条 STA 跑完再回来
            let apartment = STA.here (fun () -> Thread.CurrentThread.GetApartmentState())
            Report.line2 output "here" "从非 STA 调用方" (sprintf "%A" apartment)
            Should.equal ApartmentState.STA apartment

            Report.conclusion output "run / runAsync 同一条线程、先进先出；here 只在必要时换线程"
        }

    [<Fact>]
    member this.``STA 自等：run 的活儿里再 run 当场报错，不挂死``() : Task =
        task {
            Report.title output "STA　同一条线程上再 run"

            let caught =
                STA.run (fun () ->
                    try
                        // 这一步若被放过，就是排队等自己 —— 永久卡住，测试进程也一起卡
                        STA.run (fun () -> 42) |> ignore
                        None
                    with ex ->
                        Some ex)

            match caught with
            | None ->
                Report.line output "检查" "没有报错（该报的没报，这条调用会挂死）"
                failwith "在固定 STA 线程上再 run 没有报错"
            | Some ex ->
                Report.line2 output "报错" (ex.GetType().Name) ex.Message
                Should.equal typeof<InvalidOperationException> (ex.GetType())
                Should.equal true (ex.Message.Contains "STA")

                // 报错之后线程还好好的：后面的活儿照排
                let after = STA.run (fun () -> 7)
                Report.line2 output "报错之后" "再排一个活儿" (sprintf "拿到 %d" after)
                Should.equal 7 after

                Report.conclusion output "自己等自己转成明确异常，固定线程不受影响"
        }

    [<Fact>]
    member this.``STA 关停：shutdown 之后 run / runAsync 一律拒绝，here 仍可用``() : Task =
        task {
            if String.IsNullOrEmpty(Environment.GetEnvironmentVariable "XUNITHOST_STA_SHUTDOWN") then
                Assert.Skip("关停是一次性的、会毒掉同进程后面所有 Fact：要验就单开一个进程，XUNITHOST_STA_SHUTDOWN=1 只跑这一条")
            else
                Report.title output "STA　shutdown"

                let before = STA.run (fun () -> Thread.CurrentThread.ManagedThreadId)
                Report.line output "关停前" (sprintf "run 跑在线程 %d" before)

                STA.shutdown ()
                Report.line output "关停" "已通知固定线程退出（排上的活先做完）"

                let tryCall (what: string) (call: unit -> unit) =
                    let text =
                        try
                            call ()
                            "没有报错"
                        with ex ->
                            sprintf "%s" (ex.GetType().Name)

                    Report.line2 output "关停后" what text
                    text

                let fromRun = tryCall "run" (fun () -> STA.run (fun () -> 1) |> ignore)
                Should.equal "ObjectDisposedException" fromRun

                let fromAsync = tryCall "runAsync" (fun () -> STA.runAsync (fun () -> 1) |> ignore)
                Should.equal "ObjectDisposedException" fromAsync

                // 再关一次不吵：关停是幂等的
                STA.shutdown ()
                Report.line2 output "关停" "再来一次" "安静返回"

                // 临时 STA 线程跟固定那条无关，here 照旧能用
                let temp = STA.here (fun () -> Thread.CurrentThread.GetApartmentState())
                Report.line2 output "here" "关停后" (sprintf "%A" temp)
                Should.equal ApartmentState.STA temp

                Report.conclusion output "关停只关那条固定线程，之后排活一律 ObjectDisposedException，且不可重启"
        }
