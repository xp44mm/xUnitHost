/// 离线自检：不连 AutoCAD，用假事件源把「Begin/End 事件对 → Task」的每条路都走一遍。
///
/// 全部不碰 COM、不起 acad —— 随时可跑：
///   xUnitHost.exe -method "*事件对*" -noColor -showLiveOutput
///
/// 真机那条（挂回活着的 acad 发命令）在 AcadAttachTest.fs。
namespace XUnitHost

open System
open System.Runtime.CompilerServices
open System.Threading
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open AutoCADHost
open XUnitHost.FakeEventPair

/// 断言前的小工具：同步取结果、逼出异常、催 GC、以及两个「谁握着委托」的取样函数。
module internal PairCheck =

    let awaitResult (work: Task<'T>) : 'T = work.GetAwaiter().GetResult()

    /// 期待 faulted：把异常原样交回，没抛就是这条 Fact 写错了
    let fault (work: Task) : exn =
        let mutable caught: exn option = None

        try
            work.GetAwaiter().GetResult()
        with ex ->
            caught <- Some ex

        match caught with
        | Some ex -> ex
        | None -> failwith "居然没抛异常"

    /// 三轮回收，让「只被弱引用记着」的委托确实死掉
    let collect () =
        for _ in 1..3 do
            GC.Collect()
            GC.WaitForPendingFinalizers()

    /// 订阅后把订阅对象留住
    /// 故意捕获一个局部变量：F# 会把「不捕获任何东西的函数」缓存成静态单例，
    /// 那样测的就不是「谁握着委托」，而是「单例永不被回收」了。
    [<MethodImpl(MethodImplOptions.NoInlining)>]
    let subscribeKeeping (fake: Pair) : IDisposable * WeakReference =
        let mutable hits = 0
        let behavior = fun (_: string) -> hits <- hits + 1
        let weak = WeakReference(box behavior)
        let subscription = fake.Source.SubscribeEnd behavior
        subscription, weak

    /// 订阅后把订阅对象丢掉（相当于「没人握着委托」）
    [<MethodImpl(MethodImplOptions.NoInlining)>]
    let subscribeDropping (fake: Pair) : WeakReference =
        let mutable hits = 0
        let behavior = fun (_: string) -> hits <- hits + 1
        let weak = WeakReference(box behavior)
        fake.Source.SubscribeEnd behavior |> ignore
        weak

/// 事件对 → Task 的状态机自检。
type EventPairTest(output: ITestOutputHelper) =

    // ───────────────────────── 正路 ─────────────────────────

    [<Fact>]
    member this.``事件对 Begin 之后 End 一到 Task 完成订阅随即摘掉``() =
        Report.title output "事件对　Begin → End 一到就完成"

        let fake = Pair()
        let work = EventPair.runTask fake.Source (fun () -> fake.FireBegin "LINE")

        Report.line output "只发 Begin" (sprintf "IsCompleted=%b" work.IsCompleted)
        Should.equal false work.IsCompleted
        Should.equal false (work.Wait 0)

        fake.FireEnd "LINE"

        Should.equal "LINE" (PairCheck.awaitResult work)
        Report.line output "结果" (PairCheck.awaitResult work)
        Should.equal 0 fake.AliveEnd
        Should.equal 0 fake.AliveBegin
        Report.line output "残留订阅" (sprintf "Begin %d / End %d" fake.AliveBegin fake.AliveEnd)
        Report.conclusion output "通过：完成前不收口，完成后两条订阅都摘干净"

    [<Fact>]
    member this.``事件对 End 从别的线程送来 Task 照样完成``() =
        Report.title output "事件对　End 从线程池送来"

        let fake = Pair()
        let work = EventPair.runTask fake.Source (fun () -> fake.FireBegin "ZOOM")

        ThreadPool.QueueUserWorkItem(fun _ -> fake.FireEnd "ZOOM") |> ignore

        Should.equal true (work.Wait 2000)
        Report.line output "结果" work.Result
        Should.equal "ZOOM" work.Result
        Report.conclusion output "通过：完成由事件线程决定，与调用线程无关"

    [<Fact>]
    member this.``事件对 End 来两条只认第一条``() =
        Report.title output "事件对　并发命令的 End 只认第一条"

        let fake = Pair()
        let work = EventPair.runTask fake.Source (fun () -> fake.FireBegin "LINE")

        fake.FireEnd "第一条"
        fake.FireEnd "第二条"

        let actual = PairCheck.awaitResult work
        Report.line output "结果" actual
        Should.equal "第一条" actual
        Report.conclusion output "通过：谁先到谁算数，第二条不覆盖"

    [<Fact>]
    member this.``事件对 串行两条命令互不串台``() =
        Report.title output "事件对　串行两条"

        let fake = Pair()

        let first = EventPair.runTask fake.Source (fun () -> fake.FireBegin "LINE")
        fake.FireEnd "LINE"

        let firstResult = PairCheck.awaitResult first
        Report.line output "第一条" firstResult
        Should.equal "LINE" firstResult

        let second = EventPair.runTask fake.Source (fun () -> fake.FireBegin "CIRCLE")
        fake.FireEnd "CIRCLE"

        let secondResult = PairCheck.awaitResult second
        Report.line output "第二条" secondResult
        Should.equal "CIRCLE" secondResult
        Report.conclusion output "通过：上一条 Task 收口后才发下一条"

    [<Fact>]
    member this.``事件对 task 版 Wait 直接拿结果``() =
        Report.title output "事件对　runTask + Wait"

        let fake = Pair()
        let work = EventPair.runTask fake.Source (fun () -> fake.FireBegin "MOVE")

        ThreadPool.QueueUserWorkItem(fun _ -> fake.FireEnd "MOVE") |> ignore

        let finished = work.Wait 2000
        Report.line output "Wait(2000)" (sprintf "IsCompleted=%b" finished)
        Should.equal true finished
        Report.line output "结果" work.Result
        Should.equal "MOVE" work.Result
        Report.conclusion output "通过：task 在 2s 内完成，结果就是 Begin 的载荷"

    // ───────────────────────── 防串台：RequireBegin ─────────────────────────

    [<Fact>]
    member this.``事件对 没见到 Begin 的 End 不算数``() =
        Report.title output "事件对　RequireBegin（默认 true）"

        let fake = Pair()

        // 启动时就先来一条 End：模拟上一条命令的 End 迟到
        let work = EventPair.runTask fake.Source (fun () -> fake.FireEnd "上一条命令")

        Report.line output "串台 End 之后" (sprintf "IsCompleted=%b" work.IsCompleted)
        Should.equal false (work.Wait 300)

        fake.FireBegin "LINE"
        fake.FireEnd "LINE"

        let actual = PairCheck.awaitResult work
        Report.line output "补上 Begin 后" actual
        Should.equal "LINE" actual
        Report.conclusion output "通过：没有 Begin 的 End 被丢弃，继续等真的那条"

    [<Fact>]
    member this.``事件对 RequireBegin 关掉时 End 单独完成``() =
        Report.title output "事件对　RequireBegin=false"

        let fake = Pair()

        let options =
            { EventPair.defaults with
                RequireBegin = false }

        let work = EventPair.runTaskWith options fake.Source (fun () -> fake.FireEnd "ARC")

        let actual = PairCheck.awaitResult work
        Report.line output "结果" actual
        Should.equal "ARC" actual
        Report.conclusion output "通过：这条开关真的能放开 Begin 的前置条件"

    // ───────────────────────── 失败三条路 ─────────────────────────

    [<Fact>]
    member this.``事件对 超时抛 TimeoutException 且已退订``() =
        Report.title output "事件对　超时"

        let fake = Pair()

        let options =
            { EventPair.defaults with
                Timeout = Some(TimeSpan.FromMilliseconds 250.0) }

        let work = EventPair.runTaskWith options fake.Source (fun () -> ())
        let ex = PairCheck.fault work

        Report.line output "异常" (ex.GetType().Name)
        Report.line output "文案" ex.Message
        Should.equal true (ex :? TimeoutException)
        Should.equal 0 fake.AliveEnd
        Should.equal 0 fake.AliveBegin
        Report.line output "残留订阅" (sprintf "Begin %d / End %d" fake.AliveBegin fake.AliveEnd)
        Report.conclusion output "通过：超时是 faulted，不是永远挂着，且退订了"

    [<Fact>]
    member this.``事件对 取消是 Canceled 不是 Faulted 且已退订``() =
        Report.title output "事件对　调用方取消"

        let fake = Pair()
        use cts = new CancellationTokenSource()

        let options =
            { EventPair.defaults with
                Cancellation = cts.Token }

        let work = EventPair.runTaskWith options fake.Source (fun () -> ())

        cts.Cancel()

        try
            work.GetAwaiter().GetResult() |> ignore
            failwith "取消之后居然没抛异常"
        with :? OperationCanceledException ->
            ()

        Report.line output "状态" (sprintf "IsCanceled=%b IsFaulted=%b" work.IsCanceled work.IsFaulted)
        Should.equal true work.IsCanceled
        Should.equal 0 fake.AliveEnd
        Report.conclusion output "通过：取消走 TAP 约定，Task 是 Canceled"

    [<Fact>]
    member this.``事件对 start 抛异常是 faulted 且已退订``() =
        Report.title output "事件对　启动动作自己炸了"

        let fake = Pair()
        let work = EventPair.runTask fake.Source (fun () -> failwith "动作自己炸了")
        let ex = PairCheck.fault work

        Report.line output "异常" ex.Message
        Should.equal true (ex.Message.Contains "动作自己炸了")
        Should.equal 0 fake.AliveEnd
        Should.equal 0 fake.AliveBegin
        Report.conclusion output "通过：动作异常原样交回，订阅不留在实例上"

    // ─────────────────── 为什么要一直握着订阅对象（弱引用） ───────────────────

    [<Fact>]
    member this.``事件对 订阅对象握着委托 GC 之后订阅照样有效``() =
        Report.title output "事件对　订阅对象 = 委托的强引用"

        let fake = Pair()
        let subscription, weak = PairCheck.subscribeKeeping fake

        PairCheck.collect()

        Report.line output "GC 之后" (sprintf "委托存活=%b" weak.IsAlive)
        Should.equal true weak.IsAlive

        fake.FireEnd "还活着"

        subscription.Dispose()

        Should.equal 0 fake.AliveEnd
        Report.conclusion output "通过：握着订阅对象，委托就不会被 GC 收走"

    [<Fact>]
    member this.``事件对 订阅对象丢了就失效``() =
        Report.title output "事件对　没人握着委托 = 订阅悄悄失效"

        let fake = Pair()
        let weak = PairCheck.subscribeDropping fake

        PairCheck.collect()

        Report.line output "GC 之后" (sprintf "委托存活=%b" weak.IsAlive)
        Should.equal false weak.IsAlive

        fake.FireEnd "没人接"

        Should.equal 0 fake.AliveEnd
        Report.conclusion output "通过：这就是 COM 连接点只记弱引用的坑"
