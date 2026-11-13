/// 假的事件源：不连 AutoCAD 也能把「Begin/End → Task」这套状态机跑到底。
///
/// 故意做成**连接点只记弱引用**（跟 COM 一样），这样才能验出真规矩：
/// 委托不被谁强引用就会被 GC 收走 —— 这正是 AutoCADHost/AcadBeginEnd.fs
/// 里「委托必须由本侧一直握着」那条注释的机理解释，而不是一句口头约定。
namespace XUnitHost

open System
open AutoCADHost

module internal FakeEventPair =

    /// 一个连接点：Add 返回订阅对象，Fire 逐个喂给还活着的处理器。
    type ConnectionPoint() =
        let handlers = ResizeArray<WeakReference>()
        let gate = obj()

        let alive () =
            lock
                gate
                (fun () ->
                    handlers
                    |> Seq.filter(fun weak -> not(isNull weak.Target))
                    |> Seq.length
                )

        /// 挂一个处理器。返回的订阅对象握着它 —— 订阅对象一丢，处理器随时可能被回收。
        member _.Add(behavior: string -> unit) : IDisposable =
            lock gate (fun () -> handlers.Add(WeakReference(box behavior)))
            { new IDisposable with
                member _.Dispose() =
                    // 连接点只能摘「同一个委托实例」：所以退订必须交回当初那份，不能新建一个
                    lock
                        gate
                        (fun () ->
                            let index =
                                handlers.FindIndex(fun weak ->
                                    match weak.Target with
                                    | null -> false
                                    | target -> obj.ReferenceEquals(target, box behavior)
                                )

                            if index >= 0 then
                                handlers.RemoveAt index
                        )
            }

        /// 送一条通知：只剩死引用的就跳过（等价于 COM 里「订阅已悄悄失效」）
        member _.Fire(value: string) =
            let snapshot =
                lock
                    gate
                    (fun () ->
                        let current = handlers.ToArray()
                        // 顺手清掉已死的弱引用，别让它们冒充「还挂着」
                        handlers.RemoveAll(fun weak -> isNull weak.Target)
                        |> ignore
                        current
                    )

            for weak in snapshot do
                match weak.Target with
                | null -> ()
                | target -> (unbox<string -> unit> target) value

        /// 还活着的处理器数量
        member _.Alive = alive()

        member _.Total = lock gate (fun () -> handlers.Count)

    /// 一对假的 Begin/End 事件源
    type Pair() =
        let beginPoint = ConnectionPoint()
        let endPoint = ConnectionPoint()

        member _.Source: EventPairSource<string, string> =
            {
                SubscribeBegin = fun behavior -> beginPoint.Add behavior
                SubscribeEnd = fun behavior -> endPoint.Add behavior
            }

        member _.FireBegin(value: string) = beginPoint.Fire value
        member _.FireEnd(value: string) = endPoint.Fire value
        member _.AliveBegin = beginPoint.Alive
        member _.AliveEnd = endPoint.Alive
