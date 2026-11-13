/// AutoCAD COM 事件 → EventPairSource 的接线层：command / modal / openDrawing / lisp 四对事件。
///
/// Acad 只干「挂上 / 摘下」两件事，不决定跑在哪条线程：
///   · 连接点必须与 RCW 同公寓 → 订阅请在 OsHost.Sta.run 那条固定 STA 线程上做；
///   · 退订同样要回那条线程 → 传 options.Unsubscribe（例如 `Sta.run(fun () -> s.Dispose())`）。
///
/// 成品入口（Task 版）都在 AutoCADHost/Task.fs；这里只管订阅，不造 Task。
///
/// LISP 不是「没有事件」而是**应用级 / 文档级**的 BeginLisp / EndLisp（含 LispCancelled）：
/// 整段表达式算一个整体，内部的 (command "_.LINE" ...) 各自另有 BeginCommand/EndCommand。
/// 所以「这段 LISP 跑完了」要等 EndLisp，不是等某条命令的 EndCommand。
///
/// 为什么能这样接：参数收的是接口 IAcadApplication，跑起来却是 coclass 的 RCW，
/// 而 AcadApplication 本身实现了 _DAcadApplicationEvents_Event —— 所以 `:?>` 做一次 QI 就拿到了，
/// 不必猜类型；add_/remove_ 都挂在这张事件接口上，不在 IAcadApplication 上。
/// 委托类型名由 tlbimp 按事件参数生成（_DAcadApplicationEvents_<事件名>EventHandler）。
namespace AutoCADHost

open System
open Autodesk.AutoCAD.Interop

/// 一段 LISP 怎么收尾的：求值跑完，或被用户 Esc 取消。
/// 两者互斥 —— EndLisp 来了就不会再来 LispCancelled，反之亦然。
type LispOutcome =
    /// EndLisp：整段表达式求值结束（里面 (command ...) 调起的命令也都结束了）
    | Completed
    /// LispCancelled：求值被取消（典型是用户按 Esc）
    | Cancelled

[<RequireQualifiedAccess>]
module AcadBeginEnd =

    /// IAcadApplication → 事件接口（对 RCW 做一次 QI：真实对象是 AcadApplication，实现了这张接口）
    let private eventsOf (app: IAcadApplication) : _DAcadApplicationEvents_Event =
        app :?> _DAcadApplicationEvents_Event

    /// BeginModal / EndModal 一对（无参）。模态对话框关掉后 AutoCAD 才真正空下来。
    let modal (app: IAcadApplication) : EventPairSource<unit, unit> =
        let events = eventsOf app

        {
            SubscribeBegin =
                fun behavior ->
                    let handler =
                        _DAcadApplicationEvents_BeginModalEventHandler(fun () ->
                            behavior()
                        )
                    events.add_BeginModal handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_BeginModal handler
                    }
            SubscribeEnd =
                fun behavior ->
                    let handler =
                        _DAcadApplicationEvents_EndModalEventHandler(fun () -> behavior())
                    events.add_EndModal handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_EndModal handler
                    }
        }

    /// BeginOpen / EndOpen 一对，两边参数都是图纸全名。
    /// 「open 一个 dwg 后 Task 才算完」就靠它，比撞忙重试可靠。
    let openDrawing (app: IAcadApplication) : EventPairSource<string, string> =
        let events = eventsOf app

        {
            SubscribeBegin =
                fun behavior ->
                    // BeginOpen 的参数在类型库里是 ref string（tlbimp 生成 byref），
                    // 所以必须显式写成带类型标注的 lambda，不能直接把函数值丢给委托构造函数。
                    let handler =
                        _DAcadApplicationEvents_BeginOpenEventHandler
                            (fun (fileName: byref<string>) -> behavior fileName)

                    events.add_BeginOpen handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_BeginOpen handler
                    }
            SubscribeEnd =
                fun behavior ->
                    let handler =
                        _DAcadApplicationEvents_EndOpenEventHandler(behavior)
                    events.add_EndOpen handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_EndOpen handler
                    }
        }

    /// BeginCommand / EndCommand 一对，两边参数都是命令名。
    /// 这是「发一条命令后等它真的跑完」唯一可靠的收口信号。
    let command (app: IAcadApplication) : EventPairSource<string, string> =
        let events = eventsOf app

        {
            SubscribeBegin =
                fun behavior ->
                    // handler 由返回的订阅对象握着：连接点只记弱引用，丢了它订阅就悄悄断了
                    let handler =
                        _DAcadApplicationEvents_BeginCommandEventHandler(behavior)
                    events.add_BeginCommand handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_BeginCommand handler
                    }
            SubscribeEnd =
                fun behavior ->
                    let handler =
                        _DAcadApplicationEvents_EndCommandEventHandler(behavior)
                    events.add_EndCommand handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_EndCommand handler
                    }
        }

    /// BeginLisp / EndLisp 一对：整段 LISP 求值算一个整体。
    /// BeginLisp 的 FirstLine 是 AutoCAD 收到的第一行（就是 SendCommand 送进去的那一整行）；
    /// 收尾有两条互斥的路 —— EndLisp（求值结束）与 LispCancelled（Esc），都归到 LispOutcome。
    /// 注意顺序：内部 (command "_.LINE" ...) 的 Begin/EndCommand 先走完，最后才是 EndLisp。
    let lisp (app: IAcadApplication) : EventPairSource<string, LispOutcome> =
        let events = eventsOf app

        {
            SubscribeBegin =
                fun behavior ->
                    let handler =
                        _DAcadApplicationEvents_BeginLispEventHandler(fun (firstLine: string) ->
                            behavior firstLine)

                    events.add_BeginLisp handler
                    { new IDisposable with
                        member _.Dispose() = events.remove_BeginLisp handler
                    }
            SubscribeEnd =
                fun behavior ->
                    let ended =
                        _DAcadApplicationEvents_EndLispEventHandler(fun () -> behavior Completed)

                    let cancelled =
                        _DAcadApplicationEvents_LispCancelledEventHandler(fun () -> behavior Cancelled)

                    events.add_EndLisp ended
                    events.add_LispCancelled cancelled
                    { new IDisposable with
                        member _.Dispose() =
                            events.remove_EndLisp ended
                            events.remove_LispCancelled cancelled
                    }
        }


    /// 一对事件 + 一个启动动作 → Task：start 里做什么，就等哪一类 End。
    let runWith
        (options: EventPair.Options)
        (source: EventPairSource<'Begin, 'End>)
        (start: unit -> unit)
        : System.Threading.Tasks.Task<'End> =
        EventPair.runTaskWith options source start

