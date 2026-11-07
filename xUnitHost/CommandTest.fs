/// 一条 Fact 演示「派发两条命令，各带一段信息回来」：
///   ① dispatch start —— 起一只常驻 Excel，把命令带回来的地址 / pid / 启动时间 / 现场打印出来；
///   ② dispatch close —— 拿①带回来的地址把那只关掉，把认人 / 收尾结果打印出来。
///
/// 两条命令本身就是 OsHost 的 ProcHost.startWith / closeWith（任意进程通用）。
/// Excel 特有的 COM 启动、ROT 认人、Quit + 放 RCW 只作为策略注入，留在这个文件里；
/// 库层不认任何 Office 类型，命令行程序也不再需要。
namespace XUnitHost

open System
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open Xunit
open FSharp.xUnit
open OsHost

/// 命令的派发与「回信」：派发出去 → 拿回 Result → 摊成报告能直接打印的行。
/// 同一份 Excel 策略也被 RxCommandTest.fs 复用（那边把这两条命令包成事件流）。
module Dispatch =

    /// 一次派发的回信。Lines 是命令带回来的真实值，不是「成功」两个字。
    type Reply =
        { Succeeded: bool
          Lines: (string * string) list }

    let private progIds = [ "Excel.Application" ]

    /// 启动命令交回来的 Application 要一直捏着：
    /// 放掉 RCW 又不 Quit，Excel 会当没人用而自己退出。
    let mutable private app: Application = null

    let private safely (work: unit -> unit) =
        try
            work ()
        with _ ->
            ()

    /// 放掉 RCW：Quit 之后客户端还捏着引用，进程就是不肯退（项目踩过的坑）
    let private releaseAll (extra: obj option) =
        match extra with
        | Some value -> safely (fun () -> Marshal.FinalReleaseComObject value |> ignore)
        | None -> ()

        if not (isNull app) then
            safely (fun () -> Marshal.FinalReleaseComObject app |> ignore)
            app <- null

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    // ---------------- 命令一：起一只常驻 Excel ----------------

    /// 派发给 OsHost 的通用 start：boot 真正起进程，identify 拍现场签地址。
    /// 只认「这次新起的」差集，绝不复用用户手上开着的那只。
    let start () : Result<Launched, string> =
        ProcHost.startWith
            "EXCEL"
            "excel"
            (fun () ->
                Sta.run (fun () ->
                    let created = Com.create progIds :?> Application
                    app <- created

                    // 一本工作簿都没有的自动化实例会自己退出
                    if created.Workbooks.Count = 0 then
                        created.Workbooks.Add() |> ignore

                    safely (fun () -> created.Visible <- false)
                    safely (fun () -> created.DisplayAlerts <- false)
                    // 关键：告诉 Excel「有人在控制」，命令返回后别自己退
                    safely (fun () -> created.UserControl <- true)))
            (fun pid ->
                Sta.run (fun () ->
                    let id = ExcelHost.Host.identity app

                    if id.Pid > 0 then
                        id
                    else
                        // Hwnd 还没就绪时，用差集认出来的进程号兜底
                        { id with
                            Pid = pid
                            StartedMs = Proc.startedMs pid |> Option.defaultValue id.StartedMs }))

    // ---------------- 命令二：按地址把那只关掉 ----------------

    /// 派发给 OsHost 的通用 close：六段地址 / 归属 w / pid / 启动时间都由库层把关，
    /// graceful 这里只做 Excel 的认人 + 优雅退出；认不出就交给库层按 pid 结束（不碰别人）。
    let close (address: string) : Result<Closed, string> =
        ProcHost.closeWith address (fun a ->
            Sta.run (fun () ->
                match Com.running progIds with
                | Some value ->
                    let candidate = value :?> Application

                    match ExcelHost.Host.pidOf candidate with
                    | Some rotPid when rotPid = a.Pid ->
                        let drift = Address.verify a (ExcelHost.Host.identity candidate)

                        if not drift.IsEmpty then
                            // 指纹漂了就不认，按地址里的 pid 结束我们这一只
                            releaseAll None
                            Ok(false, 0)
                        else
                            let quitOk =
                                try
                                    ExcelHost.Host.QuitWithoutSaving candidate
                                    true
                                with _ ->
                                    false

                            releaseAll(Some(box candidate))
                            Ok(quitOk, rotPid)
                    | _ ->
                        // ROT 只暴露了别的实例：不强认，按 pid 结束我们这一只
                        releaseAll None
                        Ok(false, 0)
                | None ->
                    releaseAll None
                    Ok(false, 0)))

    // ---------------- 回信 ----------------

    let private reply (okText: string) (result: Result<'T, string>) (describe: 'T -> (string * string) list) : Reply =
        match result with
        | Ok value -> { Succeeded = true; Lines = ("成败", okText) :: describe value }
        | Error message -> { Succeeded = false; Lines = [ "成败", "Error"; "原因", message ] }

    /// 启动命令带回来的信息
    let startReply (result: Result<Launched, string>) : Reply =
        reply "Ok Launched" result (fun l ->
            [ "地址", Address.format l.Address
              "PID", sprintf "%d，启动 %s" l.Pid (Report.clock l.StartedMs)
              "现场", sprintf "起之前 EXCEL = %A；这次新起 %A" l.Others (Set.singleton l.Pid) ])

    /// 关闭命令带回来的信息
    let closeReply (result: Result<Closed, string>) : Reply =
        reply "Ok Closed" result (fun c ->
            [ "地址", c.Address
              "PID", sprintf "%d，启动 %s" c.Pid (Report.clock c.StartedMs)
              "认人",
              (if c.RecognizedPid > 0 then
                   sprintf "ROT 交回 pid %d，四项指纹全对" c.RecognizedPid
               else
                   "ROT 没认到，按地址里的 pid 结束")
              "收尾", sprintf "Quit=%b　Gone=%b　Forced=%b" c.Quit c.Gone c.Forced ])

/// 两段派发、两段回信：这就是可重跑的记录。
type CommandTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``派发两条命令：起一只常驻 Excel 打印回信，再按地址关掉打印回信``() =
        Report.title output "派发两条命令：起一只常驻 Excel，回来打印信息；按地址关掉，回来打印信息"
        let before = Proc.ids "EXCEL"
        let mutable ownPid = 0

        try
            // ① 派发「启动」命令 —— 回来的信息照原样打印
            let started = Dispatch.start ()
            let startBack = Dispatch.startReply started
            Report.line output "派发①" "ProcHost.startWith \"EXCEL\" \"excel\" boot identify —— 起新进程、补空白工作簿、签地址"

            for label, text in startBack.Lines do
                Report.line2 output "回信①" label text

            Should.equal true startBack.Succeeded
            let launched = Report.ok started
            ownPid <- launched.Pid
            let address = Address.format launched.Address

            match Address.parse address with
            | Ok a ->
                Should.equal "excel" a.Kind
                Should.equal true a.Owned // w：本命令起的，close 才有权关
                Should.equal 0L a.SavedMs // 新工作簿还没存盘
                Should.equal "-" a.PathHash
            | Error message -> Assert.True(false, message)

            // 命令已经回来了，那只进程还在吗？
            let alive = Proc.ids "EXCEL"
            Assert.True(alive.Contains ownPid, sprintf "启动命令回来后 pid %d 不在了，没做到常驻" ownPid)
            Should.equal (Set.singleton ownPid) (Set.difference alive before)

            match Proc.startedMs ownPid with
            | Some ms -> Should.equal launched.StartedMs ms
            | None -> ()

            Report.line
                output
                "常驻"
                (sprintf "命令已返回，pid %d 仍在；启动 %s" ownPid (Report.clock launched.StartedMs))

            // ② 派发「关闭」命令 —— 用①带回来的地址，回来的信息照原样打印
            let closed = Dispatch.close address
            let closeBack = Dispatch.closeReply closed
            Report.line
                output
                "派发②"
                (sprintf "ProcHost.closeWith \"%s\" graceful —— 核地址 / 指纹 → Quit → 放 RCW" address)

            for label, text in closeBack.Lines do
                Report.line2 output "回信②" label text

            Should.equal true closeBack.Succeeded
            let result = Report.ok closed
            Should.equal true result.Gone
            Should.equal address result.Address

            // 回信说 Gone，进程表可能还有半拍延迟（Quit 是异步的）：按项目惯例复核一次，
            // 真还挂着才按 pid 结束；这一步如实记进报告，不替命令遮掩。
            let stubborn = Proc.waitGone "EXCEL" (Set.singleton ownPid) 5000

            if not stubborn.IsEmpty then
                Proc.kill stubborn
                Proc.waitGone "EXCEL" stubborn 3000 |> ignore
                Report.line output "复核" (sprintf "进程表里还挂着 %A，已按 pid 结束（Quit 的半拍）" stubborn)
            else
                Report.line output "复核" "进程表里也查不到这一只了"

            let after = Proc.ids "EXCEL"
            Assert.True(not (after.Contains ownPid), sprintf "pid %d 没关掉" ownPid)
            Should.equal Set.empty (Set.difference after before)
            Report.line output "现场" (sprintf "关之后 EXCEL = %A（起之前 %A），没有残留" after before)
            Report.conclusion output "通过：两条命令各带回一段信息，起的那只已按地址关掉，现场干净"
        finally
            // 万一断言中途失败，也别把这次新起的留在用户机器上：只认自己起的 pid
            if ownPid > 0 && (Proc.ids "EXCEL").Contains ownPid then
                Proc.kill (Set.singleton ownPid)
                Report.line output "兜底" (sprintf "命令没关干净，按本次新起的 pid %d 结束" ownPid)
