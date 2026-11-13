/// 只走「命令行」这一条通道，两类输入各配一种完成判据：
///   · AutoCAD **内置命令**（_.LINE / _.CIRCLE）→ 发 BeginCommand / EndCommand，等 EndCommand；
///   · **AutoLISP** 表达式 → 发应用级的 BeginLisp / EndLisp，等 **EndLisp**（按 Esc 则是 LispCancelled）。
///
/// 粒度要说清：一整段 (progn ...) 算一个 LISP 求值整体，里面 (command "_.LINE" ...) 的
/// Begin/EndCommand 先各自走完，最后才来 EndLisp —— 「这段 LISP 跑完了」只能等后者，
/// 等 EndCommand 会在第一条内嵌命令结束时提前返回。
///
/// EndLisp 只说明求值结束、不说明成功：AutoLISP 的求值错误只在命令行打印，不回报给调用方。
/// 要拿数据 / 判成败，仍得让代码自己写结果文件（下面 bet-write 那段就是干这个）。
///
/// 不用直接 COM API（AddLine / AddCircle / Delete），不采样 IsQuiescent，也不发任何自定义命令。
///
/// 跑法：先看本机 acad 的进程号，然后
///   XUNITHOST_ACAD_PID=12345 xUnitHost.exe -method "*acad*" -noColor -showLiveOutput
///
/// 只改本次自己画的那一条命令产生的实体：按句柄删圆，绝不碰图纸里别的东西，也**绝不保存**。
namespace XUnitHost

open System
open System.Diagnostics
open System.IO
open System.Text
open System.Threading
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open Autodesk.AutoCAD.Interop
open AutoCADHost
open OsHost

/// 库里的 `AutoCADHost.AutoCADTask` 模块与 `System.Threading.Tasks.Task` 同名，这里起个别名免得每次写全名。
/// 成品入口全在 `Task.fs`：挂回 `GetApplicationAsync`、命令 `sendCommandAsync`、
/// LISP `sendLispAsync`（不等）/ `runLispAsync`（等后缀标记）；`Application.fs` 只留 checkPid / findOnSta / unwrap 这些零件。
/// LISP 和命令一样都收**文档**（IAcadDocument），SendCommand 在文档上。
module AutoCADTask = AutoCADHost.AutoCADTask

/// LISP 侧的小工具：路径、结果文件、脚本正文。都是纯函数，放在类前面。
module internal AcadScript =

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

    /// (load "...") 只认绝对路径；AutoLISP 字符串里别留反斜杠
    let lispPath (path: string) = path.Replace('\\', '/')

    /// 可以从结果文件独占打开、且内容非空 ⇒ LISP 已经 close，结果可信
    let readSettled (path: string) : string option =
        if not (File.Exists path) then
            None
        else
            try
                use stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None)
                use reader = new StreamReader(stream, Encoding.UTF8)
                let text = reader.ReadToEnd().Trim()
                if text = "" then None else Some text
            with _ ->
                None

    /// LISP 不发事件，只能等它把结果写完；超时就报错，不无限挂住
    let awaitSettled (path: string) (timeoutMs: int) : Task<string> =
        Task.Run(fun () ->
            let clock = Stopwatch.StartNew()
            let mutable found: string option = None

            while found.IsNone && clock.ElapsedMilliseconds < int64 timeoutMs do
                found <- readSettled path
                if found.IsNone then Thread.Sleep 100

            match found with
            | Some text -> text
            | None -> failwithf "LISP 结果文件 %dms 内没写完：%s" timeoutMs path)

    /// 只读：最后实体的类型与句柄。entget 的组码 0 是 DXF 名（圆是 "CIRCLE"，不是 AcDbCircle）。
    let lastEntityLines =
        [ "(setq e (entlast))"
          "(if e"
          "  (progn"
          "    (setq d (entget e))"
          "    (bet-write \"{OUT}\" (strcat \"type=\" (cdr (assoc 0 d)) \" handle=\" (cdr (assoc 5 d))))"
          "  )"
          "  (bet-write \"{OUT}\" \"type=none\")"
          ")" ]

    /// 从 "type=CIRCLE handle=1F" 里取出句柄
    let handleOf (report: string) =
        report.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
        |> Array.tryFind(fun token -> token.StartsWith "handle=")
        |> Option.map(fun token -> token.Substring "handle=".Length)

/// 挂回活着的 acad：离线三条闸门 + 真机一条命令通道。
type AcadAttachTest(output: ITestOutputHelper) =

    // ─────────────────── 离线：前置闸门的 Task 语义 ───────────────────

    [<Fact>]
    member this.``挂回 PID 不合法时同步交回已完成 faulted Task``() =
        Report.title output "挂回　PID 闸门（不碰 COM）"

        for pid in [ 0; -7 ] do
            let work = AutoCADTask.GetApplicationAsync pid

            Report.line output (sprintf "PID %d" pid) (sprintf "IsCompleted=%b IsFaulted=%b" work.IsCompleted work.IsFaulted)

            // 校验不合格必须**同步**收口：说明它没把活儿排进 STA 队列
            Should.equal true work.IsCompleted

            let ex = AcadScript.fault work

            Report.line output "异常" (sprintf "%s：%s" (ex.GetType().Name) ex.Message)
            Should.equal true (ex :? InvalidOperationException)
            Should.equal true (ex.Message.Contains "PID 必须大于 0")
            Should.equal true work.IsFaulted

        Report.conclusion output "通过：PID ≤ 0 就地拒绝，不排 STA 队列，绝不新建"

    [<Fact>]
    member this.``挂回 本进程 PID 拒绝挂回不新建``() =
        Report.title output "挂回　非 acad 进程号"

        let me = Process.GetCurrentProcess().Id
        let work = AutoCADTask.GetApplicationAsync me

        Report.line output "PID" (sprintf "%d（本进程）" me)
        Should.equal true work.IsCompleted

        let ex = AcadScript.fault work

        Report.line output "异常" ex.Message
        Should.equal true (ex :? InvalidOperationException)
        Should.equal true (ex.Message.Contains "不是正在运行的 acad 进程")
        Should.equal true work.IsFaulted
        Report.conclusion output "通过：进程表里不是 acad 就拒绝，绝不替人起一只"

    // ─────────────────── 真机：内置命令 + LISP ───────────────────

    [<Fact>]
    member this.``acad 发内置命令等 EndCommand 并用 LISP 删掉本次画的圆``() =
        Report.title output "acad　命令行通道（EndCommand 收口 + LISP 结果文件）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做，先看进程号再设环境变量")

        let pid = int pidText

        let temp = Path.GetTempPath()
        let stamp = Guid.NewGuid().ToString("N").Substring(0, 8)

        let tempFile (name: string) (ext: string) =
            Path.Combine(temp, sprintf "acadattach_%s_%s.%s" stamp name ext)

        let t =
            task {
                let! acad = AutoCADTask.GetApplicationAsync pid

                let identity =
                    STA.run(fun () -> sprintf "PID %d｜HWND 0x%08X｜%s" pid acad.HWND acad.ActiveDocument.Name)

                Report.line output "现场" identity

                let options =
                    { EventPair.defaults with
                        // 退订要回 STA，但不能阻塞等 STA：EndCommand 可能就在 STA 线程上送达 → 自己等自己
                        Unsubscribe = fun s -> STA.runAsync(fun () -> s.Dispose()) |> ignore }

                /// 发一行命令行输入并等 EndCommand。SendCommand 返回 ≠ 命令跑完。
                /// 参数收的是文档（AcadDocument 实现 IAcadDocument）：SendCommand 在文档上。
                let sendAndWait (commandLine: string) : Task<string> =
                    STA.runAsync(fun () ->
                        AutoCADTask.sendCommandAsync options acad.ActiveDocument commandLine)
                        .Unwrap()

                /// 一行 LISP 太长会被命令行截断，所以写成 .lsp 再 (load)；{OUT} 换成结果文件路径。
                /// 内容纯 ASCII：AutoCAD 2020 的 AutoLISP 按本地 ANSI(GBK) 读 .lsp，写中文会踩编码坑。
                let runScript (name: string) (body: string list) : Task<string> =
                    task {
                        let result = tempFile name "txt"
                        let script = tempFile name "lsp"

                        if File.Exists result then
                            File.Delete result

                        let lines =
                            [ ";; Pure ASCII on purpose: AutoCAD 2020 evaluates .lsp in the local ANSI code page."
                              "(defun bet-write (path text / f)"
                              "  (setq f (open path \"w\"))"
                              "  (write-line text f)"
                              "  (close f)"
                              "  (princ)"
                              ")" ]
                            @ (body |> List.map(fun line -> line.Replace("{OUT}", AcadScript.lispPath result)))
                            @ [ "(princ)" ]

                        File.WriteAllText(script, String.Join("\r\n", lines) + "\r\n", Encoding.ASCII)

                        // LISP 不是命令，没有 EndCommand 可等：loadLispAsync 只保证「已交出去」
                        do! STA.runAsync(fun () -> AutoCADTask.loadLispAsync acad.ActiveDocument script).Unwrap()

                        // LISP 侧已经 close 才算数：等结果文件写完
                        return! AcadScript.awaitSettled result 15000
                    }

                // ── 1. 内置命令：等 EndCommand
                let! lineCommand = sendAndWait "_.LINE 0,0 30,0 \n"
                Report.line output "LINE" (sprintf "EndCommand=%s" lineCommand)

                // ── 2. 只读：最后实体
                let! lastAfterLine = runScript "after-line" AcadScript.lastEntityLines
                Report.line output "最后实体" lastAfterLine
                Should.equal true (lastAfterLine.StartsWith "type=")

                // ── 3. 再画一个圆，同样等 EndCommand
                let! circleCommand = sendAndWait "_.CIRCLE 30,20 5\n"
                Report.line output "CIRCLE" (sprintf "EndCommand=%s" circleCommand)

                let! lastAfterCircle = runScript "after-circle" AcadScript.lastEntityLines
                Report.line output "最后实体" lastAfterCircle
                Should.equal true (lastAfterCircle.StartsWith "type=CIRCLE")

                // ── 4. 按句柄删，只删本次画的那个圆：不碰图纸里别的实体，也绝不保存
                match AcadScript.handleOf lastAfterCircle with
                | Some handle ->
                    Report.line output "圆句柄" handle

                    let! erased =
                        runScript
                            "erase-circle"
                            [ sprintf "(setq e (handent \"%s\"))" handle
                              "(if e"
                              "  (progn"
                              "    (setq h (cdr (assoc 5 (entget e))))"
                              "    (entdel e)"
                              "    (bet-write \"{OUT}\" (strcat \"erased handle=\" h))"
                              "  )"
                              "  (bet-write \"{OUT}\" \"erased none\")"
                              ")" ]

                    Report.line output "删除" erased
                    Should.equal true (erased.StartsWith "erased handle=")

                    let! lastAfterErase = runScript "after-erase" AcadScript.lastEntityLines
                    Report.line output "删除后" lastAfterErase
                    Should.equal false (lastAfterErase = lastAfterCircle)
                | None -> failwithf "没拿到圆句柄：%s" lastAfterCircle

                Report.conclusion output "通过：命令按 EndCommand 收口，LISP 按结果文件收口，删除只命中本次的圆"
            }

        t.GetAwaiter().GetResult()

    [<Fact>]
    member this.``acad 发一行 LISP 不等 结果文件也会落盘``() =
        Report.title output "acad　sendLispAsync（只交出去，不等它跑完）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做，先看进程号再设环境变量")

        let pid = int pidText

        let t =
            task {
                let! acad = AutoCADTask.GetApplicationAsync pid

                let stamp = Guid.NewGuid().ToString("N").Substring(0, 8)
                let target = Path.Combine(Path.GetTempPath(), sprintf "acad_sendlisp_%s.txt" stamp)

                try
                    if File.Exists target then
                        File.Delete target
                with _ ->
                    ()

                // 一行搞定：哨兵 hello + 这版 acad 的版本号 + 当前图纸名。
                // · 开头先把上一轮可能没 close 的 bet-f 关掉（LISP 报错时句柄会挂在 AutoCAD 里）；
                // · 全用 vl-princ-to-string：`(getvar "PID")` 在 R23.1 上就是 nil（没有这个变量），
                //   itoa 遇到 nil / 实数直接报错，write-line 就写不出东西。
                let code =
                    sprintf
                        "(progn (if bet-f (close bet-f)) (setq bet-f (open \"%s\" \"w\")) (write-line \"hello\" bet-f) (write-line (vl-princ-to-string (getvar \"ACADVER\")) bet-f) (write-line (vl-princ-to-string (getvar \"DWGNAME\")) bet-f) (close bet-f) (setq bet-f nil))"
                        (AcadScript.lispPath target)

                // 挂回时从 COM 读到的图纸名，等会儿和 LISP 自己报的对照
                let docName = STA.run(fun () -> acad.ActiveDocument.Name)

                let clock = Stopwatch.StartNew()

                let! _ =
                    STA.runAsync(fun () -> AutoCADTask.sendLispAsync acad.ActiveDocument code)
                        .Unwrap()

                // sendLispAsync 的契约：只保证「已交给 AutoCAD」，所以回来得很快、且不带结果
                Report.line output "调用返回" (sprintf "%d ms｜不携带执行结果" clock.ElapsedMilliseconds)

                // 真要确认它执行了，得自己看痕迹：等 acad 把文件写完
                let! heard = AcadScript.awaitSettled target 10000

                let lines =
                    heard.Split([| '\r'; '\n' |], StringSplitOptions.RemoveEmptyEntries)

                Report.line output "结果文件" (String.Join(" / ", lines))
                Report.line output "耗时" (sprintf "%d ms（含等文件）" clock.ElapsedMilliseconds)

                // 第一行是哨兵，2/3 行是 acad 自己的版本与图纸名 —— 证明确实在那一只里执行的
                Should.equal 3 lines.Length
                Should.equal "hello" lines.[0]
                Report.line output "LISP 自报" (sprintf "%s｜%s" lines.[1] lines.[2])
                Should.equal docName lines.[2]

                try
                    File.Delete target
                with _ ->
                    ()

                Report.conclusion output "通过：交出去就返回，痕迹由 AutoCAD 自己写出来"
            }

        t.GetAwaiter().GetResult()

    [<Fact>]
    member this.``acad 一段 LISP 等 EndLisp 才算求值完``() =
        Report.title output "acad　LISP 收口（EndLisp，真实例）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做，先看进程号再设环境变量")

        let pid = int pidText

        let t =
            task {
                let! acad = AutoCADTask.GetApplicationAsync pid

                let run (code: string) : Task<LispOutcome> =
                    STA.runAsync(fun () -> AutoCADTask.runLispAsync acad.ActiveDocument code)
                        .Unwrap()

                // 1. 纯赋值：EndLisp 到 = 整段求值结束
                let! first = run "(setq xunithost-probe 4242)"
                Report.line output "第一条" (sprintf "收尾=%A" first)
                Should.equal Completed first

                // 2. 里面套一条真命令：DIST 的 Begin/EndCommand 先走完，最后才是 EndLisp。
                //    用 DIST 不用 LINE：只读数、不往用户图纸里加实体。
                let! nested =
                    run "(progn (command \"_.DIST\" \"0,0\" \"10,0\"))"

                Report.line output "内嵌命令" (sprintf "收尾=%A" nested)
                Should.equal Completed nested

                // 3. 收尾：探针置 nil，别在用户图纸里留变量
                let! cleaned = run "(setq xunithost-probe nil)"
                Should.equal Completed cleaned

                Report.line output "残留" "全局变量 xunithost-probe 已置 nil"
                Report.conclusion output "通过：LISP 按 EndLisp 收口，内嵌命令也等整段求值结束"
            }

        t.GetAwaiter().GetResult()
