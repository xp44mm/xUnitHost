/// 两种执行方式。
///
/// 方式一（一条命令）：大模型写一个命令 → 起 host → 执行 → 关掉，host 不留。
///     `WordHost.Host.exec` / `execQuery`，统一入口 `Host.execOnce` / `execOnceQuery`。
///
/// 方式二（合作完成）：host 开一次、连着用几步。每步只做一小段，做完**不关**，
///     把「地址 + 结果」交回大模型；大模型据此写下一步，凭同一个地址回到同一个 host；
///     最后大模型说完了，按**第一步之前的状态**收尾（自己起的关掉，本来就在的不动）。
///     `Host.openSession` / `execWord` / `execExcel` / `close`。
///
///     地址是**跨进程**有效的，而且带着一整套指纹：
///
///         word:15324:1762340880123:1762340987654:3f5a91c2:w
///              pid   进程启动毫秒   最后保存毫秒   文档路径指纹 归属
///
///     回来时逐项核对（`Host.resolve`）：pid 被复用、文档换了、中间被别人另存过一次，
///     都会被挡下来并告诉你「这只 host 现在的地址是……」。见后面两个 Fact。
namespace XUnitHost

open System
open System.Diagnostics
open System.IO
open Xunit
open FSharp.xUnit

/// 方式一 / 方式二，一手一脚走一遍。
type CollaborationTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``方式一 一条命令：开-做-关，host 不留``() =
        let before = Report.procIds "WINWORD"
        Report.requireClean "WINWORD"

        Report.title output "方式一　一条命令：开 → 做 → 关，host 不留"
        Report.line output "起点" "本机没有 Word 在跑，从零起一个"
        Report.line output "命令" "开 Word → 新建文档 → 写一句 → 读回来 → 关掉"

        // 那个「命令」：起 Word、写一句、读回来。全程一次进程、一个实例，跑完就退出
        let text =
            Host.execOnceQuery Kind.Word (fun app ->
                match app with
                | WordApp word ->
                    let doc = word.Documents.Add()
                    doc.Content.Text <- "一条命令：写完就走"
                    doc.Content.Text.Trim Report.marks
                | _ -> failwith "kind 不对")
            |> Report.ok

        Report.line output "结果" (sprintf "「%s」" text)
        // 读回来的就是写进去的：这一条命令确实落到了 Word 里
        Should.equal "一条命令：写完就走" text

        // 关得干净：开之前没有 Word，跑完也不该留着
        let after = Report.procIds "WINWORD"
        Report.line output "收尾" (sprintf "Word 进程 %d 个（跑前 %d / 跑后 %d）" after.Count before.Count after.Count)
        Assert.True(Set.isEmpty(Set.difference after before), "Word 留了残进程")
        Report.conclusion output "通过：一条命令自己起、自己收，host 不留"

    [<Fact>]
    member this.``方式二 合作完成：host 不关、凭地址回来接着做、最后按原状态收尾``() =
        let beforeWord = Report.procIds "WINWORD"
        let beforeExcel = Report.procIds "EXCEL"
        Report.requireClean "WINWORD"
        Report.requireClean "EXCEL"

        Report.title output "方式二　合作完成：host 不关、凭地址回来接着做、按原状态收尾"
        Report.line output "起点" "Word / Excel 都没有在跑，从零各起一个"

        // ---- action1：开 host，写第一句，把「地址 + 结果」交回去 ----
        let word1 = Host.openSession Kind.Word
        let wordPid = Set.minElement word1.Pids

        let action1, word2 =
            Host.execWord word1 (fun app ->
                let doc = app.Documents.Add()
                doc.Content.Text <- "第一句：Word 先开口"
                doc.Content.Text.Trim Report.marks)
            |> Report.ok

        Report.line2 output "action1" "开 host" word1.Address
        Report.line2 output "action1" "结果" (sprintf "「%s」" action1)
        // host 没关：那个进程还在（这正是「action1 做完不关 host」）
        Report.line2 output "action1" "交回" (sprintf "地址 + 结果；host 没关，进程 %d 还在" wordPid)
        Assert.Contains(wordPid, Report.procIds "WINWORD")

        // ---- 这里是分水岭：大模型手上只有「地址 + 结果」，据此写 action2 ----
        // ---- action2：凭地址回到同一个 host，接着那份文档往下写 ----
        let action2, word3 =
            Host.execByAddress word2.Address (fun app ->
                match app with
                | WordApp app ->
                    let doc = app.ActiveDocument
                    doc.Content.Text <- (doc.Content.Text.Trim Report.marks) + "；第二句：大模型让 Word 接着说"
                    doc.Content.Text.Trim Report.marks
                | _ -> failwith "这个地址指向的不是 Word")
            |> Report.ok

        Report.line2 output "action2" "凭地址回来" (sprintf "%s（还是那只 Word，地址没变）" word3.Address)
        Report.line2 output "action2" "结果" (sprintf "「%s」" action2)
        // 第二句是接在第一句后面的：说明接的是同一份文档，不是新开一份
        Assert.Contains("第一句", action2)
        Assert.Contains("第二句", action2)

        // ---- action3：再拉一个 Excel 进来，两个 host 同时开着 ----
        let excel1 = Host.openSession Kind.Excel
        let excelPid = if excel1.Owned then Some(Set.minElement excel1.Pids) else None

        let action3, _ =
            Host.execByAddress
                excel1.Address
                (fun app ->
                    match app with
                    | ExcelApp app ->
                        let book = app.Workbooks.Add()

                        let sheet =
                            book.Worksheets.[1] :?> Microsoft.Office.Interop.Excel.Worksheet

                        let cell = sheet.Range("A1")
                        cell.Value2 <- action2
                        (string cell.Value2).Trim()
                    | _ -> failwith "这个地址指向的不是 Excel")
            |> Report.ok

        Report.line2 output "action3" "另开 host" excel1.Address
        Report.line2 output "action3" "结果" (sprintf "Excel A1 = 「%s」" action3)
        Assert.StartsWith("第一句", action3)

        // ---- 大模型判断：任务完成 → 按 action1 之前的状态收尾 ----
        Assert.True((Host.findByAddress word3.Address).IsSome, "地址该还活着")
        Host.close word3
        Host.close excel1

        Report.line2 output "收尾" "关 host" "自己起的两个都关了，地址随之失效"
        Report.line2 output "收尾" "残留进程" (sprintf "WINWORD %d 个 / EXCEL %d 个" (Report.procIds "WINWORD").Count (Report.procIds "EXCEL").Count)

        Assert.True((Host.findByAddress word3.Address).IsNone, "关掉之后地址该失效")
        Assert.DoesNotContain(wordPid, Report.procIds "WINWORD")
        excelPid |> Option.iter(fun pid -> Assert.DoesNotContain(pid, Report.procIds "EXCEL"))
        Assert.True(Set.isEmpty(Set.difference (Report.procIds "WINWORD") beforeWord), "Word 留了残进程")
        Assert.True(Set.isEmpty(Set.difference (Report.procIds "EXCEL") beforeExcel), "Excel 留了残进程")

        Report.conclusion output "通过：三个 action 串下来，Word / Excel 各一只，收尾干净"

    [<Fact>]
    member this.``方式二 跨进程：另一个进程只拿地址回来，说的还是同一份文档``() =
        let before = Report.procIds "WINWORD"
        Report.requireClean "WINWORD"

        Report.title output "方式二　跨进程：另一个进程只拿地址回来，说的还是同一份文档"
        Report.line output "起点" "本机没有 Word 在跑，从零起一个"

        // ---- action1：本进程开 host、写第一句，拿到「地址 + 结果」 ----
        let word1 = Host.openSession Kind.Word
        let wordPid = Set.minElement word1.Pids

        let action1, word2 =
            Host.execWord word1 (fun app ->
                let doc = app.Documents.Add()
                doc.Content.Text <- "第一句：开 host 的这个进程写的"
                doc.Content.Text.Trim Report.marks)
            |> Report.ok

        Report.line2 output "action1" "本进程" (sprintf "pid %d 开 host，写第一句「%s」" wordPid action1)
        Report.line2 output "action1" "交出去" (sprintf "只有地址 %s" word2.Address)

        // ---- action2：换一个真·独立的进程（还是这个测试程序，换一条 Fact 跑），手上只有地址 ----
        // 回话走文件：子进程的屏幕输出是跑框架的东西，混着不好挑，文件里那一行才是它读到的话
        let answer =
            Path.Combine(Path.GetTempPath(), "xunithost-say-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".txt")

        let psi = ProcessStartInfo()
        psi.FileName <- Report.selfExe ()
        psi.WorkingDirectory <- AppContext.BaseDirectory
        psi.RedirectStandardOutput <- true
        psi.RedirectStandardError <- true
        psi.UseShellExecute <- false
        psi.StandardOutputEncoding <- System.Text.Encoding.UTF8
        psi.StandardErrorEncoding <- System.Text.Encoding.UTF8
        // 只跑「第二个进程」那一条 Fact：它平时是跳过的，这里把地址、要说的话、回话文件塞给它
        psi.ArgumentList.Add "-method"
        psi.ArgumentList.Add "*第二个进程*"
        psi.Environment.["XUNITHOST_SAY_TO"] <- word2.Address
        psi.Environment.["XUNITHOST_SAY"] <- "；第二句：另一个进程凭地址接上的"
        psi.Environment.["XUNITHOST_SAY_AT"] <- answer

        let child = Process.Start psi
        // stdout 一边抽一边等，stderr 同步读：两头管道都得有人读，不然写满就僵住
        let childOut = child.StandardOutput.ReadToEndAsync()
        let childErr = child.StandardError.ReadToEnd()

        if not (child.WaitForExit 120000) then
            (try
                child.Kill()
             with _ ->
                 ())

            failwith "子进程两分钟没结束"

        let said =
            if File.Exists answer then
                try
                    File.ReadAllText(answer, System.Text.Encoding.UTF8).Trim()
                finally
                    (try
                        File.Delete answer
                     with _ ->
                         ())
            else
                ""

        Report.line2 output "action2" "子进程" (sprintf "pid %d 只凭地址接上，读到「%s」" child.Id said)

        if child.ExitCode = 0 then
            Report.line2 output "action2" "回话" (sprintf "写进 %s 那一行，就是它读回来的原文" (Path.GetFileName answer))
        else
            Report.line2 output "action2" "子进程说" ((childOut.Result.Trim()) + " " + (childErr.Trim()))

        // 确实是两个不同的进程
        Assert.NotEqual(wordPid, child.Id)
        Assert.True(child.ExitCode = 0, sprintf "子进程没跑通（exit=%d）：%s" child.ExitCode (childErr.Trim()))
        // 子进程读到的第一句说明它接的是同一份文档，不是自己新开一份
        Assert.StartsWith(action1, said)
        Assert.Contains("第二句", said)

        // ---- 回到本进程：还是那只 Word、那份文档，内容被子进程改了 ----
        Assert.Contains(wordPid, Report.procIds "WINWORD")

        let local, word3 =
            Host.execWord word2 (fun app -> app.ActiveDocument.Content.Text.Trim Report.marks)
            |> Report.ok

        Report.line2 output "回到本进程" "读到" (sprintf "「%s」" local)
        Report.line2 output "回到本进程" "地址" (sprintf "%s（没变）" word3.Address)
        // 两个进程看到的是同一段文字，而且地址一字不差
        Should.equal said local
        Should.equal word2.Address word3.Address

        // 陈地址（进程早没了）要老实报错，不能糊里糊涂新建一个
        match Host.resolve "word:999999" with
        | Ok _ -> failwith "不存在的进程号不该解析成功"
        | Error message -> Report.line2 output "失效地址" "word:999999" message

        // ---- 收尾：按 action1 之前的状态关掉 ----
        Host.close word3
        Report.line2 output "收尾" "关 host" (sprintf "自己起的 Word 关了，残留 %d 个" (Report.procIds "WINWORD").Count)
        Assert.DoesNotContain(wordPid, Report.procIds "WINWORD")
        Assert.True(Set.isEmpty(Set.difference (Report.procIds "WINWORD") before), "Word 留了残进程")
        Report.conclusion output "通过：两个真·独立进程，说的还是同一份文档"

    [<Fact>]
    member this.``指纹：地址带着最后一次实质性保存的时间，保存过之后老地址就作废``() =
        let before = Report.procIds "WINWORD"

        // 这个测试从头盯到尾：开新的 Word、新文档、存一次盘。起点必须干净
        Report.requireClean "WINWORD"

        let dir =
            Path.Combine(Path.GetTempPath(), "xUnitHost-指纹-" + Guid.NewGuid().ToString("N").Substring(0, 8))

        Directory.CreateDirectory dir |> ignore
        let path = Path.Combine(dir, "指纹测试.docx")

        Report.title output "指纹：地址带着最后一次实质性保存的时间，保存过之后老地址就作废"
        Report.line output "起点" "本机没有 Word 在跑，从零起一个"

        let word1 = Host.openSession Kind.Word
        let wordPid = word1.Stamp.Pid
        let owned = word1.Owned

        // 不管中间哪一步断言炸了，都要把 host 收掉，别留一个 Word 拖累后面的测试
        let mutable latest = word1

        try
            // ---- 新文档：还没存过盘，地址里保存时间是 0、路径指纹是 "-" ----
            Report.line2 output "开 host" "地址" word1.Address
            Report.line2 output "开 host" "说明" "新文档：还没存过盘，保存时间 0、路径指纹 -"
            Assert.True(word1.Stamp.DocPath.IsNone, "新文档不该有路径")
            Should.equal 0L word1.Stamp.SavedMs
            Assert.True(word1.Stamp.Pid > 0, "地址里该有进程号")
            Assert.True(word1.Stamp.StartedMs > 0L, "地址里该有进程启动时间")

            // ---- 保存一次：地址里就带上「最后一次实质性保存」的时刻 ----
            let _, word2 =
                Host.execWord word1 (fun app ->
                    let doc = app.Documents.Add()
                    doc.Content.Text <- "这一句会被存到盘上"

                    let mutable fileName: obj = box path
                    doc.SaveAs(&fileName) |> ignore
                    doc.Content.Text.Trim Report.marks)
                |> Report.ok

            latest <- word2
            let saved = word2.Stamp.SavedMs

            Report.line2 output "保存后" "地址" word2.Address
            Report.line2 output "保存后" "说明" (sprintf "锚定 %s，最后保存 %s" (Path.GetFileName path) (Report.clock saved))
            Assert.True(saved > 0L, "保存过之后该有保存时间")
            Should.equal (Some(Path.GetFullPath path)) word2.Stamp.DocPath
            // 地址里确实带着这个时刻，而且和磁盘上的文件时间对得上
            Assert.Contains(string saved, word2.Address)
            Should.equal saved (Report.unixMs (FileInfo(path).LastWriteTimeUtc))
            // 保存把文档换了（从「没路径」到「有路径」），地址必须变
            Assert.True(word1.Address <> word2.Address, "保存把文档换了，地址必须变")

            // ---- 老地址（保存之前那个）严格核对 → 拒：文档和保存时间都不对了 ----
            match Host.resolve word1.Address with
            | Ok _ -> failwith "保存过之后，老地址不该还能用"
            | Error message ->
                Report.line2 output "老地址" "被拒" message
                Assert.Contains("现在的地址", message)

            // ---- 新地址严格核对 → 过，指纹完全一致 ----
            let back = Report.ok (Host.resolve word2.Address)
            Should.equal word2.Stamp.SavedMs back.Stamp.SavedMs
            Should.equal word2.Stamp.DocPath back.Stamp.DocPath
            Assert.True(back.Drift.IsEmpty, "刚发的地址不该有漂移")
            Report.line2 output "新地址" "核对" "通过：进程、启动时间、保存时间、路径指纹全对上，无漂移"

            // ---- 模拟「中间别人又存了一次」：文件时间变了 ----
            let tampered = (FileInfo path).LastWriteTimeUtc.AddSeconds 5.0
            File.SetLastWriteTimeUtc(path, tampered)

            // 严格核对：拒（这就是「更保险」要的效果）
            match Host.resolve word2.Address with
            | Ok _ -> failwith "文件被别人存过，地址该作废"
            | Error message ->
                Report.line2 output "外部改动" "被拒" message
                Assert.Contains("最后一次保存变了", message)

            // 宽松核对：只要还是那只进程就放行，但把漂移记下来，并给出新地址
            let drifted = Report.ok (Host.resolveWith Drift.Lenient word2.Address)
            latest <- drifted

            Report.line2 output "宽松接回" "放行" (sprintf "%s（漂移 %d 项）" drifted.Address drifted.Drift.Length)
            Report.line2 output "宽松接回" "漂移" (String.Join("；", drifted.Drift))
            Assert.True(word2.Address <> drifted.Address, "漂移之后该给出新地址")
            Should.equal 1 drifted.Drift.Length
        finally
            // ---- 收尾 ----
            Host.close latest

            (try
                File.Delete path
             with _ ->
                 ())

            (try
                Directory.Delete(dir, true)
             with _ ->
                 ())

        let rest = if owned then (Report.procIds "WINWORD").Count else -1
        Report.line2 output "收尾" "关 host" (sprintf "文件删了，Word 残留 %d 个" rest)

        if owned then
            Assert.DoesNotContain(wordPid, Report.procIds "WINWORD")
            Assert.True(Set.isEmpty(Set.difference (Report.procIds "WINWORD") before), "Word 留了残进程")

        Report.conclusion output "通过：保存过之后地址自动变，老地址一律作废"

    [<Fact>]
    member this.``按地址回来只打招呼：没有实例就报错，绝不新建一个``() =
        let before = Report.procIds "EXCEL"
        Report.requireClean "EXCEL"

        Report.title output "按地址回来只打招呼：没有实例就报错，绝不新建一个"
        Report.line output "前提" "本机一个 Excel 都没有"
        Report.line output "拿到的地址" "excel:0:0:0:-:w（格式合法，但没有人接）"

        // 地址语法没问题，但本机一个 Excel 都没有：这时候只能报错。
        // 要是这里走了「没有就新建」的老路，就会凭空多出一个 EXCEL 进程，
        // 而调用方还以为自己接回了原来那只。
        match Host.resolve "excel:0:0:0:-:w" with
        | Ok session -> failwithf "没有实例时不该接上（却拿到了 %s）" (Host.describe session)
        | Error message -> Report.line output "结果" message

        let after = Report.procIds "EXCEL"
        Report.line output "核对" (sprintf "EXCEL 进程 %d 个（跑前 %d / 跑后 %d），没有凭空造一个出来" after.Count before.Count after.Count)
        Assert.True(Set.isEmpty(Set.difference after before), "按地址回来竟然起了一个 Excel")
        Report.conclusion output "通过：接不上就如实报错，绝不新建"
