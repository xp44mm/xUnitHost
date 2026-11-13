/// 「第二个进程」本人：跨进程那条 Fact 需要一个**真的另一个进程**，
/// 但不必为它再编一个程序 —— 拿这个测试程序自己再起一份、只跑这一条 Fact 就够了。
///
/// 它平时一直跳过（没有地址就不开口），只有父进程把地址递过来才动：
///
///   XUNITHOST_SAY_TO   地址        word:15324:1762340880123:...
///   XUNITHOST_SAY      要说的话     「；第二句：另一个进程凭地址接上的」
///   XUNITHOST_SAY_AT   回话文件     父进程读这个文件拿到「读回来的原文」
///
/// 规矩没变：只接**现成的那只**（`Host.resolve` 绝不新建）；说完就走，**不关 host**。
namespace XUnitHost

open System
open System.IO
open System.Text
open Xunit
open FSharp.xUnit

type SecondProcessTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``第二个进程：只凭地址接回那只 host，说一句，把读到的话写回去``() =
        let address = Environment.GetEnvironmentVariable "XUNITHOST_SAY_TO"
        let answerFile = Environment.GetEnvironmentVariable "XUNITHOST_SAY_AT"

        if String.IsNullOrEmpty address || String.IsNullOrEmpty answerFile then
            Assert.Skip("这是「另一个进程」的角色：由跨进程那条 Fact 按地址叫起来才说话，平时不跑")
        else
            let sentence =
                defaultArg (Option.ofObj (Environment.GetEnvironmentVariable "XUNITHOST_SAY")) ""

            // 回话写文件而不是打到 stdout：那是测试程序的输出，混着跑框架的话，父进程还得挑
            let answer (text: string) = File.WriteAllText(answerFile, text, UTF8Encoding false)
            let marks = Report.marks

            match Host.resolve address with
            | Error message -> failwithf "接不上 %s：%s" address message
            | Ok session ->
                output.WriteLine("接上了：{0}", Host.describe session)

                let said =
                    match session.Kind with
                    | Kind.Word ->
                        Host.execWord session (fun app ->
                            let doc = app.ActiveDocument
                            doc.Content.Text <- doc.Content.Text.Trim marks + sentence
                            doc.Content.Text.Trim marks)
                    | Kind.Excel ->
                        Host.execExcel session (fun app ->
                            let sheet =
                                app.ActiveWorkbook.Worksheets.[1]
                                :?> Microsoft.Office.Interop.Excel.Worksheet

                            let cell = sheet.Range "A1"
                            cell.Value2 <- (string cell.Value2).Trim() + sentence
                            (string cell.Value2).Trim())

                match said with
                | Error message -> failwithf "说不上话：%s" message
                | Ok(value, fresh) ->
                    // 这一步没动存盘状态，地址应该原样；变了自己也要知道
                    if fresh.Address <> address then
                        output.WriteLine("执行后地址已刷新：{0}", fresh.Address)

                    answer value
                    output.WriteLine("读到：{0}", value)
