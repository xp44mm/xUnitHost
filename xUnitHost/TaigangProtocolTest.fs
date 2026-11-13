/// 山西太钢技术协议：按黑板 PID 挂回用户正开着的那只 Word，只读地读出
/// 「4.2 系统性能要求」里燃烧器的燃料量与助燃风量（以及同类口径的其它行）。
///
/// 只读：不写、不保存、不关别人的 Word，也不新建实例。
namespace XUnitHost

open System
open Xunit
open FSharp.xUnit
open OsHost

type TaigangProtocolTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``山西太钢技术协议 读出燃烧器的燃料量与助燃风量``() =
        Report.title output "山西太钢技术协议　读出燃烧器的燃料量 / 助燃风量"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_WORD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_WORD_PID：这条要挂在活着的 Word 上，先看进程号再设环境变量")

        let pid = int pidText
        let before = Report.procIds WordHost.Host.SelfImage
        Report.line output "目标 PID" (string pid)

        let app = WordHost.WordTask.GetApplicationAsync pid |> Report.await

        // 关键字：燃烧器的燃料量与风量，兼看功率、耗量这类同行口径
        let keys = [ "燃料消耗"; "燃气量"; "助燃风量"; "风量"; "设计功率"; "热功率" ]

        let scene, diag, headers, hits =
            STA.run(fun () ->
                let doc = app.ActiveDocument
                let scene = sprintf "%s｜%s" doc.Name doc.FullName
                let diag = sprintf "表格 %d 张｜正文字符 %d" doc.Tables.Count doc.Content.Text.Length

                // 合并单元格的行访问 Cells 会抛，逐行 try 兜住
                let rows =
                    [ for t in 1 .. doc.Tables.Count do
                          let tab = doc.Tables.Item(t)

                          for r in 1 .. tab.Rows.Count do
                              try
                                  let row = tab.Rows.Item(r)

                                  let cells =
                                      [ for c in 1 .. row.Cells.Count ->
                                            WordHost.DocIO.normalize (row.Cells.Item(c).Range.Text) ]

                                  yield t, cells
                              with _ ->
                                  () ]

                // 表格对象取不到时（嵌套 / 文本框里的表）退回全文：单元格以 \r\a 收尾，按 \r 切、去掉 \a
                let fromText =
                    doc.Content.Text.Split('\r')
                    |> Array.map(fun s -> s.Replace("\a", "").Trim())
                    |> Array.filter(fun s -> s <> "")

                let hitRows =
                    rows
                    |> List.filter(fun (_, cells) ->
                        cells |> List.exists(fun c -> keys |> List.exists c.Contains))

                let hitText =
                    fromText
                    |> Array.mapi(fun i s -> i, s)
                    |> Array.filter(fun (_, s) -> keys |> List.exists s.Contains)
                    |> Array.map(fun (i, s) ->
                        let after = fromText.[i .. min (i + 3) (fromText.Length - 1)]

                        if rows.IsEmpty then
                            sprintf "正文 %d｜%s" i (String.Join(" → ", after))
                        else
                            sprintf "正文 %d｜%s" i s)

                let hit =
                    if hitRows.IsEmpty then
                        [ for t in 1 .. doc.Tables.Count -> t, [] ] |> List.truncate 0 // 占位，下面用 hitText
                    else
                        hitRows

                let firstOf t =
                    rows |> List.filter(fun (n, _) -> n = t) |> List.tryHead |> Option.map snd

                let heads =
                    hit |> List.map fst |> List.distinct |> List.choose(fun t -> firstOf t |> Option.map(fun h -> t, h))

                let allHits =
                    [ for t, cells in hit -> t, String.Join(" ｜ ", List.toArray cells)
                      for line in hitText -> 0, line ]

                scene, diag, heads, allHits)

        Report.line output "文档" scene
        Report.line output "现场" diag

        for t, head in headers do
            Report.line2 output "表头" (sprintf "表 %d" t) (String.Join(" ｜ ", List.toArray head))

        for t, cells in hits do
            Report.line2 output "命中" (if t = 0 then "全文" else sprintf "表 %d" t) cells

        // 4.2 那张表里必须有「燃料消耗」和「助燃风量」两行
        let labeled (word: string) =
            hits |> List.filter(fun (_, cells) -> cells.Contains word)

        Report.line output "燃料消耗行数" (string (labeled "燃料消耗").Length)
        Report.line output "助燃风量行数" (string (labeled "助燃风量").Length)

        Should.equal true ((labeled "燃料消耗").Length >= 1)
        Should.equal true ((labeled "助燃风量").Length >= 1)

        Should.equal before (Report.procIds WordHost.Host.SelfImage)
        Report.conclusion output "通过：只读读出燃烧器的燃料量与助燃风量，Word 还是那一只，没保存、没关"
