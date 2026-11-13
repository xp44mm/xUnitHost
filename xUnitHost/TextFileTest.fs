/// 文本读取的验证：临时造几种字节，逐份读回真实值（编码名 / 行数组）。
/// 全部在自己造的临时目录里，不碰任何现成文件。
namespace XUnitHost

open System
open System.IO
open System.Text
open Xunit
open FSharp.xUnit
open OsHost

type TextFileTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``文本读取 先认BOM再判UTF-8/GBK/ASCII``() =
        Report.title output "文本读取：BOM 优先，UTF-8 / GBK / ASCII 判歧义"

        let dir =
            Path.Combine(Path.GetTempPath(), sprintf "xUnitHost-TextFile-%d" Environment.ProcessId)

        Directory.CreateDirectory dir |> ignore
        let utf8 = UTF8Encoding(false)
        let text = "中文\r\n第二行\r\n"
        let rows = [| "中文"; "第二行" |]

        // 编码名是判定结果：| 表示两种解码都合法（不替人选边），? 表示严格解码都没过
        let samples: (string * byte[] * string * string[]) list =
            [ ("UTF-8 BOM",
               Array.append [| 0xEFuy; 0xBBuy; 0xBFuy |] (utf8.GetBytes text),
               "utf-8-bom",
               rows)
              ("UTF-8 无 BOM", utf8.GetBytes text, "utf-8", rows)
              ("GBK 无 BOM",
               [| 0xD6uy; 0xD0uy; 0xCEuy; 0xC4uy; 0x0Duy; 0x0Auy
                  0xB5uy; 0xDAuy; 0xB6uy; 0xFEuy; 0xD0uy; 0xD0uy; 0x0Duy; 0x0Auy |],
               "gbk",
               rows)
              ("UTF-16LE BOM",
               Array.append (UnicodeEncoding(false, true).GetPreamble()) (Encoding.Unicode.GetBytes text),
               "utf-16le-bom",
               rows)
              ("ASCII", utf8.GetBytes "abc\r\n", "ascii", [| "abc" |])
              ("歧义 C3A9", [| 0xC3uy; 0xA9uy |], "utf-8/gbk", [| "é" |])
              ("都不过 GBK占优", [| 0xD6uy; 0xD0uy; 0xCEuy |], "gbk/utf-8", [| "中\uFFFD" |])
              ("都不过 UTF-8占优",
               [| 0xF0uy; 0x28uy; 0x8Cuy; 0x28uy |],
               "utf-8/gbk",
               [| "\uFFFD(\uFFFD(" |])
              ("空文件", [||], "ascii", [||]) ]

        try
            for label, bytes, expectedEnc, expectedRows in samples do
                let path = Path.Combine(dir, label.Replace(" ", "_") + ".txt")
                File.WriteAllBytes(path, bytes)

                let enc, got = (TextFile.decodeAllLinesAsync path).GetAwaiter().GetResult()
                Report.line2 output "读回" label (sprintf "%s / %d 行 / %s" enc got.Length (String.Join("|", got)))
                Should.equal expectedEnc enc
                Should.equal (String.Join("|", expectedRows)) (String.Join("|", got))

            Report.conclusion output "通过：BOM 优先，UTF-8/GBK/ASCII 与歧义都写进编码名，都不过时写成 gbk/utf-8（前大后小），行尾按 \r?\n 切"
        finally
            if Directory.Exists dir then
                Directory.Delete(dir, true)
