/// 热风炉计算书摘要：先把「焦炉煤气燃烧计算书.xlsx」整本只读转储成 TSV，
/// 有了真值再往 Word 里写摘要（摘要里每一个数都来自这次转储，不凭记忆）。
///
/// 只读：按 PID 挂回用户正开着的那只 Excel，一个格都不改、不 Save、不 Quit。
///
/// 跑法：
///   XUNITHOST_EXCEL_PID=11468 xUnitHost/bin/Debug/net10.0/xUnitHost.exe \
///     -method "*计算书整本转储*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.Globalization
open System.IO
open System.Text
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal StoveBook =

    /// 用户正开着的那本计算书（换文件就改这里）
    let target = @"C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx"

    /// 转储落地处：TestResults 下，一份一个文件，UTF-8 with BOM + CRLF
    let dumpPath =
        @"C:\Application Data\GitHub\xp44mm\xUnitHost\TestResults\计算书-全表转储.tsv"

    let txt (v: obj) =
        if isNull v then "" else Convert.ToString(v, CultureInfo.InvariantCulture)

    /// 列号 → 列字母（1 → A，27 → AA）
    let colName (n: int) =
        let rec loop n acc =
            if n <= 0 then
                acc
            else
                let r = (n - 1) % 26
                loop ((n - 1) / 26) (string (char (65 + r)) + acc)

        loop n ""

    /// 批量读回来的可能是二维数组，也可能是单格的标量；统一成「第 i 行第 j 列」
    let grid (rows: int) (cols: int) (raw: obj) : (int -> int -> obj) =
        match raw with
        | :? (obj[,]) as a -> fun i j -> a.[i, j]
        | other -> fun i j -> if i = 1 && j = 1 then other else null

/// 摘要正文：内容全是「焦炉煤气燃烧计算书.xlsx」这次读到的现场值（见 TestResults\计算书-全表转储.tsv）。
/// 只列输入、里程碑节点与输出，不抄过程；数字与表内取值一一对应，改表后重跑本 Fact 就重写这一份。
module internal AbstractDoc =

    /// Word 里开着的那一份（换文件就改这里）
    let path = @"C:\崔胜利\群有\中冶南方都市环保\热风炉计算书摘要.docx"

    /// 文档块：标题 / 章节 / 正文 / 表格
    type Block =
        | Title of string
        | Head of string
        | Para of string
        | Table of string[] * string[][]

    let blocks: Block list =
        [ Title "热风炉（烟气补燃）计算书摘要"

          Para
              "依据《焦炉煤气燃烧计算书.xlsx》六张表（焦炉煤气 / 需要加热的烟气参数 / 烟气成分计算 / 烟气吸热量计算 / 补热燃烧计算 / 燃烧器选型），取 2026-11-11 文件当前的计算值。本摘要只列输入、里程碑节点与输出，推导过程见原表。"

          Para
              "结论：57000 Nm³/h 烟气自 175℃ 加热至 500℃，需补燃 637.45 万大卡/h，耗焦炉煤气 2093 Nm³/h（α=1.15），配 920.93 万大卡/h 套管式燃烧器，助燃风机 12000 m³/h、4.5 kPa。助燃风总管口径与风机压升两处尚不自洽，见第四节。"

          Head "一、设计输入"

          Table(
              [| "序号"; "项目"; "数值"; "出处" |],
              [| [| "1"; "被加热烟气量"; "57000 Nm³/h"; "需要加热的烟气参数!D3" |]
                 [| "2"; "烟气温度"; "175℃ → 500℃"; "D4 / D5" |]
                 [| "3"; "烟气压力"; "0.5~1 kPa"; "D6" |]
                 [| "4"; "烟气含氧量（干基）"; "8~10%，计算取 9%"; "D9 / 烟气成分计算!C6" |]
                 [| "5"; "烟道与布置"; "φ2020 管道式，燃烧器布置在烟气管道上"; "E11" |]
                 [| "6"; "业主给定热交换量"; "530 万大卡/h（只作对照，未采用）"; "D2" |]
                 [| "7"; "燃料"; "焦炉煤气，LHV 4000 kcal/Nm³ = 16747.2 kJ/Nm³"; "焦炉煤气!E3" |]
                 [| "8"; "煤气压力"; "4 kPa（区间 4~8）"; "焦炉煤气!E2" |]
                 [| "9"
                    "煤气成分（干基 vol%）"
                    "H2 57 / CH4 24.5 / CO 6.5 / N2 6.45 / C2 以上 2.5 / CO2 2.5 / O2 0.55；水分忽略"
                    "焦炉煤气!E4:E11" |]
                 [| "10"; "大气条件（湖北仙桃）"; "100760 Pa / 16.5℃ / 相对湿度 80%"; "烟气成分计算!C3:C5" |]
                 [| "11"; "助燃空气"; "湿空气 H2O 1.491 / O2 20.687 / N2 77.822 vol%"; "烟气成分计算!C7:C9" |]
                 [| "12"; "补燃过量空气系数 α"; "1.15（补燃自身助燃风，可改）"; "补热燃烧计算!C23" |] |]
          )

          Head "二、重要里程碑节点"

          Para
              "节点 1　成分口径定稿：以「焦炉煤气」E 列设计值为唯一口径，其后三张表全部活链接，N2 平衡到 100%（6.45%）。成分一改，α、烟气成分、吸热量、耗量、选型全链自动重算。"

          Para
              "节点 2　α 由含氧量反算，不是取值：干基 O2 9% 闭式解得 α = 1.6732（理论需氧 0.877 Nm³/Nm³、理论空气量 4.2394 Nm³/Nm³）。这里不能照天然气燃烧器的 1.15 取，取了烟气成分与吸热量都会失真。"

          Para
              "节点 3　被加热烟气成分定稿：湿基 H2O 15.635 / O2 7.593 / N2 71.821 / CO2 4.951 vol%；干基 O2 9.000 / N2 85.131 / CO2 5.869 vol%；湿基、干基各合计 100%。"

          Para
              "节点 4　吸热量（焓差口径）：57000 Nm³/h（2543.05 kmol/h）自 175℃ 到 500℃，总焓 3404.02 万 kJ/h → 6072.90 万 kJ/h，吸热 2668.88 万 kJ/h = 7413.6 kW = 637.45 万大卡/h。"

          Para
              "节点 5　独立复算：用 NIST Shomate 系数逐项算（不经 XLL）得 2659.66 万 kJ/h，与主表差 −0.35%，判为一致。"

          Para
              "节点 6　与业主数对照：计算值 637.45 ÷ 给定值 530 = 1.203。业主的 530 原格保留不改，设计按计算值走。"

          Para
              "节点 7　放热口径改定：每 Nm³ 煤气有效放热 = LHV 16747.2 − 产物 0→500℃ 显热 3995.84 = 12751.36 kJ/Nm³（3045.61 kcal/Nm³）。原「绝热温度反解 + 二分」已删除，两法数值等价。"

          Para
              "节点 8　两表互校：补燃表煤气耗量 2093.02 Nm³/h 与选型表 2093.02 Nm³/h 偏差 0%；助燃空气量 10204.11 对 10204.11，偏差 1.8×10⁻¹⁴%。两条路径同源不同走法，互为校核。"

          Para
              "节点 9　燃烧器定档：设计热功率 837.21 万大卡/h（9.74 MW）× 裕量 1.10 = 额定 920.93 万大卡/h（10.71 MW），设计点负荷率 90.9%，落在 80~95% 的合理区间。"

          Para
              "节点 10　风机定档：进风 35℃ 理论风量 11902.55 m³/h → 选型 12000 m³/h、压升 4.5 kPa、效率 60% → 轴功率 25 kW → 电机 30 kW（裕量 1.2）。"

          Head "三、主要输出结果"

          Table(
              [| "序号"; "项目"; "数值" |],
              [| [| "1"; "补燃煤气耗量"; "2093.0 Nm³/h" |]
                 [| "2"; "助燃空气量（α=1.15、湿空气）"; "10204.1 Nm³/h" |]
                 [| "3"; "补燃湿烟气量"; "11632.6 Nm³/h（其中干烟气 9157.2 Nm³/h）" |]
                 [| "4"; "补燃烟气湿基成分"; "H2O 21.28 / O2 2.37 / N2 69.43 / CO2 6.93 vol%" |]
                 [| "5"; "燃烧器额定热负荷"; "920.93 万大卡/h = 3855.7 万 kJ/h = 10.71 MW" |]
                 [| "6"; "额定煤气耗量"; "2302.3 Nm³/h（20℃ 工况 2470.9 m³/h）" |]
                 [| "7"; "额定助燃空气量"; "11224.5 Nm³/h（35℃ 工况 11902.6 m³/h）" |]
                 [| "8"; "煤气支管"; "DN250（内径 260 mm）→ 12.93 m/s" |]
                 [| "9"; "套管环隙（助燃风出口）"; "外内径 518 / 内管外径 273 → 0.1522 m²，21.72 m/s" |]
                 [| "10"; "助燃风机"; "12000 m³/h、4.5 kPa、效率 60%、轴功率 25 kW、电机 30 kW" |]
                 [| "11"; "煤气密度 / 平均分子量"; "0.4757 kg/Nm³ / 10.662 g/mol" |] |]
          )

          Head "四、待确认事项"

          Para
              "1　助燃风总管口径与风机压升不自洽：C47 取内径 350 mm（同表 D47 备注写 DN600），按 350 算管内流速 34.36 m/s，超出 10~15 m/s 经验区间；风压分项合计因此到 9.996 kPa，与取用压升 4.5 kPa 差 5.50 kPa（同表要求 ≤10%）。两处须取一：总管放大到 DN600 一档，或风机按 10 kPa 档选。"

          Para "2　燃烧器本体阻力 2500 Pa 是初估，须厂家给数（套管式常见 2~4 kPa）。"

          Para "3　风压分项里的管长 30 m、摩擦系数 λ 0.02、局部阻力系数合计 8，按管路图核。"

          Para "4　风机效率 60%、电机标准档（30 kW）按厂家样本核。"

          Para
              "5　流量裕量口径重复：C70 由 12000 ÷ 11902.55 得 1.008，另一处 E70 取 1.3；若样本风量已含裕量，两处应统一收敛到 1.10~1.20。"

          Head "五、计算口径"

          Para
              "· 焓值按 NIST Shomate 摩尔定压热容、基准 0℃，Δh = [A·Δt + B·Δt²/2 + C·Δt³/3 + D·Δt⁴/4 − E·Δ(1/t)] × 1000 kJ/kmol；主算走 gasMoleEnthalpy，另设系数列逐项复算并列。"

          Para "· 湿空气按仙桃大气（100760 Pa / 16.5℃ / 80%）算得，含 H2O 1.49%；干空气 21/79 的旧口径已于 2026-11-10 删除，本表为唯一口径。"

          Para "· 反应物（燃料与助燃空气）按 0℃ 基准，显焓忽略（影响小于 1%）。"

          Para "· 燃料 S / Cl / F / 灰分无数据按 0 计；C2 以上不饱和烃按 C2H4 计；煤气按干基。"

          Para "· 放热量不乘炉体热效率，散热与积灰由选型裕量 1.10 一并覆盖。" ]

/// 三条 Fact：只读转储 → 认人 → 写摘要（写之前先在磁盘留一份带时间戳的备份）
type StoveSummaryTest(output: ITestOutputHelper) =
    /// 只读转储：整本 6 张表逐格落成 TSV，报告里给出每张表的使用区域与非空格数。
    [<Fact>]
    member this.``计算书整本转储 只读不改一个格``() =
        Report.title output "计算书摘要　把「焦炉煤气燃烧计算书.xlsx」整本转储"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_EXCEL_PID：这个 Fact 只挂回你正开着的那只 Excel")

        let pid =
            match Int32.TryParse pidText with
            | true, n -> n
            | _ -> failwithf "XUNITHOST_EXCEL_PID 不是整数：%s" pidText

        let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await
        let back = STA.run (fun () -> ExcelHost.Host.pidOf app)
        Should.equal (Some pid) back
        Report.line output "挂回" (sprintf "PID %d（只读，不新建、不 Save、不 Quit）" pid)

        let lines, sheets =
            STA.run (fun () ->
                let book =
                    [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                    |> List.map (fun i -> app.Workbooks.Item(i))
                    |> List.tryFind (fun b ->
                        let full =
                            try
                                b.FullName
                            with _ ->
                                ""

                        String.Equals(full, StoveBook.target, StringComparison.OrdinalIgnoreCase))

                let book =
                    match book with
                    | Some b -> b
                    | None -> failwithf "这只 Excel 里没打开「%s」" StoveBook.target

                let sheets =
                    [ 1 .. (try book.Worksheets.Count with _ -> 0) ]
                    |> List.map (fun i -> book.Worksheets.Item(i) :?> Worksheet)

                let out = ResizeArray<string>()
                out.Add("工作表\t地址\t内容\t公式")

                let info =
                    sheets
                    |> List.map (fun ws ->
                        let used = ws.UsedRange
                        let r0 = used.Row
                        let c0 = used.Column
                        let rows = used.Rows.Count
                        let cols = used.Columns.Count
                        let values = StoveBook.grid rows cols used.Value2

                        let formulas =
                            let raw =
                                try
                                    used.Formula2
                                with _ ->
                                    null

                            StoveBook.grid rows cols raw

                        let mutable filled = 0

                        for i in 1 .. rows do
                            for j in 1 .. cols do
                                let text = StoveBook.txt (values i j)
                                let formula = StoveBook.txt (formulas i j)

                                if text <> "" || formula.StartsWith "=" then
                                    filled <- filled + 1

                                    let addr =
                                        sprintf "%s%d" (StoveBook.colName (c0 + j - 1)) (r0 + i - 1)

                                    out.Add(
                                        sprintf
                                            "%s\t%s\t%s\t%s"
                                            ws.Name
                                            addr
                                            text
                                            (if formula.StartsWith "=" then formula else "")
                                    )

                        ws.Name, sprintf "%s%d:%s%d" (StoveBook.colName c0) r0 (StoveBook.colName (c0 + cols - 1)) (r0 + rows - 1), filled)

                List.ofSeq out, info)

        Directory.CreateDirectory(Path.GetDirectoryName StoveBook.dumpPath) |> ignore
        File.WriteAllText(StoveBook.dumpPath, String.concat "\r\n" lines + "\r\n", UTF8Encoding true)

        Report.line output "工作表数" (sprintf "%d" sheets.Length)

        for name, used, filled in sheets do
            Report.line2 output "工作表" name (sprintf "使用区域 %s，非空格 %d" used filled)

        Report.line output "转储" (sprintf "%s（%d 行）" StoveBook.dumpPath lines.Length)
        Report.conclusion output "通过：整本计算书已落成本地 TSV，计算书本身一个格都没动"

    /// 摘要正文的落点：Word 里打开的「热风炉计算书摘要.docx」
    [<Fact>]
    member this.``摘要 认人 看清 Word 里开的是哪一份``() =
        Report.title output "计算书摘要　认人：Word 现在开着哪一份"

        let docs =
            STA.run (fun () ->
                match WordHost.Host.TryGetRunningApplication() with
                | None -> []
                | Some app ->
                    let pid = WordHost.Host.pidOf app

                    Report.line
                        output
                        "Word"
                        (sprintf
                            "PID %s（只接现成实例，不新建）"
                            (match pid with
                             | Some p -> string p
                             | None -> "—"))

                    // Documents.Item 收 obj& ，这里直接遍历集合，省掉按引用传参
                    [ for raw in app.Documents do
                          let doc = raw :?> Microsoft.Office.Interop.Word.Document

                          let path =
                              try
                                  doc.FullName
                              with _ ->
                                  ""

                          let saved =
                              try
                                  doc.Saved
                              with _ ->
                                  true

                          path, saved, (try doc.Content.Text.Length with _ -> 0) ])

        for path, saved, len in docs do
            Report.line output "文档" (sprintf "「%s」字数 %d，未保存改动 %b" path len (not saved))

        Report.conclusion output "通过：看清了 Word 里开着哪些文档（这一步只读）"

    /// 真正落笔：磁盘备份 → 认人（文件名必须是摘要那一份）→ 清空正文 → 逐块写 → Save → 读回核对。
    [<Fact>]
    member this.``摘要 清空并重写 Word 正文``() =
        Report.title output "计算书摘要　写进 Word 里开着的那一份（写前先备份）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_WORD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_WORD_PID：这个 Fact 只往你正开着的那一份里写")

        let pid =
            match Int32.TryParse pidText with
            | true, n -> n
            | _ -> failwithf "XUNITHOST_WORD_PID 不是整数：%s" pidText

        if not (File.Exists AbstractDoc.path) then
            failwithf "磁盘上没有「%s」" AbstractDoc.path

        // 1. 备份：同目录留一份带时间戳的副本，重写之后原件才允许被改
        let backup =
            Path.Combine(
                Path.GetDirectoryName AbstractDoc.path,
                sprintf "热风炉计算书摘要-bak%s.docx" (DateTime.Now.ToString "yyyyMMdd-HHmmss")
            )

        File.Copy(AbstractDoc.path, backup, false)
        Report.line output "磁盘备份" (sprintf "%s（%d 字节）" backup (FileInfo backup).Length)

        // 2. 在固定 STA 线程上挂回现成 Word、找到那一份、重写
        let result =
            STA.run (fun () ->
                let app =
                    match WordHost.Host.TryGetRunningApplication() with
                    | Some a -> a
                    | None -> failwith "本机没有在跑的 Word：这个 Fact 只往现成那一份里写，不新建"

                let back = WordHost.Host.pidOf app
                Should.equal (Some pid) back

                let doc =
                    [ for raw in app.Documents do
                          raw :?> Microsoft.Office.Interop.Word.Document ]
                    |> List.tryFind (fun d ->
                        let full =
                            try
                                d.FullName
                            with _ ->
                                ""

                        String.Equals(full, AbstractDoc.path, StringComparison.OrdinalIgnoreCase))

                let doc =
                    match doc with
                    | Some d -> d
                    | None -> failwithf "这只 Word 里没打开「%s」" AbstractDoc.path

                let beforeLen = (try doc.Content.Text.Length with _ -> 0)
                let dirty = try not doc.Saved with _ -> false
                let alertsWas = app.DisplayAlerts
                app.DisplayAlerts <- Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

                try
                    // 清空正文：只动正文，页眉页脚与样式不碰
                    doc.Content.Delete() |> ignore

                    /// 段落的正文（去掉段落标记与单元格标记）
                    let norm (s: string) =
                        if isNull s then "" else s.Replace("\r", "").Replace("\a", "").Trim()

                    /// 给刚写进去的那一段套样式：从末尾往前找三段里正文对得上的那一段
                    /// （Paragraph.Style 的 setter 收 obj& ，必须按引用传）
                    let styleJustWritten (text: string) (style: obj option) =
                        match style with
                        | None -> ()
                        | Some s ->
                            let count = doc.Paragraphs.Count

                            let target =
                                [ count .. -1 .. max 1 (count - 3) ]
                                |> List.map (fun i -> doc.Paragraphs.Item(i))
                                |> List.tryFind (fun p ->
                                    try
                                        norm p.Range.Text = norm text
                                    with _ ->
                                        false)

                            match target with
                            | Some p ->
                                let mutable value = s
                                p.set_Style(&value)
                            | None -> ()

                    /// 在文末追加一段：Content.InsertAfter 一次写文字 + 段落标记
                    let writePara (text: string) (style: obj option) =
                        doc.Content.InsertAfter(text + "\r")
                        styleJustWritten text style

                    /// 表格落在文末那个空段落上；写完补一个空段落，后面的块才有地方落
                    let writeTable (header: string[]) (rows: string[][]) =
                        let last = doc.Paragraphs.Item(doc.Paragraphs.Count)

                        let tbl =
                            doc.Tables.Add(last.Range, rows.Length + 1, header.Length)

                        header
                        |> Array.iteri (fun j h -> tbl.Cell(1, j + 1).Range.Text <- h)

                        rows
                        |> Array.iteri (fun i row ->
                            row
                            |> Array.iteri (fun j v -> tbl.Cell(i + 2, j + 1).Range.Text <- v))

                        (try
                            tbl.Borders.InsideLineStyle <- Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                            tbl.Borders.OutsideLineStyle <- Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
                         with _ ->
                            ())

                        doc.Content.InsertAfter("\r")

                    for block in AbstractDoc.blocks do
                        match block with
                        | AbstractDoc.Title text ->
                            writePara
                                text
                                (Some(box Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleTitle))
                        | AbstractDoc.Head text ->
                            writePara
                                text
                                (Some(box Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1))
                        | AbstractDoc.Para text -> writePara text None
                        | AbstractDoc.Table(header, rows) -> writeTable header rows

                    doc.Save()

                    let text = doc.Content.Text

                    {| Before = beforeLen
                       Dirty = dirty
                       Paragraphs = (try doc.Paragraphs.Count with _ -> 0)
                       Tables = (try doc.Tables.Count with _ -> 0)
                       Length = text.Length
                       Text = text
                       Saved = (try doc.Saved with _ -> true) |}
                finally
                    app.DisplayAlerts <- alertsWas)

        let has (s: string) = result.Text.Contains s

        Report.line
            output
            "原文"
            (sprintf "%d 字，未保存改动 %b（这份文档的旧内容按你的要求忽略，已备份）" result.Before result.Dirty)

        Report.line output "写后" (sprintf "%d 字，%d 段，%d 张表" result.Length result.Paragraphs result.Tables)
        Report.line output "" ""

        let key =
            [ "637.45", "吸热量 637.45 万大卡/h"
              "1.6732", "α = 1.6732"
              "12751.36", "每 Nm³ 煤气放热 12751.36 kJ"
              "2093", "补燃煤气耗量 2093 Nm³/h"
              "10204", "助燃空气量 10204 Nm³/h"
              "920.93", "燃烧器额定 920.93 万大卡/h"
              "2302.3", "额定煤气耗量 2302.3 Nm³/h"
              "12000 m³/h", "风机选型风量 12000 m³/h"
              "34.36", "待确认：助燃风总管流速 34.36 m/s" ]

        for needle, what in key do
            Report.line output "核对" (sprintf "%s　%s" (if has needle then "有" else "缺") what)

        Report.line output "工作簿" (sprintf "Saved=%b" result.Saved)

        Should.equal true (has "热风炉（烟气补燃）计算书摘要")
        Should.equal true (has "637.45")
        Should.equal true (has "1.6732")
        Should.equal true (has "12751.36")
        Should.equal true (has "920.93")
        Should.equal true (result.Tables >= 2)
        Should.equal true result.Saved

        Report.conclusion output "通过：正文已整体换成计算书摘要，两张表（输入 / 输出）落盘，关键数字逐条核对"

