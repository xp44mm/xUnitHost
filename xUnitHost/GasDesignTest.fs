/// 【黑板挂回】焦炉煤气燃烧计算书.xlsx /「焦炉煤气」!E2:E11「设计值」：
/// 把 D 列的业主区间收成一组可用于后续燃烧计算、且对设备选型稍有裕量的设计参数。
///
/// 通道与 GasParamTest.fs 相同：ExcelVSTO 抄现场 → 按 PID 挂回现成实例，
/// 绝不新建 / 绝不 Quit / 绝不 kill；COM 读写全在 OsHost.Sta.run 固定 STA 线程。
/// 纯规则（区间取值 / 成分自洽 / 热值核算）与真机动作分开，先跑得过纯规则再碰真机。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

/// 纯规则：区间解析、端点 / 中值、成分自洽与低位热值核算
module internal DesignRule =

    /// "4~8" → (4.0, 8.0)；"少量" / "--" 之类返回 None
    let range (raw: string) : (float * float) option =
        let parts = raw.Split '~'

        if parts.Length <> 2 then
            None
        else
            let okLo, lo = Double.TryParse(parts.[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture)
            let okHi, hi = Double.TryParse(parts.[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture)
            if okLo && okHi && lo <= hi then Some(lo, hi) else None

    let low (raw: string) = range raw |> Option.map fst
    let high (raw: string) = range raw |> Option.map snd
    let mid (raw: string) = range raw |> Option.map(fun (lo, hi) -> (lo + hi) / 2.0)

    /// 干基低位热值常数（kcal/Nm³，工程近似）：只用于自洽性核算，不写进 E 列
    let lhv =
        [ "H2", 2575.0
          "CH4", 8550.0
          "CO", 3020.0
          "C2以上不饱和烃", 14000.0 ]

    /// 成分体积分数（%）→ 估算低位热值
    let heatValue (parts: (string * float) list) : float =
        parts
        |> List.sumBy(fun (name, percent) ->
            match lhv |> List.tryPick(fun (key, value) -> if key = name then Some value else None) with
            | Some value -> value * percent / 100.0
            | None -> 0.0)

module internal GasDesign =

    /// 插件 2026-11-10 14:34:57 抄下来的现场，原文留在代码里当记录
    let boardText =
        [ @"[ExcelVSTO 当前 Excel 实例] 采集时间 2026-11-10 14:34:57"
          @"PID : 28884"
          @"ROT : !ExcelVSTO.Instance.28884"
          @"HWND : 0x000A04D6"
          @"EXCEL.EXE : C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
          @"工作簿 : 焦炉煤气燃烧计算书.xlsx"
          @"路径 : C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx" ]
        |> String.concat "\n"

    let info () : BoardApp = BlackboardText.ofText boardText |> Report.ok
    let doc () : BoardDoc = BlackboardText.docOfText boardText

    let sheetName = "焦炉煤气"

    /// 磁盘 / 真机上点名这十行的原始区间（写之前逐行核对，现场变了就停手）
    let rawData =
        [ "4~8"; "4000~4400"; "4~7"; "2~3"; "5~8"; "54~60"; "24~25"; "0.4~0.7"; "2~3"; "少量" ]

    type Pick =
        { Row: int
          Item: string
          Unit: string
          Raw: string
          /// 数字设计值；None = 这一格填文字
          Value: float option
          Text: string
          Why: string }

    /// 设计值口径：压力 / 热值取下限（设备按最不利工况选型，自带裕量）；
    /// 成分取各区间中值，N2 作平衡补到 100%（6.45% 仍在 4~7 内）；水分按干基忽略。
    /// 成分核算低位热值约 4100 kcal/Nm³，设计热值取 4000 → 煤气耗量算大，设备留约 2.7% 裕量。
    let picks: Pick list =
        [ { Row = 2; Item = "压力"; Unit = "kpa"; Raw = "4~8"; Value = Some 4.0; Text = "4"
            Why = "取下限：按最低供气压力选燃烧器 / 管径，实际压力更高即富余" }
          { Row = 3; Item = "热值"; Unit = "kcal"; Raw = "4000~4400"; Value = Some 4000.0; Text = "4000"
            Why = "取下限：按低热值算煤气耗量；成分核算约 4109，留约 2.7%" }
          { Row = 4; Item = "N2"; Unit = "%"; Raw = "4~7"; Value = Some 6.45; Text = "6.45"
            Why = "中值 5.5 再作平衡项补到 100%，仍在 4~7 内" }
          { Row = 5; Item = "CO2"; Unit = "%"; Raw = "2~3"; Value = Some 2.5; Text = "2.5"
            Why = "取中值" }
          { Row = 6; Item = "CO"; Unit = "%"; Raw = "5~8"; Value = Some 6.5; Text = "6.5"
            Why = "取中值" }
          { Row = 7; Item = "H2"; Unit = "%"; Raw = "54~60"; Value = Some 57.0; Text = "57"
            Why = "取中值" }
          { Row = 8; Item = "CH4"; Unit = "%"; Raw = "24~25"; Value = Some 24.5; Text = "24.5"
            Why = "取中值" }
          { Row = 9; Item = "O2"; Unit = "%"; Raw = "0.4~0.7"; Value = Some 0.55; Text = "0.55"
            Why = "取中值" }
          { Row = 10; Item = "C2以上不饱和烃"; Unit = "%"; Raw = "2~3"; Value = Some 2.5; Text = "2.5"
            Why = "取中值" }
          { Row = 11; Item = "水分"; Unit = ""; Raw = "少量"; Value = None; Text = "忽略（按干基）"
            Why = "D 列只说少量；干基成分已合计 100%，设计按干基计" } ]

    /// 干基成分合计（不含 N2 的平衡修正前）：用来验算 N2 该补多少
    let componentMids =
        [ "CO2", 2.5; "CO", 6.5; "H2", 57.0; "CH4", 24.5; "O2", 0.55; "C2以上不饱和烃", 2.5 ]

    let nitrogen = 100.0 - (componentMids |> List.sumBy snd)

    let private bookPath () : string =
        match (doc ()).Path with
        | Some path -> path
        | None -> failwith "黑板里没有落盘路径，不往下走"

    let private findBook (app: Application) (path: string) : Workbook =
        let count = try app.Workbooks.Count with _ -> 0

        let found =
            [ 1..count ]
            |> List.tryPick(fun i ->
                let book = app.Workbooks.Item(i)
                let full = try book.FullName with _ -> ""

                if String.Equals(full, path, StringComparison.OrdinalIgnoreCase) then
                    Some book
                else
                    None)

        match found with
        | Some book -> book
        | None -> failwithf "这只 Excel 里没开着「%s」" path

    let private findSheet (book: Workbook) (name: string) : Worksheet =
        let found =
            [ 1..book.Worksheets.Count ]
            |> List.tryPick(fun i ->
                let sheet = book.Worksheets.Item(i) :?> Worksheet
                if sheet.Name = name then Some sheet else None)

        match found with
        | Some sheet -> sheet
        | None -> failwithf "工作簿里没有表「%s」" name

    let cellText (sheet: Worksheet) (addr: string) : string =
        let value = sheet.Range(addr).Value2
        if isNull value then "" else Convert.ToString(value, CultureInfo.InvariantCulture)

    let private merged (sheet: Worksheet) (addr: string) : bool =
        try
            let value = sheet.Range(addr).MergeCells
            not (isNull value) && Convert.ToBoolean(value, CultureInfo.InvariantCulture)
        with _ ->
            false

    type Survey =
        { BookName: string
          BookPath: string
          Sheets: (string * string) list
          Header: string list
          /// 目标表 A2:F11：行号 + 六列文本
          Rows: (int * string list) list
          /// 工况参数表 D 列（2..12）
          Working: string list
          /// 吸热量计算表现有文字
          Calc: string list
          EColumn: string list
          MergedE: bool
          BookSaved: bool
          DiskAt: string }

    /// 只读现场：目标表整块 + 另外两张表的依据，一个字不写
    let survey (app: Application) : Survey =
        let path = bookPath ()
        let book = findBook app path
        let sheet = findSheet book sheetName

        let sheets =
            [ 1..book.Worksheets.Count ]
            |> List.map(fun i ->
                let s = book.Worksheets.Item(i) :?> Worksheet
                s.Name, (try s.UsedRange.Address() with _ -> "?"))

        let rows =
            [ for r in 2..11 ->
                r, [ for c in [ 'A' .. 'F' ] -> cellText sheet (sprintf "%c%d" c r) ] ]

        let working =
            let w = findSheet book "需要加热的烟气参数"
            [ for r in 2..12 -> cellText w (sprintf "D%d" r) ]

        let calc =
            let c = findSheet book "烟气吸热量计算"
            [ for r in 1..5 -> cellText c (sprintf "A%d" r) ]

        { BookName = book.Name
          BookPath = book.FullName
          Sheets = sheets
          Header = [ for c in [ 'A' .. 'F' ] -> cellText sheet (sprintf "%c1" c) ]
          Rows = rows
          Working = working
          Calc = calc
          EColumn = [ for r in 2..11 -> cellText sheet (sprintf "E%d" r) ]
          MergedE = merged sheet "E2"
          BookSaved = (try book.Saved with _ -> true)
          DiskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss") }

    /// 写盘：现场复核 → 十格逐格填 → Save() → 同一次 STA 调用里逐格读回
    let write (app: Application) : string list * bool * string =
        let path = bookPath ()
        let book = findBook app path
        let sheet = findSheet book sheetName

        // 写之前再验一次：D 列原文没变、E 列全空、E2 不是合并格
        Should.equal rawData [ for r in 2..11 -> cellText sheet (sprintf "D%d" r) ]
        Should.equal (List.replicate 10 "") [ for r in 2..11 -> cellText sheet (sprintf "E%d" r) ]
        Assert.True(not (merged sheet "E2"), "E2 是合并单元格，填值前要先人工拆并")

        for pick in picks do
            let boxed = match pick.Value with Some number -> box number | None -> box pick.Text
            sheet.Range(sprintf "E%d" pick.Row).Value2 <- boxed

        book.Save()

        let after = [ for r in 2..11 -> cellText sheet (sprintf "E%d" r) ]
        let saved = (try book.Saved with _ -> true)
        let diskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss")
        after, saved, diskAt

/// 三条 Fact：纯规则 → 只读清点 → 真写盘（写盘那条要显式设环境变量）
type GasDesignTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``设计取值规则：区间与成分自洽算得对``() =
        Report.title output "设计取值规则：区间解析、成分归一、热值自洽"
        Report.line output "区间" (sprintf "压力 %A 热值 %A N2 %A" (DesignRule.range "4~8") (DesignRule.range "4000~4400") (DesignRule.range "4~7"))
        Report.line output "文字" (sprintf "少量 → %A；-- → %A" (DesignRule.range "少量") (DesignRule.range "--"))

        let lows = GasDesign.rawData |> List.map DesignRule.low
        let highs = GasDesign.rawData |> List.map DesignRule.high
        // 合计只算七项干基成分（前两行是压力 / 热值，单位不同，不能加进来）
        Report.line output "成分全下限" (sprintf "%.1f%%" (lows |> List.skip 2 |> List.choose id |> List.sum))
        Report.line output "成分全上限" (sprintf "%.1f%%" (highs |> List.skip 2 |> List.choose id |> List.sum))
        Report.line output "N2 平衡" (sprintf "%.2f%%（业主区间 4~7）" GasDesign.nitrogen)
        Report.line output "成分合计" (sprintf "%.2f%%" (GasDesign.nitrogen + (GasDesign.componentMids |> List.sumBy snd)))
        Report.line output "核算热值" (sprintf "%.1f kcal/Nm³（业主区间 4000~4400）" (DesignRule.heatValue GasDesign.componentMids))
        Report.line output "设计热值" (sprintf "%.0f kcal/Nm³ → 煤气耗量裕量 %.1f%%" 4000.0 ((DesignRule.heatValue GasDesign.componentMids / 4000.0 - 1.0) * 100.0))

        Should.equal (Some(4.0, 8.0)) (DesignRule.range "4~8")
        Should.equal None (DesignRule.range "少量")
        Should.equal 100.0 (GasDesign.nitrogen + (GasDesign.componentMids |> List.sumBy snd))
        Assert.True(GasDesign.nitrogen >= 4.0 && GasDesign.nitrogen <= 7.0)
        Assert.True(DesignRule.heatValue GasDesign.componentMids >= 4000.0)
        Report.conclusion output "通过：端点组合确实不自洽，取中值 + N2 平衡才 100%；纯逻辑，不碰真机"

    [<Fact>]
    member this.``设计参数清点：只读全表现场``() =
        let info = GasDesign.info ()
        let doc = GasDesign.doc ()
        let before = Report.procIds info.Image

        Report.title output "设计参数清点：挂回常驻 Excel，只读「焦炉煤气」全表与依据"
        Report.line output "参数表" (info.describe())
        Report.line output "文档" (doc.describe())

        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

        let snap = STA.run(fun () -> GasDesign.survey app)
        Report.line output "工作簿" (sprintf "「%s」 %s" snap.BookName snap.BookPath)

        for name, used in snap.Sheets do
            Report.line output "工作表" (sprintf "「%s」 UsedRange=%s" name used)

        Report.line output "表头" (String.concat " | " snap.Header)

        for row, cells in snap.Rows do
            Report.line output (sprintf "第%d行" row) (String.concat " | " cells)

        Report.line output "工况 D2:D12" (String.concat " | " snap.Working)
        Report.line output "吸热量计算" (String.concat " | " snap.Calc)
        Report.line output "E 列" (sprintf "全空=%b；内容=%s" (snap.EColumn |> List.forall(fun v -> v = "")) (String.concat "|" snap.EColumn))
        Report.line output "E2 合并格" (string snap.MergedE)
        Report.line output "工作簿 Saved" (string snap.BookSaved)
        Report.line output "磁盘最后写入" snap.DiskAt

        Should.equal GasDesign.rawData [ for _, cells in snap.Rows -> cells.[3] ]
        Should.equal (List.replicate 10 "") snap.EColumn

        let after = Report.procIds info.Image
        Report.line output "收尾" (sprintf "%s 进程 %A" info.Image (Set.toList after))
        Should.equal before after
        Report.conclusion output "通过：D 列十行与业主区间一致，E 列全空；只读，磁盘一个字节没动"

    [<Fact>]
    member this.``设计参数回填：E2:E11 定值``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_DESIGN" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_DESIGN=1 才真的写盘"
        | _ ->
            let info = GasDesign.info ()
            let doc = GasDesign.doc ()
            let before = Report.procIds info.Image
            let target =
                match doc.Path with
                | Some path -> path
                | None -> failwith "黑板里没有落盘路径，不往下走"

            Assert.True(File.Exists target, sprintf "文件不在：%s" target)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmmss")
            let backup = Path.Combine(projectRoot, "TestResults", sprintf "焦炉煤气燃烧计算书[定值前]%s.xlsx" stamp)
            let backedUp = not (File.Exists backup)
            if backedUp then File.Copy(target, backup, false)

            Report.title output "设计参数回填：按「选设备稍有裕量」定 E2:E11 设计值"
            Report.line output "参数表" (info.describe())
            Report.line output "文档" (doc.describe())
            Report.line output "备份" (if backedUp then backup else "已有备份，未重复生成：" + backup)
            Report.line output "备份字节" (sprintf "%d B" (FileInfo backup).Length)

            let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
            Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

            let eBefore, diskBefore =
                STA.run(fun () ->
                    let snap = GasDesign.survey app
                    snap.EColumn, snap.DiskAt)

            Report.line output "存盘前" diskBefore
            Should.equal (List.replicate 10 "") eBefore

            let eAfter, saved, diskAfter = STA.run(fun () -> GasDesign.write app)

            for pick in GasDesign.picks do
                let i = pick.Row - 2
                Report.line2
                    output
                    (sprintf "E%d" pick.Row)
                    (sprintf "%s（%s）" pick.Item pick.Unit)
                    (sprintf "D「%s」→「%s」　%s" pick.Raw eAfter.[i] pick.Why)

            Report.line output "工作簿 Saved" (string saved)
            Report.line output "存盘后" diskAfter
            Should.equal (GasDesign.picks |> List.map(fun p -> p.Text)) eAfter
            Should.equal true saved
            Should.equal true (File.Exists backup)

            let after = Report.procIds info.Image
            Report.line output "现场" (sprintf "%s 进程 干活前 %A / 收尾后 %A" info.Image (Set.toList before) (Set.toList after))
            Should.equal before after
            Should.equal true (Set.contains info.Pid after)
            Report.conclusion output "完成：十格设计值已存盘并逐格读回一致；原件备份在项目目录外，实例没新建也没退出"
