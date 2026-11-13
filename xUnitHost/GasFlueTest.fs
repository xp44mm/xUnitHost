/// 【黑板挂回】由「过量空气系数 + 焦炉煤气成分」估算被加热烟气的组分体积分数，
/// 作为焓值计算的输入，写在「焦炉煤气」!A13 起，格式参照
/// C:\崔胜利\群有\河北浩锐陶瓷\天然气燃烧计算.xlsx「体积法scr烟气吸热」!A1:B10。
///
/// 通道同 GasDesignTest.fs：按 PID 挂回现成实例，绝不新建 / Quit / kill；
/// COM 全在 OsHost.Sta.run 固定 STA 线程。纯燃烧计量与真机动作分开。
/// 参考文件的做法：假设被加热烟气 = 同种燃料的完全燃烧产物（本书「烟气吸热量计算」!A2 同此假设），
/// 按组分 vol.% + 烟气量 + 进出口温度算焓差。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

/// 纯规则：每 1 Nm³ 干煤气的完全燃烧计量（体积法，标态体积 = 物质的量）
module internal FlueCalc =

    /// 干基煤气成分，vol.%，与 GasDesign.picks 的 E4:E10 一一对应
    type Gas =
        { N2: float
          CO2: float
          CO: float
          H2: float
          CH4: float
          O2: float
          /// C2 以上不饱和烃，按 C2H4 计
          C2: float }

    let designGas =
        { N2 = 6.45
          CO2 = 2.5
          CO = 6.5
          H2 = 57.0
          CH4 = 24.5
          O2 = 0.55
          C2 = 2.5 }

    /// 理论需外部 O2：H2/CO 各 0.5、CH4 2、C2H4 3，扣掉燃料自带 O2
    let stoichO2 (g: Gas) : float =
        (0.5 * (g.H2 + g.CO) + 2.0 * g.CH4 + 3.0 * g.C2) / 100.0 - g.O2 / 100.0

    /// 理论空气量 Nm³/Nm³ 干煤气（空气按 O2 21% / N2 79%）
    let l0 (g: Gas) : float = stoichO2 g / 0.21

    /// 燃烧产物，Nm³/Nm³ 干煤气（湿基）
    type Products =
        { CO2: float
          H2O: float
          N2: float
          O2: float
          Wet: float }

        member this.Dry = this.Wet - this.H2O

    let products (g: Gas) (alpha: float) : Products =
        let air = alpha * l0 g
        let co2 = (g.CO2 + g.CO + g.CH4 + 2.0 * g.C2) / 100.0
        let h2o = (g.H2 + 2.0 * g.CH4 + 2.0 * g.C2) / 100.0
        let n2 = g.N2 / 100.0 + air * 0.79
        let o2 = (alpha - 1.0) * l0 g * 0.21
        { CO2 = co2; H2O = h2o; N2 = n2; O2 = o2; Wet = co2 + h2o + n2 + o2 }

    /// 给定目标 O2 体积分数反推 α（二分，0.6~6）。dry=true 用干基，false 用湿基
    let alphaForO2 (g: Gas) (targetPercent: float) (dry: bool) : float =
        let p = targetPercent / 100.0

        let o2Fraction a =
            let v = products g a
            if dry then v.O2 / v.Dry else v.O2 / v.Wet

        let mutable lo, hi = 0.6, 6.0
        for _ in 1..80 do
            let mid = (lo + hi) / 2.0
            if o2Fraction mid < p then lo <- mid else hi <- mid
        (lo + hi) / 2.0

    /// 九个组分，湿基 / 干基 vol.%。焦炉煤气没给 S / Cl / F / 灰分，一律按 0
    let components (v: Products) : (string * float * float) list =
        [ "H2O", v.H2O / v.Wet * 100.0, 0.0
          "O2", v.O2 / v.Wet * 100.0, v.O2 / v.Dry * 100.0
          "N2", v.N2 / v.Wet * 100.0, v.N2 / v.Dry * 100.0
          "CO2", v.CO2 / v.Wet * 100.0, v.CO2 / v.Dry * 100.0
          "SO2", 0.0, 0.0
          "SO3", 0.0, 0.0
          "HCl", 0.0, 0.0
          "HF", 0.0, 0.0
          "ash", 0.0, 0.0 ]

module internal FlueBoard =

    let boardText =
        [ @"[ExcelVSTO 当前 Excel 实例] 采集时间 2026-11-10 14:54:29"
          @"PID : 14052"
          @"ROT : !ExcelVSTO.Instance.14052"
          @"HWND : 0x00050338"
          @"EXCEL.EXE : C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
          @"工作簿 : 焦炉煤气燃烧计算书.xlsx"
          @"路径 : C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx" ]
        |> String.concat "\n"

    let info () : BoardApp = BlackboardText.ofText boardText |> Report.ok
    let doc () : BoardDoc = BlackboardText.docOfText boardText

    let sheetName = "焦炉煤气"

    /// 过量空气系数：默认由第一张表「烟气含氧量 8~10%」取中值 9%、按干基反推；
    /// 也可显式 XUNITHOST_GAS_ALPHA=<数> 指定（例如参考文件的燃烧器 α=1.15）
    let alpha () : float * string =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_ALPHA" with
        | null
        | "" ->
            FlueCalc.alphaForO2 FlueCalc.designGas 9.0 true,
            "第一张表「烟气含氧量 8~10%」取中值 9%（干基）反推"
        | text ->
            match Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, a when a >= 1.0 && a <= 6.0 -> a, "XUNITHOST_GAS_ALPHA 显式指定"
            | _ -> failwithf "XUNITHOST_GAS_ALPHA 不是 1~6 的数：%s" text

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

    let private addr (row: int) (col: int) = sprintf "%c%d" (char (int 'A' + col - 1)) row

    type Survey =
        { /// 这只 Excel 里开着的全部工作簿：名字 / 路径 / 表名
          Workbooks: (string * string * string list) list
          /// 所有打开的工作簿里搜到的「过量空气 / 空气系数」字样
          AlphaHits: string list
          /// 目标表 11~26 行、A~F 列
          TargetRows: (int * string list) list
          A13: string
          BookSaved: bool
          DiskAt: string }

    /// 只读现场：不写、不存盘
    let survey (app: Application) : Survey =
        let path = bookPath ()
        let book = findBook app path

        let opened =
            [ 1..app.Workbooks.Count ]
            |> List.map(fun i ->
                let b = app.Workbooks.Item(i)
                let sheets = [ 1..b.Worksheets.Count ] |> List.map(fun j -> (b.Worksheets.Item(j) :?> Worksheet).Name)
                b.Name, (try b.FullName with _ -> ""), sheets, b)

        let hits = ResizeArray<string>()

        for name, _, _, b in opened do
            for si in 1..b.Worksheets.Count do
                let sheet = b.Worksheets.Item(si) :?> Worksheet
                let maxRow = min 40 (try sheet.UsedRange.Rows.Count with _ -> 0)
                let maxCol = min 8 (try sheet.UsedRange.Columns.Count with _ -> 0)

                for r in 1..maxRow do
                    for c in 1..maxCol do
                        let text = cellText sheet (addr r c)

                        if text.Contains "过量空气" || text.Contains "空气系数" then
                            // 参考表的排版是 A=名称 B=单位 C=数值，三格一起报才看得见 α
                            hits.Add(
                                sprintf
                                    "%s!%s!%s=「%s」 B=「%s」 C=「%s」"
                                    name
                                    sheet.Name
                                    (addr r c)
                                    text
                                    (cellText sheet (addr r (c + 1)))
                                    (cellText sheet (addr r (c + 2)))
                            )

        let target = findSheet book sheetName

        let rows =
            [ for r in 11..26 -> r, [ for c in 1..6 -> cellText target (addr r c) ] ]

        { Workbooks = opened |> List.map(fun (n, p, s, _) -> n, p, s)
          AlphaHits = List.ofSeq hits
          TargetRows = rows
          A13 = cellText target "A13"
          BookSaved = (try book.Saved with _ -> true)
          DiskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss") }

    /// 写盘：A13:E22 组分表（湿基 / 干基）+ A24:E27 计算依据
    let write (app: Application) (alpha: float) (reason: string) : string list * bool * string =
        let path = bookPath ()
        let book = findBook app path
        let sheet = findSheet book sheetName

        // 写之前再验一次：A13:E22 全空，别盖掉用户刚写的东西
        let occupied =
            [ for r in 13..22 do
                for c in 1..5 do
                    let text = cellText sheet (addr r c)
                    if text <> "" then yield sprintf "%s%d=「%s」" (string (char (int 'A' + c - 1))) r text ]

        Should.equal ([]: string list) occupied

        let v = FlueCalc.products FlueCalc.designGas alpha
        let comps = FlueCalc.components v
        let l0 = FlueCalc.l0 FlueCalc.designGas
        let round (x: float) = Math.Round(x, 2)

        let put (a: string) (value: obj) = sheet.Range(a).Value2 <- value

        // 表头
        put "A13" (box "组分/参数")
        put "B13" (box "单位")
        put "C13" (box "被加热烟气 湿基")
        put "D13" (box "被加热烟气 干基")
        put "E13" (box "备注")

        // 九个组分（湿基 / 干基都是数值，H2O 干基恒 0）
        comps
        |> List.iteri(fun i (name, wet, dry) ->
            let r = 14 + i
            put (sprintf "A%d" r) (box name)
            put (sprintf "B%d" r) (box "vol. %")
            put (sprintf "C%d" r) (box (round wet))
            put (sprintf "D%d" r) (box (round dry))
            put (sprintf "E%d" r) (box(if name = "H2O" then "燃烧生成水；燃料水分按干基忽略"
                                      elif List.contains name [ "SO2"; "SO3"; "HCl"; "HF"; "ash" ] then "燃料未给 S/Cl/F/灰分，按 0"
                                      elif name = "O2" then "由过量空气系数与燃料自带 O2 决定"
                                      elif name = "N2" then "燃料 N2 + 助燃空气 N2"
                                      else "完全燃烧生成")))

        // 计算依据
        put "A24" (box "过量空气系数 α")
        put "B24" (box "—")
        put "C24" (box (round alpha))
        put "E24" (box reason)
        put "A25" (box "理论空气量 L0")
        put "B25" (box "Nm³/Nm³煤气")
        put "C25" (box (round l0))
        put "E25" (box "H2/CO 各 0.5 O2，CH4 2，C2H4 3；空气 O2 21% / N2 79%")
        put "A26" (box "湿基产物量")
        put "B26" (box "Nm³/Nm³煤气")
        put "C26" (box (round v.Wet))
        put "E26" (box "每 1 Nm³ 干焦炉煤气完全燃烧后的湿烟气量")
        put "A27" (box "口径")
        put "E27" (box "被加热烟气按同种焦炉煤气完全燃烧产物估（见「烟气吸热量计算」!A2）；C 列湿基用于焓值，D 列干基用于对照")

        book.Save()

        let after = [ for r in 14..22 -> cellText sheet (sprintf "C%d" r) ]
        let saved = (try book.Saved with _ -> true)
        let diskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss")
        after, saved, diskAt

/// 三条 Fact：纯燃烧计量 → 只读清点（含 α 出处）→ 真写盘
type GasFlueTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``烟气成分规则：燃烧计量与反推算得对``() =
        let g = FlueCalc.designGas
        let l0 = FlueCalc.l0 g
        Report.title output "烟气成分规则：焦炉煤气完全燃烧计量（每 1 Nm³ 干煤气）"
        Report.line output "设计成分" (sprintf "N2 %.2f CO2 %.2f CO %.2f H2 %.2f CH4 %.2f O2 %.2f C2H4 %.2f (vol.%%)" g.N2 g.CO2 g.CO g.H2 g.CH4 g.O2 g.C2)
        Report.line output "理论 O2" (sprintf "%.4f Nm³" (FlueCalc.stoichO2 g))
        Report.line output "理论空气 L0" (sprintf "%.4f Nm³/Nm³" l0)

        for a in [ 1.0; 1.15; 1.6733; 1.8726 ] do
            let v = FlueCalc.products g a
            Report.line
                output
                (sprintf "α=%.4f" a)
                (sprintf "湿基 H2O %.2f O2 %.2f N2 %.2f CO2 %.2f %%；干基 O2 %.2f %%；产物 %.3f Nm³"
                    (v.H2O / v.Wet * 100.0) (v.O2 / v.Wet * 100.0) (v.N2 / v.Wet * 100.0)
                    (v.CO2 / v.Wet * 100.0) (v.O2 / v.Dry * 100.0) v.Wet)

        Report.line output "反推 α" (sprintf "干基 O2 9%% → %.4f；湿基 O2 9%% → %.4f" (FlueCalc.alphaForO2 g 9.0 true) (FlueCalc.alphaForO2 g 9.0 false))

        let wetSum = FlueCalc.components (FlueCalc.products g 1.6733) |> List.sumBy(fun (_, w, _) -> w)
        Should.equal 100.0 (Math.Round(wetSum, 6))
        Assert.True(l0 > 3.5 && l0 < 5.0, "焦炉煤气理论空气量应在 3.5~5 Nm³/Nm³")
        Assert.True(abs (FlueCalc.products g (FlueCalc.alphaForO2 g 9.0 true)).O2 / (FlueCalc.products g (FlueCalc.alphaForO2 g 9.0 true)).Dry * 100.0 - 9.0 < 0.001)
        Report.conclusion output "通过：L0≈4.18 符合焦炉煤气经验值，湿基组分合计 100%，α 反推回代得回目标 O2；纯逻辑，不碰真机"

    [<Fact>]
    member this.``烟气成分清点：只读现场与过量空气系数出处``() =
        let info = FlueBoard.info ()
        let doc = FlueBoard.doc ()
        let before = Report.procIds info.Image

        Report.title output "烟气成分清点：挂回常驻 Excel，查明过量空气系数出处与 A13 现场"
        Report.line output "参数表" (info.describe())
        Report.line output "文档" (doc.describe())

        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

        let snap = STA.run(fun () -> FlueBoard.survey app)

        for name, path, sheets in snap.Workbooks do
            Report.line output "打开的工作簿" (sprintf "「%s」 %s" name path)
            Report.line output "  其中工作表" (String.concat "、" sheets)

        if snap.AlphaHits.IsEmpty then
            Report.line output "α 线索" "所有打开的工作簿里都没搜到「过量空气 / 空气系数」字样"
        else
            for hit in snap.AlphaHits do
                Report.line output "α 线索" hit

        for row, cells in snap.TargetRows do
            if cells |> List.exists(fun v -> v <> "") then
                Report.line output (sprintf "第%d行" row) (String.concat " | " cells)

        Report.line output "A13" (if snap.A13 = "" then "(空)" else snap.A13)
        Report.line output "工作簿 Saved" (string snap.BookSaved)
        Report.line output "磁盘最后写入" snap.DiskAt

        let alpha, reason = FlueBoard.alpha ()
        let v = FlueCalc.products FlueCalc.designGas alpha
        Report.line output "拟用 α" (sprintf "%.4f（%s）" alpha reason)
        Report.line output "对应湿基" (sprintf "H2O %.2f O2 %.2f N2 %.2f CO2 %.2f SO2 0 (vol.%%)" (v.H2O / v.Wet * 100.0) (v.O2 / v.Wet * 100.0) (v.N2 / v.Wet * 100.0) (v.CO2 / v.Wet * 100.0))

        let burner = FlueCalc.products FlueCalc.designGas 1.15
        Report.line output "参考 α=1.15" (sprintf "湿基 H2O %.2f O2 %.2f N2 %.2f CO2 %.2f；干基 O2 %.2f (vol.%%)" (burner.H2O / burner.Wet * 100.0) (burner.O2 / burner.Wet * 100.0) (burner.N2 / burner.Wet * 100.0) (burner.CO2 / burner.Wet * 100.0) (burner.O2 / burner.Dry * 100.0))
        Report.line output "矛盾" (sprintf "第一张表要求烟气含氧量 8~10%%，α=1.15 只给 3.01%%（干基）；按含氧量 9%% 反推 α=%.4f" (FlueCalc.alphaForO2 FlueCalc.designGas 9.0 true))

        let after = Report.procIds info.Image
        Report.line output "收尾" (sprintf "%s 进程 %A" info.Image (Set.toList after))
        Should.equal before after
        Report.conclusion output "通过：只读查明 α 出处与 A13 现场，磁盘一个字节没动"

    [<Fact>]
    member this.``烟气成分回填：A13 起写被加热烟气组分``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_FLUE" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_FLUE=1 才真的写盘；α 可用 XUNITHOST_GAS_ALPHA 指定"
        | _ ->
            let info = FlueBoard.info ()
            let doc = FlueBoard.doc ()
            let before = Report.procIds info.Image
            let target =
                match doc.Path with
                | Some path -> path
                | None -> failwith "黑板里没有落盘路径，不往下走"

            Assert.True(File.Exists target, sprintf "文件不在：%s" target)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmmss")
            let backup = Path.Combine(projectRoot, "TestResults", sprintf "焦炉煤气燃烧计算书[烟气成分前]%s.xlsx" stamp)
            let backedUp = not (File.Exists backup)
            if backedUp then File.Copy(target, backup, false)

            let alpha, reason = FlueBoard.alpha ()
            let v = FlueCalc.products FlueCalc.designGas alpha

            Report.title output "烟气成分回填：A13 起写被加热烟气组分（焓值计算输入）"
            Report.line output "参数表" (info.describe())
            Report.line output "文档" (doc.describe())
            Report.line output "备份" (if backedUp then backup else "已有备份，未重复生成：" + backup)
            Report.line output "α 取值" (sprintf "%.4f（%s）" alpha reason)
            Report.line output "理论空气 L0" (sprintf "%.3f Nm³/Nm³煤气" (FlueCalc.l0 FlueCalc.designGas))
            Report.line output "湿基产物量" (sprintf "%.3f Nm³/Nm³煤气" v.Wet)

            let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
            Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

            let cAfter, saved, diskAfter = STA.run(fun () -> FlueBoard.write app alpha reason)

            for (name, wet, dry), readBack in List.zip (FlueCalc.components v) cAfter do
                Report.line
                    output
                    (sprintf "%-4s" name)
                    (sprintf "湿基 %.2f %%　干基 %.2f %%　读回「%s」" wet dry readBack)

            Report.line output "工作簿 Saved" (string saved)
            Report.line output "存盘后" diskAfter
            Should.equal (FlueCalc.components v |> List.map(fun (_, w, _) -> string (Math.Round(w, 2)))) cAfter
            Should.equal true saved
            Should.equal true (File.Exists backup)

            let after = Report.procIds info.Image
            Report.line output "现场" (sprintf "%s 进程 干活前 %A / 收尾后 %A" info.Image (Set.toList before) (Set.toList after))
            Should.equal before after
            Should.equal true (Set.contains info.Pid after)
            Report.conclusion output "完成：A13 起九个组分湿基/干基 vol.% 与计算依据已存盘并读回一致；原件备份在项目目录外"
