/// 【补热燃烧计算】参考 C:\崔胜利\群有\煤气燃烧计算.xlsx 的四段式
/// （大气与烟气参数 → 煤气 → 过剩空气 → 烟气分析），跳过污染物单元格
/// （SO2 / SO3 / HCl / HF / ash：焦炉煤气未给 S/Cl/F/灰分，按 0）。
///
/// 目标 C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx 只有「被加热烟气吸热侧」，
/// 缺燃烧侧：燃料/空气/燃烧产物量、产物成分、以及「前后焓值变化」算出的放热量。
/// 结果另存 C:\崔胜利\群有\中冶南方都市环保\补热焦炉煤气燃烧计算书.xlsx，原书一个字节不动。
///
/// 大气数据按湖北仙桃（2026-11-10 查官方资料）：
///   仙桃市政府《汉江生态经济带水利发展规划》：多年平均气温 16.5℃、年平均相对湿度 80%
///   仙桃市气象局：年平均气温 17.0℃、年降水量 1238.6 mm
///   国网湖北电力《仙桃仙西 220kV 输变电工程水保报告》：多年平均气压 1006.5 hPa、相对湿度 79%
///   仙桃市林业十三五规划：平均海拔 26 m
/// 取 100760 Pa（沿用参考表，与仙桃气象站 1006.5 hPa 一致）、16.5℃、80%，都做成可改输入格。
///
/// 通道同前面几轮：按 PID 挂回现成实例，绝不新建 / Quit / kill；
/// COM 全在 OsHost.Sta.run 固定 STA 线程。写公式一律 Range.Formula2（本机实测：
/// Range.Formula 会给自定义函数补 @ 隐式交集，XLL 数组函数会 spill 不出来）。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal CombustionBoard =

    /// 本轮两段黑板原文：同一只 Excel（pid 28884）里先后活动的是参考表与目标书
    let refBlackboard =
        [ @"[ExcelVSTO 当前 Excel 实例] 采集时间 2026-11-10 15:53:11"
          @"PID : 28884"
          @"ROT : !ExcelVSTO.Instance.28884"
          @"HWND : 0x001B06B2"
          @"EXCEL.EXE : C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
          @"工作簿 : 煤气燃烧计算.xlsx"
          @"路径 : C:\崔胜利\群有\煤气燃烧计算.xlsx" ]
        |> String.concat "\n"

    let targetBlackboard =
        [ @"[ExcelVSTO 当前 Excel 实例] 采集时间 2026-11-10 17:26:56"
          @"PID : 14052"
          @"ROT : !ExcelVSTO.Instance.14052"
          @"HWND : 0x00050338"
          @"EXCEL.EXE : C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
          @"工作簿 : 焦炉煤气燃烧计算书.xlsx"
          @"路径 : C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx" ]
        |> String.concat "\n"

    /// 挂回用 identity：两段黑板 PID / ROT / EXE 相同，取目标书那份
    let info () : BoardApp = BlackboardText.ofText targetBlackboard |> Report.ok

    let refPath = @"C:\崔胜利\群有\煤气燃烧计算.xlsx"
    let targetPath = @"C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx"
    let outPath = @"C:\崔胜利\群有\中冶南方都市环保\补热焦炉煤气燃烧计算书.xlsx"

    /// 湖北仙桃大气设计输入（见文件头出处），单位与参考表一致
    let pressurePa = 100760.0
    let airCelsius = 16.5
    let humidityPct = 80.0

    /// 参考表用的空气成分：干空气 O2 21% / N2 79%，水蒸气按 φ·p_sat(t)/P
    let dryO2 = 21.0
    let dryN2 = 79.0

module internal CombustionSurvey =

    let private txt (v: obj) =
        if isNull v then "" else Convert.ToString(v, CultureInfo.InvariantCulture)

    let private describe (v: obj) =
        match v with
        | null -> "(null)"
        | :? string as s -> s
        | :? double as d -> d.ToString("G15", CultureInfo.InvariantCulture)
        | :? int as i when i < -2000000000 -> sprintf "错误码 %d" i
        | :? (obj[,]) as a ->
            let i, j = a.GetLowerBound(0), a.GetLowerBound(1)
            sprintf "数组 %d×%d 首元素「%s」" (a.GetLength 0) (a.GetLength 1) (txt (a.[i, j]))
        | _ -> txt v

    let private findBook (app: Application) (path: string) : Workbook option =
        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
        |> List.tryPick(fun i ->
            let b = app.Workbooks.Item(i)
            let full = try b.FullName with _ -> ""

            if String.Equals(full, path, StringComparison.OrdinalIgnoreCase) then Some b else None)

    let private findSheet (book: Workbook) (name: string) : Worksheet =
        [ 1 .. book.Worksheets.Count ]
        |> List.tryPick(fun i ->
            let s = book.Worksheets.Item(i) :?> Worksheet
            if s.Name = name then Some s else None)
        |> function
            | Some s -> s
            | None -> failwithf "工作簿里没有表「%s」" name

    let private cell (sheet: Worksheet) (addr: string) : string * string * string =
        let r = sheet.Range(addr)
        let v = txt (r.Value2)
        let f = txt (r.Formula)
        let f2 = txt (r.Formula2)
        v, f, f2

    type Snapshot =
        { Opened: string list
          RefAir: (string * string) list
          RefGas: (string * string) list
          RefAlpha: string
          RefProducts: (string * string) list
          RefFormulas: string list
          TargetGas: (string * string) list
          TargetHeat: (string * string) list
          Probes: string list
          OutExists: bool
          OutAt: string }

    let survey (app: Application) : Snapshot =
        let opened =
            [ for i in 1 .. (try app.Workbooks.Count with _ -> 0) do
                  let b = app.Workbooks.Item(i)
                  yield sprintf "「%s」 %s Saved=%b" (try b.Name with _ -> "?") (try b.FullName with _ -> "?") (try b.Saved with _ -> true) ]

        let refBook = findBook app CombustionBoard.refPath
        let targetBook = findBook app CombustionBoard.targetPath

        let refAir, refGas, refAlpha, refProducts, refFormulas =
            match refBook with
            | None -> [], [], "(参考工作簿没开)", [], []
            | Some b ->
                let air = findSheet b "大气与烟气参数"
                let gas = findSheet b "煤气"
                let excess = findSheet b "过剩空气"

                let take (sheet: Worksheet) (addrs: string list) =
                    addrs |> List.map(fun a -> a, let v, _, _ = cell sheet a in v)

                let airRows = take air [ "C3"; "C4"; "C5"; "C7"; "C8"; "C9" ]
                let gasRows = take gas [ "C8"; "C9"; "C10"; "C11"; "C12"; "C13"; "C14"; "C15"; "C16"; "C19"; "C25" ]
                let alpha = let v, _, _ = cell excess "C15" in v

                let products =
                    take excess [ "C17"; "C18"; "C19"; "C21"; "C22"; "C23"; "C24"; "C26"; "C27"; "C31"; "C32"; "C33" ]

                let formulas =
                    [ "大气与烟气参数!C7", air
                      "煤气!C19", gas
                      "煤气!C24", gas
                      "过剩空气!C17", excess
                      "过剩空气!C21", excess
                      "过剩空气!C31", excess ]
                    |> List.map(fun (label, sheet) ->
                        let addr = label.Substring(label.IndexOf('!') + 1)
                        let _, f, f2 = cell sheet addr
                        sprintf "%s → Formula「%s」 Formula2「%s」" label f f2)

                airRows, gasRows, alpha, products, formulas

        let targetGas, targetHeat =
            match targetBook with
            | None -> [], []
            | Some b ->
                let g = findSheet b "焦炉煤气"
                let p = findSheet b "需要加热的烟气参数"
                let h = findSheet b "烟气吸热量计算"

                let take (sheet: Worksheet) (addrs: string list) =
                    addrs |> List.map(fun a -> a, let v, _, _ = cell sheet a in v)

                (take g [ "E2"; "E3"; "E4"; "E5"; "E6"; "E7"; "E8"; "E9"; "E10"; "E11" ]
                 @ take p [ "D2"; "D3"; "D4"; "D5"; "D9" ]),
                take h [ "C14"; "C15"; "C33"; "C36"; "C37"; "C38" ]

        // 探测 XLL：标量函数直接 Evaluate，数组函数用 INDEX 取首元素
        let probes =
            [ sprintf "gasMoleEnthalpy(...) = %s" (describe (app.Evaluate(box "gasMoleEnthalpy({\"N2\"},{1},500)")))
              sprintf "molar(\"CO2\") = %s" (describe (app.Evaluate(box "molar(\"CO2\")")))
              sprintf "空气体积成分(仙桃) 首元素 = %s"
                  (describe (app.Evaluate(box (sprintf "INDEX(空气体积成分(%s,%s,%s),1)"
                                                    (CombustionBoard.pressurePa.ToString("G", CultureInfo.InvariantCulture))
                                                    (CombustionBoard.airCelsius.ToString("G", CultureInfo.InvariantCulture))
                                                    (CombustionBoard.humidityPct.ToString("G", CultureInfo.InvariantCulture)))))) ]

        { Opened = opened
          RefAir = refAir
          RefGas = refGas
          RefAlpha = refAlpha
          RefProducts = refProducts
          RefFormulas = refFormulas
          TargetGas = targetGas
          TargetHeat = targetHeat
          Probes = probes
          OutExists = File.Exists CombustionBoard.outPath
          OutAt = (if File.Exists CombustionBoard.outPath then File.GetLastWriteTime(CombustionBoard.outPath).ToString("yyyy-MM-dd HH:mm:ss") else "(不存在)") }

/// 纯规则：照参考表代数算燃烧计量，再用「绝热焓差」算放热量。不碰真机。
///   绝热条件：显焓(T_ad) = LHV（每 1 Nm³ 煤气的低位热值，反应物按 0℃ 基准、显焓忽略）
///   放热量  ：绝热焓差 = 显焓(T_ad) − 显焓(500℃) = LHV − 显焓(500℃)
///   煤气耗量：B = 吸热量 ÷ 绝热焓差
/// 用户口径：按 636.29 万大卡/hr、α 先按 1.4、只用绝热焓差（不单列 LHV×B 作为放热量）。
module internal CombustionCalc =

    let nm3PerKmol = 22.414
    let kcalToKj = 4.1868

    /// 「焦炉煤气」!E4:E11 设计成分（vol.%，干基）；C2 以上不饱和烃按 C2H4 计，H2S 无数据按 0
    let gas =
        [ "CO", 6.5
          "H2", 57.0
          "CH4", 24.5
          "C2H4", 2.5
          "CO2", 2.5
          "O2", 0.55
          "N2", 6.45
          "H2S", 0.0 ]

    let private v name = gas |> List.find (fun (n, _) -> n = name) |> snd

    let lhvKcal = 4000.0
    let qWanDaKa = 636.2885487306428
    let tOut = 500.0

    /// Magnus 饱和水汽压 hPa。用参考表 100760 Pa / 40.2℃ / 6.2% 复算得 H2O 0.4592 vol.%，
    /// 与 XLL 空气体积成分 0.459228846 一致，故新表可用它作独立校验。
    let pSatHpa (t: float) = 6.112 * exp (17.67 * t / (t + 243.5))

    /// 湿空气成分 vol.%（H2O, O2, N2），干空气固定 21 / 79
    let airComposition (pPa: float) (t: float) (phiPct: float) =
        let yH2O = phiPct / 100.0 * pSatHpa t * 100.0 / pPa
        yH2O * 100.0, (1.0 - yH2O) * 21.0, (1.0 - yH2O) * 79.0

    type Meter =
        { O2Need: float
          Theory: (string * float) list
          IdealAir: (string * float) list
          RealAir: (string * float) list
          Products: (string * float) list
          WetTotal: float
          DryTotal: float }

    /// 每 100 Nm³ 煤气的 kmol；污染物（H2S/SO2）按 0，其余照参考表代数
    let meter (alpha: float) (airH2O, airO2, airN2) : Meter =
        let o2need =
            (0.5 * v "CO" + 0.5 * v "H2" + 2.0 * v "CH4" + 3.0 * v "C2H4" + 1.5 * v "H2S" - v "O2")
            / nm3PerKmol

        let theory =
            [ "H2O", (v "H2" + 2.0 * v "CH4" + 2.0 * v "C2H4" + v "H2S") / nm3PerKmol
              "O2", -o2need
              "N2", v "N2" / nm3PerKmol
              "CO2", (v "CO" + v "CH4" + 2.0 * v "C2H4" + v "CO2") / nm3PerKmol
              "SO2", v "H2S" / nm3PerKmol ]

        let totalAir = o2need / (airO2 / 100.0)

        let idealAir =
            [ "H2O", totalAir * airH2O / 100.0
              "O2", totalAir * airO2 / 100.0
              "N2", totalAir * airN2 / 100.0 ]

        let realAir = idealAir |> List.map (fun (n, x) -> n, x * alpha)
        let pick n = realAir |> List.find (fun (m, _) -> m = n) |> snd

        let products =
            theory
            |> List.map(fun (n, x) ->
                let add =
                    match n with
                    | "H2O" -> pick "H2O"
                    | "O2" -> pick "O2"
                    | "N2" -> pick "N2"
                    | _ -> 0.0

                n, x + add)

        { O2Need = o2need
          Theory = theory
          IdealAir = idealAir
          RealAir = realAir
          Products = products
          WetTotal = products |> List.sumBy snd
          DryTotal = products |> List.filter (fun (n, _) -> n <> "H2O") |> List.sumBy snd }

    /// NIST Shomate 单套系数（0~500℃ 段；上一轮对参考工况 130→180℃ 偏差 -0.27%）
    let private coefs =
        [ "H2O", (30.09200, 6.832514, 6.793435, -2.534480, 0.082139)
          "O2", (30.03235, 8.772972, -3.988133, 0.788313, -0.741599)
          "N2", (19.50583, 19.88705, -8.598535, 1.369784, 0.527601)
          "CO2", (24.99735, 55.18696, -33.69137, 7.948387, -0.136638)
          "SO2", (21.43049, 74.35094, -57.75217, 16.35534, 0.086731) ]

    let private shomate (a, b, c, d, e) t =
        a * t + b * t * t / 2.0 + c * (t ** 3.0) / 3.0 + d * (t ** 4.0) / 4.0 - e / t

    let private t0 = 0.27315

    /// 0℃ → celsius 的摩尔显焓 kJ/kmol
    let sensible (name: string) (celsius: float) =
        match coefs |> List.tryFind (fun (n, _) -> n = name) with
        | Some(_, c) -> (shomate c ((celsius + 273.15) / 1000.0) - shomate c t0) * 1000.0
        | None -> 0.0

    /// 每 1 Nm³ 煤气燃烧产物在 celsius 的显焓，kJ/Nm³
    let mixSensible (products: (string * float) list) (celsius: float) =
        products |> List.sumBy(fun (n, k) -> k / 100.0 * sensible n celsius)

    type Solved =
        { Meter: Meter
          H500: float
          QEff: float
          GasNm3: float
          AirNm3: float
          WetNm3: float
          DryNm3: float
          TAd: float
          QKj: float }

    let solve (alpha: float) (air: float * float * float) : Solved =
        let m = meter alpha air
        let lhvKj = lhvKcal * kcalToKj
        let h500 = mixSensible m.Products tOut
        let qEff = lhvKj - h500
        let qKj = qWanDaKa * 10000.0 * kcalToKj
        let b = qKj / qEff

        // T_ad 由 Excel 里的 gasMoleEnthalpy 二分迭代求；这里用同一套单系数外推，仅作数量级校验
        let tAd =
            let rec go lo hi n =
                if n = 0 then
                    (lo + hi) / 2.0
                else
                    let mid = (lo + hi) / 2.0

                    if mixSensible m.Products mid < lhvKj then
                        go mid hi (n - 1)
                    else
                        go lo mid (n - 1)

            go 0.0 3000.0 60

        { Meter = m
          H500 = h500
          QEff = qEff
          GasNm3 = b
          AirNm3 = b * (m.RealAir |> List.sumBy snd) * nm3PerKmol / 100.0
          WetNm3 = b * m.WetTotal * nm3PerKmol / 100.0
          DryNm3 = b * m.DryTotal * nm3PerKmol / 100.0
          TAd = tAd
          QKj = qKj }

    /// α 依据（2026-11-10 查）：本表 1.4；flue reheat.xlsx 煤气/天然气 C21=1.15；
    /// 天然气燃烧器 C4=1.15；燃气燃烧器典型 1.05~1.25（炼油管式炉烧气 1.05~1.15、
    /// Noxmat 残余 O2 2.5~3.5% 对应 λ≈1.13~1.2）
    let alphaSources =
        [ "本表（用户指定，先按）", 1.4
          "flue reheat.xlsx 煤气/天然气 C21", 1.15
          "天然气燃烧计算.xlsx 燃烧器 C4", 1.15
          "燃气燃烧器典型下限", 1.05
          "燃气燃烧器典型上限", 1.25 ]

type GasCombustionTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``燃烧计算清点：只读现场与参考表计算链``() =
        Report.title output "燃烧计算清点：挂回常驻 Excel，核对两本表与 XLL 函数"
        let info = CombustionBoard.info()
        let before = Report.procIds info.Image
        Report.line output "黑板" (info.describe())
        Report.line output "参考表" CombustionBoard.refPath
        Report.line output "目标书" CombustionBoard.targetPath
        Report.line output "输出" CombustionBoard.outPath

        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

        let snap = STA.run (fun () -> CombustionSurvey.survey app)

        for line in snap.Opened do
            Report.line output "打开的工作簿" line

        for (a, v) in snap.RefAir do
            Report.line output (sprintf "参考 大气 %s" a) v

        for (a, v) in snap.RefGas do
            Report.line output (sprintf "参考 煤气 %s" a) v

        Report.line output "参考 α" snap.RefAlpha

        for (a, v) in snap.RefProducts do
            Report.line output (sprintf "参考 过剩空气 %s" a) v

        for line in snap.RefFormulas do
            Report.line output "参考 公式" line

        for (a, v) in snap.TargetGas do
            Report.line output (sprintf "目标 输入 %s" a) v

        for (a, v) in snap.TargetHeat do
            Report.line output (sprintf "目标 吸热 %s" a) v

        for line in snap.Probes do
            Report.line output "XLL 探测" line

        Report.line output "输出文件" (sprintf "已存在=%b　最后写入=%s" snap.OutExists snap.OutAt)
        let after = Report.procIds info.Image
        Report.line output "进程" (sprintf "EXCEL 干活前 %A / 收尾后 %A" (Set.toList before) (Set.toList after))

        Assert.True(snap.Opened.Length >= 2, "参考表与目标书都应在同一只 Excel 里开着")
        Should.equal before after
        Report.conclusion output "通过：只读查明两本表现场、参考表计算链与 XLL 可用性，磁盘一个字节没动"

    [<Fact>]
    member this.``补热燃烧规则：计量与绝热焓差算得对``() =
        Report.title output "补热燃烧规则：仙桃大气 + 焦炉煤气计量 + 绝热焓差放热量"
        let air = CombustionCalc.airComposition CombustionBoard.pressurePa CombustionBoard.airCelsius CombustionBoard.humidityPct
        let refAir = CombustionCalc.airComposition 100760.0 40.2 6.2
        Report.line output "仙桃大气" (sprintf "100760 Pa / 16.5℃ / 80%% → H2O %.4f O2 %.4f N2 %.4f vol.%%" (let a,b,c = air in a) (let a,b,c = air in b) (let a,b,c = air in c))
        Report.line output "口径校验" (sprintf "参考表 40.2℃/6.2%% 复算 H2O %.6f vol.%%（XLL 缓存 0.459229）" (let a,_,_ = refAir in a))
        Should.equal true (abs ((let a,_,_ = refAir in a) - 0.45922884683371556) < 0.01)

        let s = CombustionCalc.solve 1.4 air
        Report.line output "理论耗 O2" (sprintf "%.6f kmol / 100 Nm³ 煤气" s.Meter.O2Need)
        Report.line output "理论空气量" (sprintf "%.4f Nm³/Nm³ 煤气" ((s.Meter.IdealAir |> List.sumBy snd) * CombustionCalc.nm3PerKmol / 100.0))
        Report.line output "实际空气量 α=1.4" (sprintf "%.4f Nm³/Nm³ 煤气" ((s.Meter.RealAir |> List.sumBy snd) * CombustionCalc.nm3PerKmol / 100.0))
        Report.line output "实际湿烟气量" (sprintf "%.4f Nm³/Nm³ 煤气" (s.Meter.WetTotal * CombustionCalc.nm3PerKmol / 100.0))
        Report.line output "实际干烟气量" (sprintf "%.4f Nm³/Nm³ 煤气" (s.Meter.DryTotal * CombustionCalc.nm3PerKmol / 100.0))

        for (n, k) in s.Meter.Products do
            Report.line output (sprintf "产物 %s" n) (sprintf "湿基 %.4f vol.%%　干基 %.4f vol.%%" (k / s.Meter.WetTotal * 100.0) (if n = "H2O" then 0.0 else k / s.Meter.DryTotal * 100.0))

        Report.line output "低位热值" (sprintf "%.1f kcal/Nm³ = %.1f kJ/Nm³" CombustionCalc.lhvKcal (CombustionCalc.lhvKcal * CombustionCalc.kcalToKj))
        Report.line output "显焓 500℃" (sprintf "%.1f kJ/Nm³ 煤气" s.H500)
        Report.line output "绝热焓差" (sprintf "%.1f kJ/Nm³ 煤气 = %.1f kcal/Nm³" s.QEff (s.QEff / CombustionCalc.kcalToKj))
        Report.line output "T_ad 参考" (sprintf "%.0f ℃（单套 Shomate 外推，只作数量级校验；Excel 主链用 gasMoleEnthalpy 迭代）" s.TAd)
        Report.line output "吸热量" (sprintf "%.2f 万大卡/hr = %.0f kJ/h" CombustionCalc.qWanDaKa s.QKj)
        Report.line output "煤气耗量" (sprintf "%.1f Nm³/h" s.GasNm3)
        Report.line output "助燃空气量" (sprintf "%.0f Nm³/h" s.AirNm3)
        Report.line output "湿烟气量" (sprintf "%.0f Nm³/h" s.WetNm3)
        Report.line output "放热量闭合" (sprintf "B × 绝热焓差 = %.0f kJ/h（= 吸热量 %.0f）" (s.GasNm3 * s.QEff) s.QKj)

        Should.equal true (abs (s.GasNm3 * s.QEff / s.QKj - 1.0) < 1e-9)
        Should.equal true (s.QEff > 0.0 && s.QEff < CombustionCalc.lhvKcal * CombustionCalc.kcalToKj)
        Should.equal 100.0 (Math.Round(CombustionCalc.gas |> List.sumBy snd, 6))

        for (label, a) in CombustionCalc.alphaSources |> List.filter (fun (n, _) -> n.Contains "C21" || n.Contains "本表") do
            let x = CombustionCalc.solve a air
            Report.line output (sprintf "α 对照 %s" label) (sprintf "α=%.2f → 实际空气 %.3f Nm³/Nm³、绝热焓差 %.0f kJ/Nm³、煤气 %.0f Nm³/h" a ((x.Meter.RealAir |> List.sumBy snd) * CombustionCalc.nm3PerKmol / 100.0) x.QEff x.GasNm3)

        Report.conclusion output "通过：仙桃湿空气成分复现参考 XLL，燃烧计量自洽，绝热焓差放热量与吸热量闭合；纯逻辑不碰真机"

    [<Fact>]
    member this.``补热燃烧回填：另存补热焦炉煤气燃烧计算书``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_COMBUSTION" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_COMBUSTION=1 才真的生成新文件"
        | _ ->
            let info = CombustionBoard.info ()
            let before = Report.procIds info.Image
            let target = CombustionBoard.targetPath
            let outPath = CombustionBoard.outPath
            Assert.True(File.Exists target, sprintf "原书不在：%s" target)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            let backupDir = Path.Combine(projectRoot, "TestResults")
            Directory.CreateDirectory backupDir |> ignore

            let targetBackup = Path.Combine(backupDir, sprintf "焦炉煤气燃烧计算书[补热计算前]%s.xlsx" stamp)
            if not (File.Exists targetBackup) then
                File.Copy(target, targetBackup, false)

            // 目标已存在时先留一份旧版，再覆盖（覆盖的是本 Fact 自己生成的副本，不是用户原件）
            let oldBackup =
                if File.Exists outPath then
                    let p = Path.Combine(backupDir, sprintf "补热焦炉煤气燃烧计算书[上一版]%s.xlsx" stamp)
                    if not (File.Exists p) then File.Copy(outPath, p, false)
                    p
                else
                    ""

            // 复制原书到新路径：用户手上那本工作簿全程不开、不改、不 SaveAs
            File.Copy(target, outPath, true)

            Report.title output "补热燃烧回填：另存「补热焦炉煤气燃烧计算书.xlsx」"
            Report.line output "黑板" (info.describe ())
            Report.line output "原书" target
            Report.line output "输出" outPath
            Report.line output "原书备份" targetBackup
            Report.line output "旧版备份" (if oldBackup = "" then "(输出文件原先不存在)" else oldBackup)
            Report.line output "存盘前" (File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss"))

            let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
            Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

            let result =
                STA.run (fun () ->
                    let book = ExcelHost.BookIO.openBook app outPath false
                    let sheetName = "补热燃烧计算"

                    let sheet =
                        [ 1 .. book.Worksheets.Count ]
                        |> List.tryPick (fun i ->
                            let s = book.Worksheets.Item(i) :?> Worksheet
                            if s.Name = sheetName then Some s else None)
                        |> function
                        | Some s -> s
                        | None ->
                            let s = book.Worksheets.Add() :?> Worksheet
                            s.Name <- sheetName
                            // 挪到最后：失败也不影响计算，只是位置不同
                            (try
                                s.Move(box Type.Missing, box (book.Worksheets.Item(book.Worksheets.Count)))
                             with _ ->
                                ())
                            s

                    // 先收集 (地址, 值)，最后统一落笔；公式一律 Formula2
                    let writes = ResizeArray<string * obj>()
                    let put (a: string) (v: obj) = writes.Add(a, v)

                    /// 一行：名称 / 单位 / 数值或公式（null 表示由 spill 填） / 备注
                    let line r name unit value note =
                        put (sprintf "A%d" r) (box name)
                        if unit <> "" then put (sprintf "B%d" r) (box unit)
                        if not (isNull value) then put (sprintf "C%d" r) value
                        if note <> "" then put (sprintf "D%d" r) (box note)

                    let f x = box x

                    line 1 "名称" "单位" (f "数值") "备注"
                    line 2 "一、大气与燃料输入（湖北仙桃）" "" null "大气数据按仙桃多年平均，三项均为可改输入格"
                    line 3 "大气压力" "Pa" (box CombustionBoard.pressurePa) "仙桃气象站多年平均气压 1006.5 hPa；与参考表 100760 Pa 一致"
                    line 4 "大气温度" "℃" (box CombustionBoard.airCelsius) "仙桃市政府/气象局多年平均气温 16.5~17.0℃"
                    line 5 "大气相对湿度" "%" (box CombustionBoard.humidityPct) "仙桃多年平均相对湿度 79~80%"
                    line 6 "湿空气 H2O" "vol. %" (f "=空气体积成分(C3,C4,C5)") "XLL 数组函数 spill 到 C6:C8；用 Formula2 写才不会被补 @"
                    line 7 "湿空气 O2" "vol. %" null ""
                    line 8 "湿空气 N2" "vol. %" null ""
                    line 9 "Magnus 独立复算 H2O" "vol. %" (f "=100*(C5/100*6.112*EXP(17.67*C4/(C4+243.5))*100/C3)") "不依赖插件的校验：应与 C6 一致"

                    line 11 "二、焦炉煤气成分（活链接「焦炉煤气」设计值）" "" null "C2 以上不饱和烃按 C2H4 计"
                    line 12 "CO" "vol. %" (f "='焦炉煤气'!E6") ""
                    line 13 "H2" "vol. %" (f "='焦炉煤气'!E7") ""
                    line 14 "CH4" "vol. %" (f "='焦炉煤气'!E8") ""
                    line 15 "C2H4（C2 以上不饱和烃）" "vol. %" (f "='焦炉煤气'!E10") ""
                    line 16 "CO2" "vol. %" (f "='焦炉煤气'!E5") ""
                    line 17 "O2" "vol. %" (f "='焦炉煤气'!E9") ""
                    line 18 "N2" "vol. %" (f "='焦炉煤气'!E4") ""
                    line 19 "H2S" "vol. %" (box 0.0) "污染物：燃料未给 S，按 0"
                    line 20 "合计" "vol. %" (f "=SUM(C12:C19)") "应为 100.00"
                    line 21 "低位热值 LHV" "kcal/Nm³" (f "='焦炉煤气'!E3") "设计值取区间下限 4000"
                    line 22 "低位热值 LHV" "kJ/Nm³" (f "=C21*4.1868") "绝热条件的目标焓"
                    line 23 "过量空气系数 α" "—" (box 1.4) "可改输入格；依据与 1.15 对照见第 95~100 行"
                    line 24 "排烟/混合终温" "℃" (f "='需要加热的烟气参数'!D5") "燃烧产物混入被加热烟气后的终温"
                    line 25 "烟气吸热量" "万大卡/hr" (f "='烟气吸热量计算'!C36") "按 57000 Nm³/h × 175→500℃ 焓差；业主给定 530 仅作对照，未采用"
                    line 26 "烟气吸热量" "kJ/h" (f "=C25*10000*4.1868") ""

                    line 28 "三、燃烧计量（每 100 Nm³ 煤气，照 煤气燃烧计算.xlsx，污染物按 0）" "" null ""
                    line 29 "理论耗 O2" "kmol" (f "=-(0.5*C12+0.5*C13+2*C14+3*C15+1.5*C19-C17)/22.414") "CO/H2 各 0.5、CH4 2、C2H4 3，扣燃料自带 O2"
                    line 30 "理论产物 H2O" "kmol" (f "=(C13+2*C14+2*C15+C19)/22.414") ""
                    line 31 "理论产物 O2" "kmol" (f "=C29") "负值＝理论耗氧"
                    line 32 "理论产物 N2" "kmol" (f "=C18/22.414") ""
                    line 33 "理论产物 CO2" "kmol" (f "=(C12+C14+2*C15+C16)/22.414") ""
                    line 34 "理论产物 SO2" "kmol" (f "=C19/22.414") "污染物按 0"
                    line 35 "理论空气量 H2O" "kmol" (f "=理论空气量($C$6:$C$8,C29)") "XLL 数组函数 spill 到 C35:C37"
                    line 36 "理论空气量 O2" "kmol" null ""
                    line 37 "理论空气量 N2" "kmol" null ""
                    line 38 "理论空气量合计" "kmol" (f "=SUM(C35:C37)") ""
                    line 39 "实际空气量 H2O" "kmol" (f "=实际空气量($C$6:$C$8,$C$23,$C$29)") "XLL 数组函数 spill 到 C39:C41"
                    line 40 "实际空气量 O2" "kmol" null ""
                    line 41 "实际空气量 N2" "kmol" null ""
                    line 42 "实际空气量合计" "kmol" (f "=SUM(C39:C41)") ""
                    line 43 "实际烟气量 H2O" "kmol" (f "=实际烟气量(C39:C41,C30:C34)") "XLL 数组函数 spill 到 C43:C47"
                    line 44 "实际烟气量 O2" "kmol" null ""
                    line 45 "实际烟气量 N2" "kmol" null ""
                    line 46 "实际烟气量 CO2" "kmol" null ""
                    line 47 "实际烟气量 SO2" "kmol" null "污染物按 0"
                    line 49 "实际湿烟气量" "kmol" (f "=SUM(C43:C47)") ""
                    line 50 "实际干烟气量" "kmol" (f "=SUM(C44:C47)") ""
                    line 51 "湿基 H2O" "vol. %" (f "=烟气湿基体积分数(C43:C47,C49)") "XLL 数组函数 spill 到 C51:C55"
                    line 52 "湿基 O2" "vol. %" null ""
                    line 53 "湿基 N2" "vol. %" null ""
                    line 54 "湿基 CO2" "vol. %" null ""
                    line 55 "湿基 SO2" "vol. %" null "污染物按 0"
                    line 57 "干基 H2O" "vol. %" (f "=烟气干基体积分数(C43:C47,C50)") "XLL 数组函数 spill 到 C57:C61"
                    line 58 "干基 O2" "vol. %" null ""
                    line 59 "干基 N2" "vol. %" null ""
                    line 60 "干基 CO2" "vol. %" null ""
                    line 61 "干基 SO2" "vol. %" null "污染物按 0"
                    line 63 "每 1 Nm³ 煤气产物 H2O" "kmol" (f "=C43/100") "焓值函数的 kmol 输入"
                    line 64 "每 1 Nm³ 煤气产物 O2" "kmol" (f "=C44/100") ""
                    line 65 "每 1 Nm³ 煤气产物 N2" "kmol" (f "=C45/100") ""
                    line 66 "每 1 Nm³ 煤气产物 CO2" "kmol" (f "=C46/100") ""
                    line 67 "每 1 Nm³ 煤气产物 SO2" "kmol" (f "=C47/100") "污染物按 0"
                    line 68 "合计" "kmol" (f "=SUM(C63:C67)") ""

                    line 70 "四、绝热焓差放热量（前后焓值变化）" "" null "T_ad 由 显焓(T_ad)=LHV 反解，见 F:L 列 20 步二分"
                    put "A71" (box "组分")
                    put "B71" (box "kmol/Nm³")
                    put "C71" (box "说明")
                    for (r, name, src) in [ 72, "H2O", 63; 73, "O2", 64; 74, "N2", 65; 75, "CO2", 66; 76, "SO2", 67 ] do
                        put (sprintf "A%d" r) (box name)
                        put (sprintf "B%d" r) (f (sprintf "=C%d" src))

                    line 77 "焓值基准 H(0℃)" "kJ/Nm³" (f "=gasMoleEnthalpy($A$72:$A$76,$B$72:$B$76,0)") "XLL 总焓含常数，后续一律取差值"
                    line 78 "终温焓 H(500℃)" "kJ/Nm³" (f "=gasMoleEnthalpy($A$72:$A$76,$B$72:$B$76,C24)") ""
                    line 79 "显焓(0→500℃)" "kJ/Nm³" (f "=C78-C77") "燃烧产物在终温带走的显热"
                    line 80 "低位热值（绝热目标）" "kJ/Nm³" (f "=C22") ""
                    line 81 "绝热火焰温度 T_ad" "℃" (f "=(K92+L92)/2") "20 步二分收敛，见 F:L 列"
                    line 82 "显焓(0→T_ad)" "kJ/Nm³" (f "=gasMoleEnthalpy($A$72:$A$76,$B$72:$B$76,C81)-C77") "绝热条件下应等于 LHV"
                    line 83 "显焓闭合校核" "—" (f "=C82/C80") "应约等于 1"
                    line 84 "绝热焓差（放热量）" "kJ/Nm³ 煤气" (f "=C82-C79") "放热量＝前后焓值变化；不单列 LHV×B"
                    line 85 "绝热焓差" "kcal/Nm³" (f "=C84/4.1868") ""
                    line 86 "煤气耗量" "Nm³/h" (f "=C26/C84") "补热燃烧器设计煤气量"
                    line 87 "放热量" "kJ/h" (f "=C86*C84") ""
                    line 88 "放热量" "万大卡/hr" (f "=C87/4.1868/10000") "应等于 C25"
                    line 89 "助燃空气量" "Nm³/h" (f "=C86*C42*22.414/100") ""
                    line 90 "湿烟气量" "Nm³/h" (f "=C86*C49*22.414/100") ""
                    line 91 "干烟气量" "Nm³/h" (f "=C86*C50*22.414/100") ""
                    line 92 "放热/吸热闭合" "—" (f "=C87/C26") "应约等于 1"

                    // F:L 列：T_ad 二分迭代（20 步，0~2500℃）
                    put "F71" (box "T_ad 二分迭代（20 步）")
                    for (col, text) in [ "F", "步"; "G", "low"; "H", "high"; "I", "mid"; "J", "显焓(mid)"; "K", "新 low"; "L", "新 high" ] do
                        put (sprintf "%s72" col) (box text)

                    for i in 0 .. 19 do
                        let r = 73 + i
                        put (sprintf "F%d" r) (box (i + 1))

                        if i = 0 then
                            put (sprintf "G%d" r) (box 0.0)
                            put (sprintf "H%d" r) (box 2500.0)
                        else
                            put (sprintf "G%d" r) (f (sprintf "=K%d" (r - 1)))
                            put (sprintf "H%d" r) (f (sprintf "=L%d" (r - 1)))

                        put (sprintf "I%d" r) (f (sprintf "=(G%d+H%d)/2" r r))
                        put (sprintf "J%d" r) (f (sprintf "=gasMoleEnthalpy($A$72:$A$76,$B$72:$B$76,I%d)-$C$77" r))
                        put (sprintf "K%d" r) (f (sprintf "=IF(J%d<$C$80,I%d,G%d)" r r r))
                        put (sprintf "L%d" r) (f (sprintf "=IF(J%d<$C$80,H%d,I%d)" r r r))

                    line 94 "五、口径、依据与对照" "" null ""
                    let note r text = put (sprintf "A%d" r) (box text)
                    note 95 "α 取值：本表 1.4（用户指定）；改 C23 输入格可全表重算"
                    note 96 "α 依据 1：flue reheat.xlsx「煤气」「天然气」C21 = 1.15，备注「助燃风常量」"
                    note 97 "α 依据 2：天然气燃烧计算.xlsx「天然气燃烧器」C4 = 1.15"
                    note 98 "α 依据 3：煤气燃烧计算.xlsx「过剩空气」C15 = 1.4（本表沿用）"
                    note 99 "α 依据 4：外网查得燃气燃烧器典型 1.05~1.25（炼油管式炉烧气 1.05~1.15；Noxmat 残余 O2 2.5~3.5% 对应 λ≈1.13~1.2）；1.4 偏高，建议复核"
                    note 100 "α=1.15 离线对照：实际空气 4.875 Nm³/Nm³、绝热焓差 12743 kJ/Nm³、煤气 2091 Nm³/h（比 1.4 少约 5.6%）；把 C23 改成 1.15 即全表重算"
                    note 101 "大气数据（湖北仙桃）：100760 Pa、16.5℃、80%；仙桃气象站多年平均气压 1006.5 hPa，市政府/气象局多年平均气温 16.5~17.0℃、相对湿度 79~80%"
                    note 102 "吸热量按「烟气吸热量计算」!C36 = 636.29 万大卡/hr；业主给定「需要加热的烟气参数」!D2 = 530 万大卡/hr 仅作对照，未采用"
                    note 103 "放热量口径：只用绝热焓差。T_ad 由 显焓(T_ad)=LHV 反解，放热量 = 显焓(T_ad) − 显焓(500℃)；不把 LHV×B 单列为放热量"
                    note 104 "反应物（燃料与助燃空气）按 0℃ 基准，显焓忽略（仙桃 16.5℃，影响小于 1%）"
                    note 105 "污染物 SO2/SO3/HCl/HF/ash 全部按 0；参考表的污染物浓度折算、质量百分比等单元格未复制"
                    note 106 "焓值函数 gasMoleEnthalpy（本机 XLL 已注册）；高温段以插件内部系数为准，F# Shomate 单系数外推 T_ad≈1590℃ 仅作数量级校验"
                    note 107 "所有公式用 Range.Formula2 写入：Range.Formula 会给自定义函数补 @ 隐式交集，XLL 数组函数无法 spill"
                    note 108 "C9 用 Magnus 饱和水汽压独立复算湿空气 H2O，与 XLL 的 C6 互为校验"

                    // 幂等校验：重跑时允许覆盖本 Fact 上次写下的内容，但绝不盖掉别人新写的东西
                    let existing =
                        let gridObj = sheet.Range("A1:L120").Value2

                        if isNull gridObj then
                            []
                        else
                            let grid = gridObj :?> obj[,]
                            let r0, c0 = grid.GetLowerBound(0), grid.GetLowerBound(1)

                            [ for i in r0 .. grid.GetUpperBound(0) do
                                  for j in c0 .. grid.GetUpperBound(1) do
                                      let v = grid.[i, j]

                                      if not (isNull v) && Convert.ToString(v, CultureInfo.InvariantCulture) <> "" then
                                          yield sprintf "%c%d" (char (int 'A' + j - c0)) (i - r0 + 1) ]

                    // spill 成员格由数组公式自动填，不在写入清单里；重跑时它们也算「自己的内容」。
                    // 落笔前先清 spill 起点（Excel 会连带清掉整片动态数组），否则重跑会 #SPILL!。
                    let spillAnchors = [ "C6"; "C35"; "C39"; "C43"; "C51"; "C57" ]

                    let spillMembers =
                        [ "C7"; "C8"; "C36"; "C37"; "C40"; "C41"
                          "C44"; "C45"; "C46"; "C47"
                          "C52"; "C53"; "C54"; "C55"
                          "C58"; "C59"; "C60"; "C61" ]

                    let allowed =
                        Set.union (writes |> Seq.map fst |> Set.ofSeq) (Set.ofList (spillAnchors @ spillMembers))

                    let stray = existing |> List.filter (fun a -> not (Set.contains a allowed))

                    if not stray.IsEmpty then
                        failwithf "新表里已有本次清单之外的非空单元格，停手：%s" (String.concat "、" stray)

                    for a in spillAnchors do
                        (try
                            sheet.Range(a).ClearContents() |> ignore
                         with _ ->
                            ())

                    for (addr, value) in writes do
                        let r = sheet.Range(addr)

                        match value with
                        | :? string as s when s.StartsWith("=") -> r.Formula2 <- box s
                        | _ -> r.Value2 <- value

                    sheet.Calculate()
                    book.Save()

                    let cellText a =
                        let v = sheet.Range(a).Value2

                        match v with
                        | null -> ""
                        | :? int as i when i < -2000000000 -> sprintf "错误码 %d" i
                        | _ -> Convert.ToString(v, CultureInfo.InvariantCulture)

                    let addrs =
                        [ "C6"; "C7"; "C8"; "C9"; "C20"; "C22"; "C23"; "C24"; "C25"; "C26"
                          "C29"; "C38"; "C42"; "C49"; "C50"; "C51"; "C52"; "C53"; "C54"; "C58"
                          "C68"; "C77"; "C78"; "C79"; "C80"; "C81"; "C82"; "C83"; "C84"; "C85"
                          "C86"; "C87"; "C88"; "C89"; "C90"; "C91"; "C92"; "J73"; "J92"; "I92" ]

                    let readBack = addrs |> List.map (fun a -> a, cellText a)
                    let saved = try book.Saved with _ -> true
                    let sheetCount = book.Worksheets.Count
                    ExcelHost.BookIO.closeBook book

                    {| ReadBack = readBack
                       Saved = saved
                       SheetCount = sheetCount
                       Stray = stray
                       Writes = writes.Count
                       DiskAt = File.GetLastWriteTime(outPath).ToString("yyyy-MM-dd HH:mm:ss")
                       OutBytes = FileInfo(outPath).Length
                       TargetAt = File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss") |})

            for (a, v) in result.ReadBack do
                Report.line output (sprintf "读回 %s" a) v

            let value a = result.ReadBack |> List.find (fun (x, _) -> x = a) |> snd
            let number a = float (value a)

            Report.line output "写入格数" (sprintf "%d 格；表数 %d；Saved=%b" result.Writes result.SheetCount result.Saved)
            Report.line output "磁盘" (sprintf "输出 %d B　保存于 %s；原书最后写入仍为 %s" result.OutBytes result.DiskAt result.TargetAt)
            let after = Report.procIds info.Image
            Report.line output "进程" (sprintf "EXCEL 干活前 %A / 收尾后 %A" (Set.toList before) (Set.toList after))

            Assert.True(File.Exists targetBackup, "原书备份应存在")
            Assert.True(result.Saved, "副本应已保存")
            Assert.True(result.OutBytes > FileInfo(target).Length, "输出应比原书大（多一张表）")
            Assert.True(result.ReadBack |> List.forall (fun (_, v) -> not (v.Contains "错误码")), "读回值里不应有 Excel 错误码")
            Assert.True(abs (number "C83" - 1.0) < 0.02, sprintf "显焓闭合校核 C83 应约 1，实际 %s" (value "C83"))
            Assert.True(abs (number "C92" - 1.0) < 0.02, sprintf "放热/吸热闭合 C92 应约 1，实际 %s" (value "C92"))
            Assert.True(abs (number "C6" - number "C9") < 0.02, "Magnus 复算应与 XLL 空气体积成分一致")
            Assert.True(abs (number "C20" - 100.0) < 0.01, "煤气成分合计应为 100")

            let expected = CombustionCalc.solve 1.4 (CombustionCalc.airComposition CombustionBoard.pressurePa CombustionBoard.airCelsius CombustionBoard.humidityPct)
            Assert.True(abs (number "C86" / expected.GasNm3 - 1.0) < 0.05, sprintf "煤气耗量 %.1f 应与纯规则 %.1f Nm³/h 相差小于 5%%" (number "C86") expected.GasNm3)
            Should.equal before after

            Report.conclusion output "完成：新文件已另存、原书未动；仙桃大气 + 燃烧计量 + 绝热焓差放热量全为活公式，读回闭合"

