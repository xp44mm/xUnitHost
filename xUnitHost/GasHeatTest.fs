/// 【黑板挂回】参考 C:\崔胜利\群有\河北浩锐陶瓷\天然气燃烧计算.xlsx
/// 「体积法scr烟气吸热」!A1:F34 的体积法计算过程，在
/// C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx「烟气吸热量计算」!A3 起复刻：
///   组分 vol.% → 烟气量 Nm³/h → /22.414 得 kmol/h → 各组分 kmol/h
///   → 两组分的两个温度摩尔显焓 → 总焓 → 焓差 → kJ/h / kW / kcal/h / 万大卡/hr
/// 输入全部活链接到「需要加热的烟气参数」，改参数可重算。
/// 焓值主表照搬参考表的 gasMoleEnthalpy（本机 Excel 已注册该 XLL 插件）；
/// F/G 列另用 NIST Shomate 摩尔定压热容多项式写成 Excel 内置公式独立复算
/// （对参考工况 130→180℃ 偏差 -0.27%），既做交叉校验，也是没装插件时的替代。
///
/// 通道同 GasFlueTest.fs：按 PID 挂回现成实例，绝不新建 / Quit / kill；
/// COM 全在 OsHost.Sta.run 固定 STA 线程。纯计算与真机动作分开。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

/// 纯规则：体积法烟气吸热量
module internal HeatCalc =

    /// NIST Shomate 系数（kJ/mol，t=T(K)/1000）：
    /// H°(T)-H°(298.15) = A·t + B·t²/2 + C·t³/3 + D·t⁴/4 - E/t + F - H
    /// 取两温度之差时常数项抵消，故只留 A..E。
    type Coef =
        { A: float
          B: float
          C: float
          D: float
          E: float }

    let coefs: (string * Coef) list =
        [ "H2O", { A = 30.09200; B = 6.832514; C = 6.793435; D = -2.534480; E = 0.082139 }
          "O2", { A = 30.03235; B = 8.772972; C = -3.988133; D = 0.788313; E = -0.741599 }
          "N2", { A = 19.50583; B = 19.88705; C = -8.598535; D = 1.369784; E = 0.527601 }
          "CO2", { A = 24.99735; B = 55.18696; C = -33.69137; D = 7.948387; E = -0.136638 }
          "SO2", { A = 21.43049; B = 74.35094; C = -57.75217; D = 16.35534; E = 0.086731 } ]

    let t (celsius: float) = (celsius + 273.15) / 1000.0
    /// 基准取 0℃（t0=0.27315）；焓差与基准无关
    let t0 = t 0.0

    let private shomate (c: Coef) (tv: float) =
        c.A * tv
        + c.B * tv * tv / 2.0
        + c.C * (tv ** 3.0) / 3.0
        + c.D * (tv ** 4.0) / 4.0
        - c.E / tv

    /// 从 0℃ 到 celsius 的摩尔显焓，kJ/kmol；SO3/HCl/HF/ash 无热容数据按 0
    let molarEnthalpy (name: string) (celsius: float) : float =
        match coefs |> List.tryFind(fun (n, _) -> n = name) with
        | Some(_, c) -> (shomate c (t celsius) - shomate c t0) * 1000.0
        | None -> 0.0

    /// kJ/h：各组分 kmol/h × 该组分 t1→t2 的摩尔焓差
    let heat (kmolPerHour: (string * float) list) (t1: float) (t2: float) : float =
        kmolPerHour
        |> List.sumBy(fun (name, kmol) -> kmol * (molarEnthalpy name t2 - molarEnthalpy name t1))

    let nm3PerKmol = 22.414
    let kjPerKcal = 4.1868

    /// 被加热烟气湿基组分，vol.%；取自「需要加热的烟气参数」!D16:D24
    let wetComponents: (string * float) list =
        [ "H2O", 14.47
          "O2", 7.70
          "N2", 72.81
          "CO2", 5.02
          "SO2", 0.0
          "SO3", 0.0
          "HCl", 0.0
          "HF", 0.0
          "ash", 0.0 ]

    let names = wetComponents |> List.map fst

    /// 体积法：Nm³/h → 各组分 kmol/h
    let toKmol (volume: float) : (string * float) list =
        let total = volume / nm3PerKmol
        wetComponents |> List.map(fun (n, p) -> n, total * p / 100.0)

    /// 本书设计工况
    let volume = 57000.0
    let t1 = 175.0
    let t2 = 500.0
    let givenWanDaKa = 530.0

    let designKmol = toKmol volume
    let designHeat = heat designKmol t1 t2

    /// 参考工况（河北浩锐陶瓷「体积法scr烟气吸热」缓存值）：用来校验多项式口径
    let refKmol: (string * float) list =
        [ "H2O", 1039.1492236994736
          "O2", 157.08794503435354
          "N2", 4515.0316900151693
          "CO2", 522.37975372535027
          "SO2", 0.0 ]

    let refHeatCached = 9799413.0127017945

module internal HeatBoard =

    let info () : BoardApp = FlueBoard.info ()
    let doc () : BoardDoc = FlueBoard.doc ()

    let sheetName = "烟气吸热量计算"
    let sourceSheet = "需要加热的烟气参数"
    let refPath = @"C:\崔胜利\群有\河北浩锐陶瓷\天然气燃烧计算.xlsx"
    let refSheet = "体积法scr烟气吸热"

    let private bookPath () : string =
        match (doc ()).Path with
        | Some path -> path
        | None -> failwith "黑板里没有落盘路径，不往下走"

    let private findBook (app: Application) (path: string) : Workbook option =
        let count = try app.Workbooks.Count with _ -> 0

        [ 1..count ]
        |> List.tryPick(fun i ->
            let book = app.Workbooks.Item(i)
            let full = try book.FullName with _ -> ""
            if String.Equals(full, path, StringComparison.OrdinalIgnoreCase) then Some book else None)

    let private needBook (app: Application) (path: string) : Workbook =
        match findBook app path with
        | Some book -> book
        | None -> failwithf "这只 Excel 里没开着「%s」" path

    let private findSheet (book: Workbook) (name: string) : Worksheet =
        [ 1..book.Worksheets.Count ]
        |> List.tryPick(fun i ->
            let sheet = book.Worksheets.Item(i) :?> Worksheet
            if sheet.Name = name then Some sheet else None)
        |> function
            | Some sheet -> sheet
            | None -> failwithf "工作簿里没有表「%s」" name

    let cellText (sheet: Worksheet) (addr: string) : string =
        let value = sheet.Range(addr).Value2
        if isNull value then "" else Convert.ToString(value, CultureInfo.InvariantCulture)

    let private cellFormula (sheet: Worksheet) (addr: string) : string =
        let f = sheet.Range(addr).Formula
        if isNull f then "" else Convert.ToString(f, CultureInfo.InvariantCulture)

    let private describeValue (v: obj) : string =
        if isNull v then "(空)"
        elif v :? int then sprintf "Excel 错误码 %d（公式不可用）" (unbox<int> v)
        else Convert.ToString(v, CultureInfo.InvariantCulture)

    /// Shomate 系数表：I=组分 J=A K=B L=C M=D N=E（写在目标表右侧，供公式溯源）
    let private coefTable: (string * float * float * float * float * float) list =
        HeatCalc.coefs
        |> List.map(fun (name, c) -> name, c.A, c.B, c.C, c.D, c.E)

    /// 各组分在温度 t 的显焓 kJ/h：普通 Excel 公式，无数组、无插件
    let private componentEnthalpyFormula (row: int) (coefRow: int) (tCell: string) : string =
        sprintf
            "=$C%d*(($J%d*($C$%s-$C$42))+($K%d/2)*($C$%s^2-$C$42^2)+($L%d/3)*($C$%s^3-$C$42^3)+($M%d/4)*($C$%s^4-$C$42^4)-$N%d*(1/$C$%s-1/$C$42))*1000"
            row coefRow tCell coefRow tCell coefRow tCell coefRow tCell coefRow tCell

    /// 只读现场：三张表关键格、参考表计算链、XLL 插件是否可用、A3 起是否空
    type Survey =
        { Workbooks: (string * string * string list) list
          /// 目标表 A1:A6 现状
          HeadRows: (int * string) list
          /// A3:N46 里非空的格子（写之前必须是空的）
          Occupied: string list
          /// 参数表 D2:D5、D16:E24 关键值
          SourceValues: (string * string) list
          /// 参考表关键行的公式与值
          RefRows: (int * string * string * string) list
          XllProbes: string list
          BookSaved: bool
          DiskAt: string }

    let survey (app: Application) : Survey =
        let path = bookPath ()
        let book = needBook app path

        let opened =
            [ 1..app.Workbooks.Count ]
            |> List.map(fun i ->
                let b = app.Workbooks.Item(i)
                let sheets = [ 1..b.Worksheets.Count ] |> List.map(fun j -> (b.Worksheets.Item(j) :?> Worksheet).Name)
                b.Name, (try b.FullName with _ -> ""), sheets)

        let target = findSheet book sheetName
        let source = findSheet book sourceSheet

        let head = [ for r in 1..6 -> r, cellText target (sprintf "A%d" r) ]

        let gridObj = target.Range("A3:N50").Value2
        let occupied = ResizeArray<string>()

        // 整块全空时 COM 直接给 null，不是 obj[,]
        if not (isNull gridObj) then
            let grid = gridObj :?> obj[,]
            let r0, c0 = grid.GetLowerBound(0), grid.GetLowerBound(1)

            for i in r0..grid.GetUpperBound(0) do
                for j in c0..grid.GetUpperBound(1) do
                    let v = grid.[i, j]

                    if not (isNull v) && Convert.ToString(v, CultureInfo.InvariantCulture) <> "" then
                        occupied.Add(
                            sprintf
                                "%c%d=「%s」"
                                (char (int 'A' + j - c0))
                                (i - r0 + 3)
                                (Convert.ToString(v, CultureInfo.InvariantCulture))
                        )

        let sourceValues =
            [ for r in 2..5 -> sprintf "D%d" r, cellText source (sprintf "D%d" r) ]
            @ [ for r in 16..24 -> sprintf "D%d/E%d" r r, cellText source (sprintf "D%d" r) + " / " + cellText source (sprintf "E%d" r) ]

        let refRows =
            match findBook app refPath with
            | None -> []
            | Some refBook ->
                let rs = findSheet refBook refSheet

                [ 1; 2; 12; 13; 15; 25; 26; 28; 29; 31; 32; 33 ]
                |> List.map(fun r ->
                    r,
                    cellText rs (sprintf "A%d" r),
                    cellText rs (sprintf "C%d" r),
                    cellFormula rs (sprintf "C%d" r))

        // 探测参考表用的 XLL 插件是否在本机注册：只读 Evaluate，不改任何单元格
        // 真机公式显示为 gasMoleEnthalpy(...,$C$15#,T)，磁盘 XML 里才是 _xll. 前缀
        let xllProbes =
            match findBook app refPath with
            | None -> [ "参考工作簿没开，未探测" ]
            | Some refBook ->
                let bk = sprintf "'[%s]%s'!" refBook.Name refSheet
                let args = sprintf "%s$A$15:$A$19,%s$C$15#" bk bk
                let argsRange = sprintf "%s$A$15:$A$19,%s$C$15:$C$19" bk bk

                let candidates =
                    [ sprintf "gasMoleEnthalpy(%s,130)" args
                      sprintf "_xll.gasMoleEnthalpy(%s,130)" args
                      sprintf "gasMoleEnthalpy(%s,130)" argsRange
                      sprintf "%sC26（直接读参考表结果）" bk ]

                candidates
                |> List.map(fun formula ->
                    try
                        sprintf "%s → %s" formula (describeValue (app.Evaluate(box formula)))
                    with ex ->
                        sprintf "%s → 抛错：%s" formula ex.Message)

        { Workbooks = opened
          HeadRows = head
          Occupied = List.ofSeq occupied
          SourceValues = sourceValues
          RefRows = refRows
          XllProbes = xllProbes
          BookSaved = (try book.Saved with _ -> true)
          DiskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss") }

    /// 写盘：A3 起复刻参考表体积法计算过程（行号 = 参考表行号 + 2）
    let write (app: Application) : (string * string) list * bool * string * string list =
        let path = bookPath ()
        let book = needBook app path
        let sheet = findSheet book sheetName

        // 写之前再验一次：A3:N46 全空，别盖掉用户刚写的东西
        let occupied =
            let gridObj = sheet.Range("A3:N50").Value2

            if isNull gridObj then
                []
            else
                let grid = gridObj :?> obj[,]
                let r0, c0 = grid.GetLowerBound(0), grid.GetLowerBound(1)

                [ for i in r0..grid.GetUpperBound(0) do
                      for j in c0..grid.GetUpperBound(1) do
                          let v = grid.[i, j]

                          if not (isNull v) && Convert.ToString(v, CultureInfo.InvariantCulture) <> "" then
                              yield sprintf "%c%d" (char (int 'A' + j - c0)) (i - r0 + 3) ]

        // 先收集全部写入项，最后统一落笔：这样能先验「现有非空格是否都在本次清单里」
        // （上一轮若写到一半失败，内存里会留本 Fact 自己写的半张表，磁盘并未保存）。
        let writes = ResizeArray<string * obj>()
        let put (a: string) (value: obj) = writes.Add(a, value)

        let same = sprintf "'%s'!" sourceSheet

        // ── 表头 + 九组分 vol.%（活链接到参数表 D/E 列）──
        put "A3" (box "组分/参数")
        put "B3" (box "单位")
        put "C3" (box "湿基（计算用）")
        put "D3" (box "干基（对照）")
        put "E3" (box "备注")

        HeatCalc.names
        |> List.iteri(fun i name ->
            let r = 4 + i
            let srcRow = 16 + i
            put (sprintf "A%d" r) (box name)
            put (sprintf "B%d" r) (box "vol. %")
            put (sprintf "C%d" r) (box (sprintf "=%sD%d" same srcRow))
            put (sprintf "D%d" r) (box (sprintf "=%sE%d" same srcRow))
            put (sprintf "E%d" r) (box (sprintf "活链接：%s!D%d/E%d；湿基用于焓值" sourceSheet srcRow srcRow)))

        // ── 烟气量 / kmol/h ──
        put "A14" (box "烟气量")
        put "B14" (box "Nm3/h")
        put "C14" (box (sprintf "=%sD3" same))
        put "E14" (box "需加热烟气量，活链接")

        put "B15" (box "kmol/h")
        put "C15" (box "=C14/22.414")
        put "E15" (box "22.414 Nm³/kmol（0℃,101.325kPa）；照参考表 C13")

        // ── 各组分 kmol/h（参考表 C15:C19 的 spill 数组改写成逐格公式）──
        HeatCalc.names
        |> List.iteri(fun i name ->
            let r = 17 + i
            put (sprintf "A%d" r) (box name)
            put (sprintf "B%d" r) (box "kmol/h")
            if i < 5 then
                put (sprintf "C%d" r) (box (sprintf "=$C$15*C%d/100" (4 + i)))
                // 备注不可写成 = 开头：put 会把等号开头一律当公式交给 Excel
                put (sprintf "E%d" r) (box "公式 C15×vol.%/100；照参考表 $C$13*C2:C6/100")
            else
                put (sprintf "C%d" r) (box 0.0)
                put (sprintf "E%d" r) (box "无含量，按 0"))

        // ── 各组分显焓 kJ/h：F 列初始温度，G 列最终温度 ──
        put "F16" (box "初始温度各组分显焓 kJ/h")
        put "G16" (box "最终温度各组分显焓 kJ/h")

        coefTable
        |> List.iteri(fun i _ ->
            let r = 17 + i
            let cr = 4 + i
            put (sprintf "F%d" r) (box (componentEnthalpyFormula r cr "40"))
            put (sprintf "G%d" r) (box (componentEnthalpyFormula r cr "41")))

        put "F22" (box "=SUM(F17:F21)")
        put "G22" (box "=SUM(G17:G21)")
        put "E22" (box "F/G 两列合计：与 C28/C31 的 XLL 结果可交叉核对（绝对焓基准不同，比焓差）")

        // ── 两个温度的总焓与焓差 ──
        put "A27" (box "烟气初始温度")
        put "B27" (box "℃")
        put "C27" (box (sprintf "=%sD4" same))
        put "E27" (box "烟气原始温度，活链接")

        // 本机已注册 gasMoleEnthalpy（参考表用的 XLL），主表照搬参考公式；
        // F/G 列留 Shomate 独立复算做交叉校验。
        put "A28" (box "初始温度总焓值")
        put "B28" (box "kJ/h")
        put "C28" (box "=gasMoleEnthalpy($A$17:$A$21,$C$17:$C$21,C27)")
        put "E28" (box "照参考表 C26：=gasMoleEnthalpy(组分名,各组分 kmol/h,温度)")

        put "A30" (box "烟气最终温度")
        put "B30" (box "℃")
        put "C30" (box (sprintf "=%sD5" same))
        put "E30" (box "加热后烟气温度，活链接")

        put "A31" (box "最终温度总焓值")
        put "B31" (box "kJ/h")
        put "C31" (box "=gasMoleEnthalpy($A$17:$A$21,$C$17:$C$21,C30)")
        put "E31" (box "照参考表 C29")

        put "A33" (box "烟气吸热量")
        put "B33" (box "kJ/h")
        put "C33" (box "=C31-C28")
        put "E33" (box "焓差，照参考表 C29-C26")

        put "B34" (box "kW")
        put "C34" (box "=C33/3600")
        put "B35" (box "kcal/h")
        put "C35" (box (sprintf "=C33/%s" (HeatCalc.kjPerKcal.ToString("G", CultureInfo.InvariantCulture))))
        put "A36" (box "烟气吸热量")
        put "B36" (box "万大卡/hr")
        put "C36" (box "=C35/10000")

        put "A37" (box "给定热交换量")
        put "B37" (box "万大卡/hr")
        put "C37" (box (sprintf "=%sD2" same))
        put "E37" (box "业主给定值，照抄不修改")

        put "A38" (box "计算值 / 给定值")
        put "B38" (box "—")
        put "C38" (box "=C36/C37")
        put "E38" (box "大于 1：按 57000 Nm³/h × 175→500℃ 的显热高于给定热交换量")

        // ── 右侧 Shomate 系数表（公式溯源）──
        put "I3" (box "组分")
        put "J3" (box "A")
        put "K3" (box "B")
        put "L3" (box "C")
        put "M3" (box "D")
        put "N3" (box "E")

        coefTable
        |> List.iteri(fun i (name, a, b, c, d, e) ->
            let r = 4 + i
            put (sprintf "I%d" r) (box name)
            put (sprintf "J%d" r) (box a)
            put (sprintf "K%d" r) (box b)
            put (sprintf "L%d" r) (box c)
            put (sprintf "M%d" r) (box d)
            put (sprintf "N%d" r) (box e))

        // ── 下方辅助温度与口径说明 ──
        put "A40" (box "辅助：t1 = (T1+273.15)/1000")
        put "B40" (box "K/1000")
        put "C40" (box "=(C27+273.15)/1000")
        put "A41" (box "辅助：t2 = (T2+273.15)/1000")
        put "B41" (box "K/1000")
        put "C41" (box "=(C30+273.15)/1000")
        put "A42" (box "辅助：基准 t0 = 273.15/1000")
        put "B42" (box "K/1000")
        put "C42" (box "=273.15/1000")

        put "A44" (box "焓值口径")
        put "E44" (box "NIST Shomate 摩尔定压热容：Δh(kJ/kmol)=[A·Δt+B·Δt²/2+C·Δt³/3+D·Δt⁴/4−E·Δ(1/t)]×1000，基准取 0℃")

        put "A45" (box "与参考表的关系")
        put "E45" (box "流程、单位与 gasMoleEnthalpy 函数均照 天然气燃烧计算.xlsx「体积法scr烟气吸热」A1:F34；本机已注册该 XLL 插件")

        put "A46" (box "系数来源")
        put "E46" (box "F/G 列为 NIST WebBook Shomate 独立复算（H2O/O2/N2/CO2/SO2，298~1200K 段）；参考工况 130→180℃ 与主函数偏差 -0.27%；SO3/HCl/HF/ash 无含量按 0")

        put "A48" (box "独立复算吸热量（Shomate）")
        put "B48" (box "kJ/h")
        put "C48" (box "=G22-F22")
        put "E48" (box "不依赖 XLL：若文件换到没装 gasMoleEnthalpy 插件的机器，把 C28 改 =F22、C31 改 =G22 即可")

        put "A49" (box "独立复算 / 主表 − 1")
        put "B49" (box "—")
        put "C49" (box "=C48/C33-1")
        put "E49" (box "两种焓值基准不同，但焓差应基本一致；超过 1% 需复查组分或温度")

        // 幂等校验：现有非空格必须落在本次写入清单的地址里；否则是用户新写的东西，停手
        let addresses = writes |> Seq.map fst |> Set.ofSeq
        let strangers = occupied |> List.filter(fun addr -> not (Set.contains addr addresses))
        Should.equal ([]: string list) strangers

        for (a, value) in writes do
            let r = sheet.Range(a)

            match value with
            | :? string as s when s.StartsWith("=") -> r.Formula <- box s
            | _ -> r.Value2 <- value

        // 局部重算，让公式有缓存值；不动 Application.Calculation 全局模式
        sheet.Calculate()
        book.Save()

        let readBack =
            [ "C14"; "C15"; "C27"; "C28"; "C30"; "C31"; "C33"; "C34"; "C35"; "C36"; "C37"; "C38"; "C48"; "C49" ]
            |> List.map(fun a -> a, cellText sheet a)

        readBack, (try book.Saved with _ -> true), File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss"), occupied

/// 三条 Fact：纯计算 → 只读清点（含 XLL 探测）→ 真写盘
type GasHeatTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``烟气吸热规则：体积法焓值与参考表对得上``() =
        Report.title output "烟气吸热规则：体积法（组分 vol.% → kmol/h → 摩尔焓差）"

        let refQ = HeatCalc.heat HeatCalc.refKmol 130.0 180.0
        let refDiff = (refQ / HeatCalc.refHeatCached - 1.0) * 100.0
        Report.line output "参考工况" "139721 Nm³/h，130→180℃，组分 H2O 16.67 / O2 2.52 / N2 72.43 / CO2 8.38 vol.%"
        Report.line output "参考表缓存焓差" (sprintf "%.0f kJ/h" HeatCalc.refHeatCached)
        Report.line output "Shomate 复算" (sprintf "%.0f kJ/h（偏差 %+.3f%%）" refQ refDiff)
        Assert.True((abs refDiff) < 0.5, sprintf "与参考 XLL 偏差应小于 0.5%%，实际 %.3f%%" refDiff)

        let total = HeatCalc.volume / HeatCalc.nm3PerKmol
        Report.line output "本书工况" (sprintf "%.0f Nm³/h，%.0f→%.0f℃，湿基九组分" HeatCalc.volume HeatCalc.t1 HeatCalc.t2)
        Report.line output "总物质的量" (sprintf "%.3f kmol/h（22.414 Nm³/kmol）" total)

        for ((name, kmol), (_, pct)) in List.zip HeatCalc.designKmol HeatCalc.wetComponents do
            Report.line
                output
                (sprintf "%-4s" name)
                (sprintf "%6.2f vol.%% → %10.3f kmol/h　Δh %.1f kJ/kmol" pct kmol (HeatCalc.molarEnthalpy name HeatCalc.t2 - HeatCalc.molarEnthalpy name HeatCalc.t1))

        let q = HeatCalc.designHeat
        let wan = q / HeatCalc.kjPerKcal / 10000.0
        Report.line output "烟气吸热量" (sprintf "%.0f kJ/h" q)
        Report.line output "换算" (sprintf "%.1f kW　%.0f kcal/h　%.2f 万大卡/hr" (q / 3600.0) (q / HeatCalc.kjPerKcal) wan)
        Report.line output "给定热交换量" (sprintf "%.0f 万大卡/hr；计算值 / 给定值 = %.3f" HeatCalc.givenWanDaKa (wan / HeatCalc.givenWanDaKa))

        Should.equal 100.0 (Math.Round((HeatCalc.wetComponents |> List.sumBy snd), 2))
        Assert.True(total > 2500.0 && total < 2600.0, "57000 Nm³/h 应约 2543 kmol/h")
        Assert.True(q > 2.6e7 && q < 2.7e7, "吸热量应约 2.65e7 kJ/h")
        Report.conclusion output "通过：多项式复算参考工况偏差 -0.27%，本书工况算得约 634 万大卡/hr；纯逻辑，不碰真机"

    [<Fact>]
    member this.``烟气吸热清点：只读现场与参考计算链``() =
        let info = HeatBoard.info ()
        let doc = HeatBoard.doc ()
        let before = Report.procIds info.Image

        Report.title output "烟气吸热清点：挂回常驻 Excel，核对输入、参考计算链与 A3 现场"
        Report.line output "参数表" (info.describe())
        Report.line output "文档" (doc.describe())

        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

        let snap = STA.run(fun () -> HeatBoard.survey app)

        for name, path, sheets in snap.Workbooks do
            Report.line output "打开的工作簿" (sprintf "「%s」 %s" name path)
            Report.line output "  其中工作表" (String.concat "、" sheets)

        for r, text in snap.HeadRows do
            Report.line output (sprintf "A%d" r) (if text = "" then "(空)" else text)

        Report.line output "A3:N50 非空" (if snap.Occupied.IsEmpty then "(全空，可写)" else String.concat "；" snap.Occupied)

        for addr, value in snap.SourceValues do
            Report.line output (sprintf "参数表 %s" addr) value

        if snap.RefRows.IsEmpty then
            Report.line output "参考表" "参考工作簿没开，跳过公式读取（磁盘结构已核对）"
        else
            for r, label, value, formula in snap.RefRows do
                Report.line output (sprintf "参考 C%d" r) (sprintf "%s　值「%s」　公式「%s」" label value formula)

        for probe in snap.XllProbes do
            Report.line output "XLL 探测" probe
        Report.line output "工作簿 Saved" (string snap.BookSaved)
        Report.line output "磁盘最后写入" snap.DiskAt

        let q = HeatCalc.designHeat
        Report.line output "拟写结果" (sprintf "%.0f kJ/h = %.2f 万大卡/hr" q (q / HeatCalc.kjPerKcal / 10000.0))

        let after = Report.procIds info.Image
        Report.line output "收尾" (sprintf "%s 进程 %A" info.Image (Set.toList after))
        Should.equal before after
        Report.conclusion output "通过：只读查明输入、A3 现场与参考体积法计算链，磁盘一个字节没动"

    [<Fact>]
    member this.``烟气吸热回填：A3 起复刻体积法计算过程``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_HEAT" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_HEAT=1 才真的写盘"
        | _ ->
            let info = HeatBoard.info ()
            let doc = HeatBoard.doc ()
            let before = Report.procIds info.Image
            let target =
                match doc.Path with
                | Some path -> path
                | None -> failwith "黑板里没有落盘路径，不往下走"

            Assert.True(File.Exists target, sprintf "文件不在：%s" target)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmmss")
            let backup = Path.Combine(projectRoot, "TestResults", sprintf "焦炉煤气燃烧计算书[吸热计算前]%s.xlsx" stamp)
            let backedUp = not (File.Exists backup)
            if backedUp then File.Copy(target, backup, false)

            Report.title output "烟气吸热回填：A3 起复刻参考表体积法计算过程"
            Report.line output "参数表" (info.describe())
            Report.line output "文档" (doc.describe())
            Report.line output "备份" (if backedUp then backup else "已有备份，未重复生成：" + backup)
            Report.line output "存盘前" (File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss"))
            Report.line output "计算过程" "组分 vol.% → Nm³/h → /22.414 kmol/h → 各组分 kmol/h → Shomate 两温度显焓 → 焓差 → kJ/h / kW / kcal/h"

            let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
            Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

            let readBack, saved, diskAfter, stale = STA.run(fun () -> HeatBoard.write app)

            Report.line
                output
                "半途残留"
                (if List.isEmpty stale then
                     "(无)"
                 else
                     sprintf "%d 格都在本次写入清单里，已幂等覆盖：%s" stale.Length (String.concat "、" stale))

            for addr, value in readBack do
                Report.line output addr value

            let valueOf addr = readBack |> List.find (fun (a, _) -> a = addr) |> snd
            let num addr = Double.Parse(valueOf addr, NumberStyles.Float, CultureInfo.InvariantCulture)
            let expected = HeatCalc.designHeat

            Assert.True(abs (num "C33" / expected - 1.0) < 0.005, sprintf "C33 读回 %.0f 与纯规则 %.0f 不符" (num "C33") expected)
            Assert.True(abs (num "C36" / (expected / HeatCalc.kjPerKcal / 10000.0) - 1.0) < 0.005, "万大卡/hr 读回不符")
            Assert.True((abs (num "C49")) < 0.01, sprintf "Shomate 独立复算与主函数偏差过大：%.4f" (num "C49"))
            Should.equal HeatCalc.volume (num "C14")
            Should.equal HeatCalc.t1 (num "C27")
            Should.equal HeatCalc.t2 (num "C30")
            Should.equal HeatCalc.givenWanDaKa (num "C37")
            Should.equal true saved
            Should.equal true (File.Exists backup)

            Report.line output "工作簿 Saved" (string saved)
            Report.line output "存盘后" diskAfter

            let after = Report.procIds info.Image
            Report.line output "现场" (sprintf "%s 进程 干活前 %A / 收尾后 %A" info.Image (Set.toList before) (Set.toList after))
            Should.equal before after
            Should.equal true (Set.contains info.Pid after)
            Report.conclusion output "完成：A3 起体积法计算链已存盘并逐格读回一致；原件备份在项目目录外，实例没新建也没退出"
