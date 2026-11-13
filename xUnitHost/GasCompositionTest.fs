/// 【被加热烟气成分计算】用「补热」那一套燃烧计量（仙桃湿空气 + XLL 数组函数）
/// 反算被加热烟气的组分。只算成分 —— 焓值、放热量、闭合校核一律不写。
///
/// α 由第一张表「烟气含氧量 8~10%」取中值 9%（干基）反解，闭式解：
///   干基 O2% = tO2·(α-1) / (K + α·M)
///   其中 K = 产物CO2 + 燃料N2 + 产物SO2 - tO2，M = tO2 + L0·aN2
///   → α = (x·K + tO2) / (tO2 - x·M)
///
/// 结果只落在新表「烟气成分计算」上。「需要加热的烟气参数」!B15:F24 一个字不动
/// （用户 2026-11-10 16:52 明确「不写」）。
///
/// 通道同前面几轮：按 PID 挂回现成实例，绝不新建 / Quit / kill；
/// COM 全在 OsHost.Sta.run 固定 STA 线程；写公式一律 Range.Formula2
/// （Range.Formula 会给自定义函数补 @ 隐式交集）。
///
/// 湿空气成分用 XLL 数组函数「空气体积成分」（2026-11-10 17:27 起挂回的 Excel pid=14052
/// 已正常加载该插件；上一轮 pid=28884 里解析不到只是那个进程没加载成功）。
/// C7 用 spill 公式，C8/C9 由它溢出；C10 另用 Magnus 内置公式独立复算，不依赖插件。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

/// 纯规则：焦炉煤气完全燃烧 + 仙桃湿空气，由目标干基含氧量反算 α 与烟气成分
module internal CompositionCalc =

    // 「焦炉煤气」!E2:E11 设计值（vol.%，干基；C2 以上不饱和烃按 C2H4 计，水分按干基忽略）
    let co = 6.5
    let h2 = 57.0
    let ch4 = 24.5
    let c2h4 = 2.5
    let co2 = 2.5
    let o2 = 0.55
    let n2 = 6.45
    let h2s = 0.0

    /// 仙桃大气（100760 Pa / 16.5℃ / 80%）下湿空气成分，由 Excel「空气体积成分」给出
    let airH2O = 1.4910678840809841
    let airO2 = 20.68687574434299
    let airN2 = 77.82205637157602

    /// 第一张表「烟气含氧量 8~10%」取中值（干基）
    let targetDryO2 = 9.0

    type Meter =
        { O2Need: float
          L0: float
          Alpha: float
          CoefK: float
          CoefM: float
          Air: float
          ProdH2O: float
          ProdO2: float
          ProdN2: float
          ProdCO2: float
          ProdSO2: float
          Wet: float
          Dry: float }

    let compute (target: float) : Meter =
        let o2Need = (0.5 * co + 0.5 * h2 + 2.0 * ch4 + 3.0 * c2h4 + 1.5 * h2s - o2) / 100.0
        let aO2 = airO2 / 100.0
        let aN2 = airN2 / 100.0
        let aH2O = airH2O / 100.0
        let l0 = o2Need / aO2
        let prodCO2 = (co + ch4 + 2.0 * c2h4 + co2) / 100.0
        let prodN2Fuel = n2 / 100.0
        let prodSO2 = h2s / 100.0
        let k = prodCO2 + prodN2Fuel + prodSO2 - o2Need
        let m = o2Need + l0 * aN2
        let x = target / 100.0
        let alpha = (x * k + o2Need) / (o2Need - x * m)
        let air = alpha * l0
        let prodH2O = (h2 + 2.0 * ch4 + 2.0 * c2h4 + h2s) / 100.0 + air * aH2O
        let prodO2 = o2Need * (alpha - 1.0)
        let prodN2 = prodN2Fuel + air * aN2
        let wet = prodH2O + prodO2 + prodN2 + prodCO2 + prodSO2
        let dry = prodO2 + prodN2 + prodCO2 + prodSO2

        { O2Need = o2Need
          L0 = l0
          Alpha = alpha
          CoefK = k
          CoefM = m
          Air = air
          ProdH2O = prodH2O
          ProdO2 = prodO2
          ProdN2 = prodN2
          ProdCO2 = prodCO2
          ProdSO2 = prodSO2
          Wet = wet
          Dry = dry }

    /// 湿基 / 干基 体积百分数（H2O 的干基按 0 计，与「需要加热的烟气参数」表口径一致）
    let wetPct (v: float) (m: Meter) = v / m.Wet * 100.0
    let dryPct (v: float) (m: Meter) = v / m.Dry * 100.0

module internal CompositionBoard =

    let info () : BoardApp = CombustionBoard.info ()
    let bookPath = CombustionBoard.targetPath
    let sheetName = "烟气成分计算"
    let sourceSheet = "需要加热的烟气参数"

    let pressurePa = CombustionBoard.pressurePa
    let airCelsius = CombustionBoard.airCelsius
    let humidityPct = CombustionBoard.humidityPct

    let private findBook (app: Application) (path: string) : Workbook option =
        let count = try app.Workbooks.Count with _ -> 0

        [ 1 .. count ]
        |> List.tryPick(fun i ->
            let book = app.Workbooks.Item(i)
            let full = try book.FullName with _ -> ""

            if String.Equals(full, path, StringComparison.OrdinalIgnoreCase) then
                Some book
            else
                None)

    let private findSheet (book: Workbook) (name: string) : Worksheet option =
        [ 1 .. book.Worksheets.Count ]
        |> List.tryPick(fun i ->
            let s = book.Worksheets.Item(i) :?> Worksheet
            if s.Name = name then Some s else None)

    /// 现场只读取数：新表是否已存在、目标书其余表名、B15:F24 现状（本轮不动它）
    type Snapshot =
        { Sheets: string list
          HasSheet: bool
          SourceRows: string list
          XllProbe: string list
          Neighbour: string }

    let survey (app: Application) : Snapshot =
        let book, weOpened =
            match findBook app bookPath with
            | Some b -> b, false
            | None -> ExcelHost.BookIO.openBook app bookPath true, true

        let sheets = [ for i in 1 .. book.Worksheets.Count -> (book.Worksheets.Item(i) :?> Worksheet).Name ]
        let has = sheets |> List.contains sheetName

        let sourceRows =
            match findSheet book sourceSheet with
            | None -> [ "（没有「需要加热的烟气参数」表）" ]
            | Some s ->
                [ for r in 15 .. 24 ->
                      let parts =
                          [ for c in [ "B"; "C"; "D"; "E" ] ->
                                let v = s.Range(sprintf "%s%d" c r).Value2
                                if isNull v then "" else Convert.ToString(v, CultureInfo.InvariantCulture) ]

                      sprintf "%d行 %s" r (String.concat " | " parts) ]

        // 自己开的只读副本收尾关掉，不往用户 Excel 里多留窗口
        if weOpened then
            ExcelHost.BookIO.closeBook book

        // 探测 XLL 数组函数在这只 Excel 里此刻能不能算（上一轮写盘报 0x800A07ED 就靠它定位）
        let xllProbe =
            let describe (v: obj) =
                match v with
                | null -> "(null)"
                | :? string as s -> s
                | :? int as i when i < -2000000000 -> sprintf "错误码 %d" i
                | _ -> Convert.ToString(v, CultureInfo.InvariantCulture)

            [ "空气体积成分(100760,16.5,80)"; "gasMoleEnthalpy({\"N2\"},{1},0)" ]
            |> List.map(fun formula ->
                try
                    sprintf "%s → %s" formula (describe (app.Evaluate(box formula)))
                with ex ->
                    sprintf "%s → 抛错：%s" formula ex.Message)

        // 同一本书里「补热燃烧计算」用同一批函数，看它现在是否还正常
        let neighbour =
            match findSheet book sheetName, findSheet book "补热燃烧计算" with
            | _, None -> "「补热燃烧计算」表不在"
            | _, Some s ->
                [ "C6"; "C35"; "C43" ]
                |> List.map(fun a ->
                    let v = s.Range(a).Value2
                    sprintf "%s=%s" a (if isNull v then "(空)" else Convert.ToString(v, CultureInfo.InvariantCulture)))
                |> String.concat "　"

        { Sheets = sheets
          HasSheet = has
          SourceRows = sourceRows
          XllProbe = xllProbe
          Neighbour = neighbour }

    /// 写新表：大气 → 煤气成分 → α 反算 → 实际烟气量 → 湿基/干基体积分数。只到成分为止。
    let write (app: Application) : (string * string) list * bool * string =
        let opened =
            match findBook app bookPath with
            | Some b -> b, false
            | None -> ExcelHost.BookIO.openBook app bookPath false, true

        let book, weOpened = opened

        let sheet =
            match findSheet book sheetName with
            | Some s -> s
            | None ->
                let s = book.Worksheets.Add() :?> Worksheet
                s.Name <- sheetName
                (try
                    s.Move(box Type.Missing, box (book.Worksheets.Item(book.Worksheets.Count)))
                 with _ ->
                    ())
                s

        // 先收集 (地址, 值)，最后统一落笔；公式一律 Formula2
        let writes = ResizeArray<string * obj>()
        let put (a: string) (v: obj) = writes.Add(a, v)

        let line r name unit value note =
            put (sprintf "A%d" r) (box name)
            if unit <> "" then put (sprintf "B%d" r) (box unit)
            if not (isNull value) then put (sprintf "C%d" r) value
            if note <> "" then put (sprintf "D%d" r) (box note)

        let f (x: string) = box x

        line 1 "名称" "单位" (f "数值") "备注"
        line 2 "一、输入：大气（湖北仙桃）与目标含氧量" "" null ""
        line 3 "大气压力" "Pa" (box pressurePa) "仙桃多年平均气压 1006.5 hPa；与「烟气吸热量计算」同口径"
        line 4 "大气温度" "℃" (box airCelsius) "仙桃多年平均气温 16.5~17.0℃"
        line 5 "大气相对湿度" "%" (box humidityPct) "仙桃多年平均相对湿度 79~80%"
        line 6 "目标烟气含氧量（干基）" "%" (box CompositionCalc.targetDryO2) "第一张表「烟气含氧量 8~10%」取中值 9%；改这里 α 与全表跟着变"
        line 7 "湿空气 H2O" "vol. %" (f "=空气体积成分($C$3,$C$4,$C$5)") "XLL 数组函数，spill 到 C7:C9（组分序 H2O/O2/N2）"
        line 8 "湿空气 O2" "vol. %" null ""
        line 9 "湿空气 N2" "vol. %" null ""
        line 10 "湿空气 H2O 复核（Magnus）" "%" (f "=C5/100*6.112*EXP(17.67*C4/(C4+243.5))*100/C3*100") "内置公式独立复算，与 C7 应一致"

        line 11 "二、焦炉煤气成分（活链接「焦炉煤气」设计值）" "" null "C2 以上不饱和烃按 C2H4 计"
        line 12 "CO" "vol. %" (f "='焦炉煤气'!E6") ""
        line 13 "H2" "vol. %" (f "='焦炉煤气'!E7") ""
        line 14 "CH4" "vol. %" (f "='焦炉煤气'!E8") ""
        line 15 "C2H4（C2 以上不饱和烃）" "vol. %" (f "='焦炉煤气'!E10") ""
        line 16 "CO2" "vol. %" (f "='焦炉煤气'!E5") ""
        line 17 "O2" "vol. %" (f "='焦炉煤气'!E9") ""
        line 18 "N2" "vol. %" (f "='焦炉煤气'!E4") ""
        line 19 "H2S" "vol. %" (box 0.0) "燃料未给 S，按 0"
        line 20 "合计" "vol. %" (f "=SUM(C12:C19)") "应为 100.00"

        line 22 "三、过量空气系数由含氧量反算（闭式解）" "" null "干基 O2% = tO2(α-1)/(K+α·M) → α = (x·K+tO2)/(tO2-x·M)"
        line 23 "理论需氧量 tO2" "Nm³/Nm³" (f "=(0.5*C12+0.5*C13+2*C14+3*C15+1.5*C19-C17)/100") "H2/CO 各 0.5 O2，CH4 2，C2H4 3，扣燃料自带 O2"
        line 24 "湿空气 O2 份额" "—" (f "=C8/100") ""
        line 25 "湿空气 N2 份额" "—" (f "=C9/100") ""
        line 26 "湿空气 H2O 份额" "—" (f "=C7/100") ""
        line 27 "理论空气量 L0" "Nm³/Nm³" (f "=C23/C24") ""
        line 28 "产物 CO2" "Nm³/Nm³" (f "=(C12+C14+2*C15+C16)/100") ""
        line 29 "产物 N2（燃料带入）" "Nm³/Nm³" (f "=C18/100") ""
        line 30 "产物 SO2" "Nm³/Nm³" (f "=C19/100") "污染物按 0"
        line 31 "系数 K" "Nm³/Nm³" (f "=C28+C29+C30-C23") "干烟气里与 α 无关的部分"
        line 32 "系数 M" "Nm³/Nm³" (f "=C23+C27*C25") "干烟气对 α 的斜率"
        line 33 "过量空气系数 α" "—" (f "=(C6/100*C31+C23)/(C23-C6/100*C32)") "由目标干基含氧量反算，不是取值"

        line 35 "四、实际烟气量（每 1 Nm³ 干煤气）" "" null ""
        line 36 "实际空气量" "Nm³/Nm³" (f "=C33*C27") ""
        line 37 "　其中 H2O" "Nm³/Nm³" (f "=C36*C26") "湿空气带入"
        line 38 "　其中 O2" "Nm³/Nm³" (f "=C36*C24") ""
        line 39 "　其中 N2" "Nm³/Nm³" (f "=C36*C25") ""
        line 40 "产物 H2O" "Nm³/Nm³" (f "=(C13+2*C14+2*C15+C19)/100+C37") "燃烧生成水 + 空气带入"
        line 41 "产物 O2" "Nm³/Nm³" (f "=C23*(C33-1)") ""
        line 42 "产物 N2" "Nm³/Nm³" (f "=C29+C39") ""
        line 43 "产物 CO2" "Nm³/Nm³" (f "=C28") ""
        line 44 "产物 SO2" "Nm³/Nm³" (f "=C30") "污染物按 0"
        line 45 "湿烟气合计" "Nm³/Nm³" (f "=SUM(C40:C44)") ""
        line 46 "干烟气合计" "Nm³/Nm³" (f "=SUM(C41:C44)") ""

        put "A48" (box "五、被加热烟气成分（结果，供焓值计算输入）")
        put "A49" (box "组分")
        put "B49" (box "单位")
        put "C49" (box "湿基 vol. %")
        put "D49" (box "干基 vol. %")
        put "E49" (box "备注")

        let row r name wet dry note =
            put (sprintf "A%d" r) (box name)
            put (sprintf "B%d" r) (box "vol. %")
            put (sprintf "C%d" r) (f wet)
            put (sprintf "D%d" r) (f dry)
            if note <> "" then put (sprintf "E%d" r) (box note)

        row 50 "H2O" "=C40/$C$45*100" "=0" "燃烧生成水 + 湿空气带入；干基按 0"
        row 51 "O2" "=C41/$C$45*100" "=C41/$C$46*100" "由 α 与燃料自带 O2 决定；干基＝第一张表「烟气含氧量」目标值"
        row 52 "N2" "=C42/$C$45*100" "=C42/$C$46*100" "燃料 N2 + 助燃空气 N2"
        row 53 "CO2" "=C43/$C$45*100" "=C43/$C$46*100" "完全燃烧生成"
        row 54 "SO2" "=C44/$C$45*100" "=C44/$C$46*100" "燃料未给 S，按 0"
        row 55 "SO3" "=0" "=0" "燃料未给 S，按 0"
        row 56 "HCl" "=0" "=0" "燃料未给 Cl，按 0"
        row 57 "HF" "=0" "=0" "燃料未给 F，按 0"
        row 58 "ash" "=0" "=0" "无灰分数据，按 0"
        row 59 "合计" "=SUM(C50:C58)" "=SUM(D50:D58)" "湿基、干基各应为 100"

        put "A61" (box "口径与依据")
        put "A62" (box "被加热烟气按「同种焦炉煤气完全燃烧产物」计，沿用「烟气吸热量计算」!A2 的假设")
        put "A63" (box "α 不取值，由第一张表「烟气含氧量 8~10%」中值 9%（干基）反算，闭式解见 C23:C33")
        put "A64" (box "助燃空气按湖北仙桃湿空气（C3:C5 → XLL「空气体积成分」溢出 C7:C9），含 H2O 约 1.49%，C10 用 Magnus 内置公式复核；若按干空气 21/79 计，湿基 H2O 为 14.47%")
        put "A65" (box "燃料 C2 以上不饱和烃按 C2H4 计；S/Cl/F/灰分无数据按 0；燃料水分按干基忽略")
        put "A66" (box "本表只算成分；焓值、放热量见「烟气吸热量计算」与 补热焦炉煤气燃烧计算书.xlsx")
        put "A68" (box "口径变更（2026-11-10）：原「需要加热的烟气参数」!B15:F24 那张手工成分块（干空气 21/79）数值已过时，已删除；本表是唯一口径，「烟气吸热量计算」C4:D12 直接引用本表 C50:D58。")
        put "A69" (box "旧干空气口径的湿基值仅供对照：H2O 14.47 / O2 7.70 / N2 72.81 / CO2 5.02（差只来自助燃空气是否含水）；干基两套完全一致 9.00 / 85.13 / 5.87。")

        // 幂等校验：重跑允许覆盖本 Fact 上次写下的内容，但绝不盖掉别人新写的东西
        let existing =
            let gridObj = sheet.Range("A1:E70").Value2

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

        let spillAnchors = [ "C7" ]
        let spillMembers: string list = [ "C8"; "C9" ]

        let allowed =
            Set.union (writes |> Seq.map fst |> Set.ofSeq) (Set.ofList (spillAnchors @ spillMembers))

        let stray = existing |> List.filter (fun a -> not (Set.contains a allowed))

        if not stray.IsEmpty then
            failwithf "「%s」里已有本次清单之外的非空单元格，停手：%s" sheetName (String.concat "、" stray)

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
            [ "C7"; "C8"; "C9"; "C10"; "C20"; "C23"; "C27"; "C33"; "C36"; "C40"; "C41"; "C42"
              "C43"; "C45"; "C46"; "C50"; "C51"; "C52"; "C53"; "D50"; "D51"; "D52"
              "D53"; "C59"; "D59" ]

        let readBack = addrs |> List.map (fun a -> a, cellText a)
        let saved = try book.Saved with _ -> true
        let sheetCount = book.Worksheets.Count

        if weOpened then
            ExcelHost.BookIO.closeBook book

        readBack, saved, sprintf "%d 张表；磁盘 %s（%d B）" sheetCount (File.GetLastWriteTime(bookPath).ToString("yyyy-MM-dd HH:mm:ss")) (FileInfo(bookPath).Length)

/// 【接线】把「烟气成分计算」扶正成唯一口径：
///   ① Excel 用 ExcelHost 的启动函数 `ExcelTask.getOrCreateOnSta()` 起 —— 它按 ROT 交回现成那一只，
///      XLL 插件随它一起加载好了（用户 2026-11-10 17:40 定的口径：别自己糊一只不带插件的实例）；
///   ② 「需要加热的烟气参数」!B15:F24 那张手工成分块（干空气 21/79 口径）数值已过时 —— 删掉；
///   ③ 「烟气吸热量计算」!C4:D12 直接引用「烟气成分计算」!C50:D58；
///      参数表里 α / L0 / 湿基产物量三行也从死值改成指向新表的活链接；
///   ④ 「烟气成分计算」表移到「烟气吸热量计算」之前（成分 → 吸热量，读的顺序）。
module internal CompositionWiring =

    let private bookPath = CompositionBoard.bookPath
    let paramSheet = "需要加热的烟气参数"
    let compSheet = "烟气成分计算"
    let heatSheet = "烟气吸热量计算"

    /// ExcelHost 的启动函数：本机有 Excel 就交回那一只（插件已加载），没有就新建一只再补空白簿
    let start () : Application = ExcelHost.ExcelTask.getOrCreateOnSta ()

    let describe (v: obj) =
        match v with
        | null -> "(空)"
        | :? string as s -> s
        | :? int as i when i < -2000000000 -> sprintf "错误码 %d" i
        | _ -> Convert.ToString(v, CultureInfo.InvariantCulture)

    /// 这只 Excel 此刻能不能算 XLL 的数组函数 —— 起完实例第一件事就验它
    let probe (app: Application) : string list =
        [ "空气体积成分(100760,16.5,80)"; "gasMoleEnthalpy({\"N2\"},{1},0)" ]
        |> List.map(fun formula ->
            try
                sprintf "%s → %s" formula (describe (app.Evaluate(box formula)))
            with ex ->
                sprintf "%s → 抛错：%s" formula ex.Message)

    let private findBook (app: Application) (path: string) : Workbook option =
        let count = try app.Workbooks.Count with _ -> 0

        [ 1 .. count ]
        |> List.tryPick(fun i ->
            let book = app.Workbooks.Item(i)
            let full = try book.FullName with _ -> ""

            if String.Equals(full, path, StringComparison.OrdinalIgnoreCase) then
                Some book
            else
                None)

    let private findSheet (book: Workbook) (name: string) : Worksheet option =
        [ 1 .. book.Worksheets.Count ]
        |> List.tryPick(fun i ->
            let s = book.Worksheets.Item(i) :?> Worksheet
            if s.Name = name then Some s else None)

    let order (book: Workbook) : string list =
        [ for i in 1 .. book.Worksheets.Count -> (book.Worksheets.Item(i) :?> Worksheet).Name ]

    /// 现场只读：表序、旧块 B15:F24、吸热量表 C4:D12 现在指向谁、参数表 α 三行、XLL 探测
    type Snapshot =
        { Pid: int
          Order: string list
          OldBlock: string list
          HeatInputs: string list
          AlphaRows: string list
          Xll: string list
          HeatResult: string }

    let survey (app: Application) : Snapshot =
        let book, opened =
            match findBook app bookPath with
            | Some b -> b, false
            | None -> ExcelHost.BookIO.openBook app bookPath true, true

        let param = findSheet book paramSheet
        let heat = findSheet book heatSheet

        let oldBlock =
            match param with
            | None -> [ "（没有「需要加热的烟气参数」表）" ]
            | Some s ->
                [ for r in 15 .. 24 ->
                      let parts =
                          [ for c in [ "B"; "C"; "D"; "E"; "F" ] -> describe (s.Range(sprintf "%s%d" c r).Value2) ]

                      sprintf "%2d 行 %s" r (String.concat " | " parts) ]

        let heatInputs =
            match heat with
            | None -> [ "（没有「烟气吸热量计算」表）" ]
            | Some s ->
                [ for r in 4 .. 12 ->
                      sprintf "%2d 行 C=%s　D=%s" r
                          (describe (s.Range(sprintf "C%d" r).Formula2))
                          (describe (s.Range(sprintf "D%d" r).Formula2)) ]

        let alphaRows =
            match param with
            | None -> []
            | Some s ->
                [ for r in 26 .. 29 ->
                      sprintf "%2d 行 %s / %s / D=%s / %s" r
                          (describe (s.Range(sprintf "A%d" r).Value2))
                          (describe (s.Range(sprintf "B%d" r).Value2))
                          (describe (s.Range(sprintf "D%d" r).Value2))
                          (describe (s.Range(sprintf "F%d" r).Value2)) ]

        let heatResult =
            match heat with
            | None -> "（没有吸热量表）"
            | Some s ->
                sprintf "C36=%s 万大卡/hr　C38=%s" (describe (s.Range("C36").Value2)) (describe (s.Range("C38").Value2))

        let pid = try ExcelHost.Host.pidOf app |> Option.defaultValue 0 with _ -> 0

        let snap =
            { Pid = pid
              Order = order book
              OldBlock = oldBlock
              HeatInputs = heatInputs
              AlphaRows = alphaRows
              Xll = probe app
              HeatResult = heatResult }

        if opened then
            ExcelHost.BookIO.closeBook book

        snap

    /// 落笔：① 吸热量表 C4:D12 改指成分表 ② 参数表旧块清空 + α 三行改活链接 ③ 表序前移 ④ 成分表口径说明改写
    let write (app: Application) : string * string list * (string * string) list =
        let book =
            match findBook app bookPath with
            | Some b -> b
            | None -> ExcelHost.BookIO.openBook app bookPath false

        let pick name =
            match findSheet book name with
            | Some s -> s
            | None -> failwithf "目标书里没有「%s」表" name

        let param = pick paramSheet
        let comp = pick compSheet
        let heat = pick heatSheet

        let heatBefore = describe (heat.Range("C36").Value2)

        // ① 吸热量表：9 行组分，湿基 C50:C58 / 干基 D50:D58 一一对应
        for i in 0 .. 8 do
            heat.Range(sprintf "C%d" (4 + i)).Formula2 <- box (sprintf "=%s!C%d" compSheet (50 + i))
            heat.Range(sprintf "D%d" (4 + i)).Formula2 <- box (sprintf "=%s!D%d" compSheet (50 + i))

        // ② 旧成分块整块（含格式）清掉；参数表只留三行活链接的汇总值
        param.Range("B15:F24").Clear() |> ignore

        param.Range("D26").Formula2 <- box (sprintf "=%s!C33" compSheet)
        param.Range("D27").Formula2 <- box (sprintf "=%s!C27" compSheet)
        param.Range("D28").Formula2 <- box (sprintf "=%s!C45" compSheet)
        let noteF27 = "H2/CO 各 0.5 O2，CH4 2，C2H4 3；助燃空气按仙桃湿空气（O2 20.687 / N2 77.822 / H2O 1.491%），见「烟气成分计算」C27"
        let noteF29 = "被加热烟气成分见「烟气成分计算」A48:D59（C 列湿基用于焓值、D 列干基用于对照）；本表不再自带成分表"
        param.Range("F27").Value2 <- box noteF27
        param.Range("F29").Value2 <- box noteF29

        // ③ 成分表挪到吸热量表之前
        (try
            comp.Move(box heat, box Type.Missing)
         with ex ->
            failwithf "移动「%s」表失败：%s" compSheet (Errors.explain ex))

        // ④ 成分表底下那段「与现表 D16:E24 的差异」跟着改，别让说明指向已删的块
        let note68 = "口径变更（2026-11-10）：原「需要加热的烟气参数」!B15:F24 那张手工成分块（干空气 21/79）已删除，本表是唯一口径；「烟气吸热量计算」C4:D12 直接引用本表 C50:D58。"
        let note69 = "旧干空气口径的湿基值仅供对照：H2O 14.47 / O2 7.70 / N2 72.81 / CO2 5.02；干基两套完全一致 9.00 / 85.13 / 5.87。"
        comp.Range("E50").Value2 <- box "燃烧生成水 + 湿空气带入；干基按 0"
        comp.Range("A68").Value2 <- box note68
        comp.Range("A69").Value2 <- box note69

        for i in 1 .. book.Worksheets.Count do
            (try
                (book.Worksheets.Item(i) :?> Worksheet).Calculate()
             with _ ->
                ())

        book.Save()

        let readBack =
            [ "参数表 B15", describe (param.Range("B15").Value2)
              "参数表 C15", describe (param.Range("C15").Value2)
              "参数表 F15", describe (param.Range("F15").Value2)
              "参数表 B16", describe (param.Range("B16").Value2)
              "参数表 D16", describe (param.Range("D16").Value2)
              "参数表 E24", describe (param.Range("E24").Value2)
              "参数表 D26 α", describe (param.Range("D26").Value2)
              "参数表 D27 L0", describe (param.Range("D27").Value2)
              "参数表 D28 湿烟气", describe (param.Range("D28").Value2)
              "吸热量 C4 公式", describe (heat.Range("C4").Formula2)
              "吸热量 D4 公式", describe (heat.Range("D4").Formula2)
              "吸热量 C4 H2O", describe (heat.Range("C4").Value2)
              "吸热量 C5 O2", describe (heat.Range("C5").Value2)
              "吸热量 C6 N2", describe (heat.Range("C6").Value2)
              "吸热量 C7 CO2", describe (heat.Range("C7").Value2)
              "吸热量 C12 ash", describe (heat.Range("C12").Value2)
              "吸热量 C36 万大卡/hr", describe (heat.Range("C36").Value2)
              "吸热量 C38 计算/给定", describe (heat.Range("C38").Value2)
              "成分 C50 H2O 湿基", describe (comp.Range("C50").Value2)
              "成分 D52 N2 干基", describe (comp.Range("D52").Value2) ]

        let summary = sprintf "改前 C36=%s 万大卡/hr；改后见下" heatBefore
        summary, order book, readBack

/// 只算成分：焦炉煤气完全燃烧 + 仙桃湿空气，α 由目标干基含氧量反算
type GasCompositionTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``成分反算规则：由含氧量反解 α 再算组分``() =
        let m = CompositionCalc.compute CompositionCalc.targetDryO2

        Report.title output "成分反算规则：α 由干基含氧量 9% 反解，再算湿基/干基组分"
        Report.line output "输入" "焦炉煤气 CO 6.5 / H2 57 / CH4 24.5 / C2H4 2.5 / CO2 2.5 / O2 0.55 / N2 6.45 vol.%"
        Report.line output "湿空气" (sprintf "H2O %.4f / O2 %.4f / N2 %.4f vol.%%" CompositionCalc.airH2O CompositionCalc.airO2 CompositionCalc.airN2)
        Report.line output "理论需氧 tO2" (sprintf "%.6f Nm³/Nm³" m.O2Need)
        Report.line output "理论空气量 L0" (sprintf "%.6f Nm³/Nm³" m.L0)
        Report.line output "系数 K / M" (sprintf "%.6f / %.6f" m.CoefK m.CoefM)
        Report.line output "反算 α" (sprintf "%.6f" m.Alpha)
        Report.line output "实际空气量" (sprintf "%.6f Nm³/Nm³" m.Air)
        Report.line output "湿烟气 / 干烟气" (sprintf "%.6f / %.6f Nm³/Nm³" m.Wet m.Dry)
        Report.line output "湿基 H2O" (sprintf "%.4f %%" (CompositionCalc.wetPct m.ProdH2O m))
        Report.line output "湿基 O2 / N2 / CO2" (sprintf "%.4f / %.4f / %.4f %%" (CompositionCalc.wetPct m.ProdO2 m) (CompositionCalc.wetPct m.ProdN2 m) (CompositionCalc.wetPct m.ProdCO2 m))
        Report.line output "干基 O2 / N2 / CO2" (sprintf "%.4f / %.4f / %.4f %%" (CompositionCalc.dryPct m.ProdO2 m) (CompositionCalc.dryPct m.ProdN2 m) (CompositionCalc.dryPct m.ProdCO2 m))

        // 干基含氧量必须回到目标 9%，这是反算的定义
        Assert.True(abs (CompositionCalc.dryPct m.ProdO2 m - 9.0) < 1e-9, "干基 O2 应等于目标 9%")
        // α 与前面几轮手算的 1.6732 一致
        Assert.True(abs (m.Alpha - 1.6732) < 0.001, sprintf "α 应约 1.6732，实际 %.6f" m.Alpha)
        // 干基三行应与「需要加热的烟气参数」现有手算值一致（9.00 / 85.13 / 5.87）
        Assert.True(abs (CompositionCalc.dryPct m.ProdN2 m - 85.13) < 0.01, "干基 N2 应约 85.13")
        Assert.True(abs (CompositionCalc.dryPct m.ProdCO2 m - 5.87) < 0.01, "干基 CO2 应约 5.87")
        // 湿基合计、干基合计各自 100
        let parts = [ m.ProdH2O; m.ProdO2; m.ProdN2; m.ProdCO2; m.ProdSO2 ]

        Should.equal 100.0 (Math.Round(parts |> List.sumBy (fun v -> CompositionCalc.wetPct v m), 6))
        Should.equal 100.0 (Math.Round((parts |> List.tail) |> List.sumBy (fun v -> CompositionCalc.dryPct v m), 6))

        Report.conclusion output "通过：α 1.6732 由含氧量 9% 反解，干基 9.00/85.13/5.87 与现表手算一致；湿基因计入空气含水升到 H2O 15.63%；纯逻辑不碰真机"

    [<Fact>]
    member this.``烟气成分清点：只读现场与目标表``() =
        let info = CompositionBoard.info ()
        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        let snap = STA.run(fun () -> CompositionBoard.survey app)

        Report.title output "烟气成分清点：只读「焦炉煤气燃烧计算书」现场"
        Report.line output "黑板" (info.describe ())
        Report.line output "目标书" CompositionBoard.bookPath
        Report.line output "现有表" (String.concat " / " snap.Sheets)
        Report.line output "新表是否已存在" (if snap.HasSheet then "已存在（重跑会幂等覆盖）" else "不存在（本轮新建）")
        Report.line output "「需要加热的烟气参数」B15:E24 现状" ""
        for r in snap.SourceRows do
            Report.line output "　" r
        Report.line output "XLL 探测" ""
        for r in snap.XllProbe do
            Report.line output "　" r
        Report.line output "同书「补热燃烧计算」" snap.Neighbour

        Assert.True(snap.Sheets |> List.contains "焦炉煤气", "目标书里应有「焦炉煤气」表")
        Assert.True(snap.Sheets |> List.contains "需要加热的烟气参数", "目标书里应有「需要加热的烟气参数」表")
        Report.conclusion output "通过：只读确认目标书表结构与组分块现状，磁盘一个字节没动"

    [<Fact>]
    member this.``烟气成分回填：新表只算到成分``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_COMPOSITION" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_COMPOSITION=1 才真的写盘"
        | _ ->
            let info = CompositionBoard.info ()
            let before = Report.procIds info.Image
            let path = CompositionBoard.bookPath
            Assert.True(File.Exists path, sprintf "目标书不在：%s" path)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let backupDir = Path.Combine(projectRoot, "TestResults")
            Directory.CreateDirectory backupDir |> ignore

            let stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            let backup = Path.Combine(backupDir, sprintf "焦炉煤气燃烧计算书[成分计算前]%s.xlsx" stamp)

            if not (File.Exists backup) then
                File.Copy(path, backup, false)

            let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
            let mtimeBefore = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss")
            let readBack, saved, summary = STA.run(fun () -> CompositionBoard.write app)

            Report.title output "烟气成分回填：新表「烟气成分计算」只算到成分"
            Report.line output "黑板" (info.describe ())
            Report.line output "目标书" path
            Report.line output "备份" backup
            Report.line output "存盘前" mtimeBefore
            Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)
            Report.line output "写入" summary

            for (a, v) in readBack do
                Report.line output (sprintf "读回 %s" a) v

            let value a = readBack |> List.find (fun (x, _) -> x = a) |> snd
            let number a = float (value a)

            let after = Report.procIds info.Image
            Report.line output "进程" (sprintf "EXCEL 干活前 %A / 收尾后 %A" (Set.toList before) (Set.toList after))

            Assert.True(File.Exists backup, "备份应存在")
            Assert.True(saved, "目标书应已保存")
            Assert.True(readBack |> List.forall (fun (_, v) -> not (v.Contains "错误码")), "读回值里不应有 Excel 错误码")
            Assert.True(abs (number "C20" - 100.0) < 0.01, "煤气成分合计应为 100")
            Assert.True(abs (number "C7" - CompositionCalc.airH2O) < 0.01, sprintf "湿空气 H2O 应约 %.4f，实际 %s" CompositionCalc.airH2O (value "C7"))
            Assert.True(abs (number "C10" - number "C7") < 0.02, sprintf "Magnus 复核应接近 XLL 的 C7：C10=%s C7=%s" (value "C10") (value "C7"))
            Assert.True(abs (number "C7" + number "C8" + number "C9" - 100.0) < 0.05, "湿空气三组分应合计 100")
            Assert.True(abs (number "C33" - 1.6732) < 0.001, sprintf "反算 α 应约 1.6732，实际 %s" (value "C33"))
            Assert.True(abs (number "D51" - 9.0) < 1e-6, sprintf "干基 O2 应等于目标 9%%，实际 %s" (value "D51"))
            Assert.True(abs (number "C59" - 100.0) < 0.01, "湿基合计应为 100")
            Assert.True(abs (number "D59" - 100.0) < 0.01, "干基合计应为 100")
            // 与纯规则一致（Excel 与 F# 走同一条闭式解）
            let p = CompositionCalc.compute CompositionCalc.targetDryO2
            Assert.True(abs (number "C50" - CompositionCalc.wetPct p.ProdH2O p) < 1e-6, "Excel 湿基 H2O 应与纯规则一致")
            Assert.True(abs (number "D52" - CompositionCalc.dryPct p.ProdN2 p) < 1e-6, "Excel 干基 N2 应与纯规则一致")

            Report.conclusion output "完成：新表已算并存盘，α 由含氧量反算、成分读回自洽；B15:F24 与其它表一个字未动"

    [<Fact>]
    member this.``烟气成分接线清点：启动函数起 Excel，旧块与引用现况``() =
        let snap = STA.run(fun () -> CompositionWiring.survey (CompositionWiring.start ()))

        Report.title output "烟气成分接线清点：用 ExcelHost 启动函数拿 Excel，只读现场"
        Report.line output "Excel pid" (string snap.Pid)
        Report.line output "表序" (String.concat " → " snap.Order)
        Report.line output "XLL 探测" ""

        for r in snap.Xll do
            Report.line output "　" r

        Report.line output "「需要加热的烟气参数」B15:F24" ""

        for r in snap.OldBlock do
            Report.line output "　" r

        Report.line output "α / L0 / 产物量三行" ""

        for r in snap.AlphaRows do
            Report.line output "　" r

        Report.line output "「烟气吸热量计算」C4:D12 现在指向" ""

        for r in snap.HeatInputs do
            Report.line output "　" r

        Report.line output "吸热量结果" snap.HeatResult

        Assert.True(snap.Pid > 0, "启动函数应交回一只真实的 Excel 进程")
        Assert.True(snap.Order |> List.contains CompositionWiring.paramSheet, "目标书里应有「需要加热的烟气参数」表")
        Assert.True(snap.Order |> List.contains CompositionWiring.compSheet, "目标书里应有「烟气成分计算」表")
        Assert.True(snap.Order |> List.contains CompositionWiring.heatSheet, "目标书里应有「烟气吸热量计算」表")

        // 启动函数交回的这只必须能算 XLL —— 正是上一轮踩过的坑
        let bad = snap.Xll |> List.filter (fun s -> s.Contains "错误码" || s.Contains "抛错")
        Assert.True(bad.IsEmpty, sprintf "启动函数交回的 Excel 应已加载 XLL：%s" (String.concat "；" bad))

        Report.conclusion output "通过：启动函数交回的 Excel 能算 XLL；旧块与引用现状已记录，磁盘一个字节没动"

    [<Fact>]
    member this.``烟气成分接线：删旧块、吸热量改指新表、成分表前移``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_WIRING" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_WIRING=1 才真的写盘"
        | _ ->
            let path = CompositionBoard.bookPath
            Assert.True(File.Exists path, sprintf "目标书不在：%s" path)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let backupDir = Path.Combine(projectRoot, "TestResults")
            Directory.CreateDirectory backupDir |> ignore

            let stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            let backup = Path.Combine(backupDir, sprintf "焦炉煤气燃烧计算书[接线前]%s.xlsx" stamp)

            if not (File.Exists backup) then
                File.Copy(path, backup, false)

            let mtimeBefore = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss")

            let summary, order, readBack =
                STA.run(fun () ->
                    let app = CompositionWiring.start ()
                    let probe = CompositionWiring.probe app
                    let bad = probe |> List.filter (fun s -> s.Contains "错误码" || s.Contains "抛错")

                    if not bad.IsEmpty then
                        failwithf "启动函数交回的 Excel 没加载 XLL，停手：%s" (String.concat "；" bad)

                    CompositionWiring.write app)

            let value key =
                readBack |> List.tryFind (fun (k, _) -> k = key) |> Option.map snd |> Option.defaultValue "(缺)"

            let number key = float (value key)

            Report.title output "烟气成分接线：旧块下线，吸热量表改指「烟气成分计算」，成分表前移"
            Report.line output "目标书" path
            Report.line output "备份" backup
            Report.line output "存盘前" mtimeBefore
            Report.line output "写入" summary
            Report.line output "表序" (String.concat " → " order)

            for (k, v) in readBack do
                Report.line output k v

            let idx name = order |> List.findIndex (fun s -> s = name)
            let posComp, posHeat = idx CompositionWiring.compSheet, idx CompositionWiring.heatSheet

            Assert.True(posComp + 1 = posHeat, sprintf "「烟气成分计算」应紧排在「烟气吸热量计算」之前，实际 %d / %d" posComp posHeat)
            Assert.True(value "参数表 B15" = "(空)", "参数表 B15 的「组分/参数」表头应已删除")
            Assert.True(value "参数表 D16" = "(空)", "参数表 D16（旧湿基 H2O）应已删除")
            Assert.True(value "参数表 E24" = "(空)", "参数表 E24（旧块最后一格）应已删除")
            Assert.True((value "吸热量 C4 公式").Contains CompositionWiring.compSheet, "吸热量表 C4 应改为引用「烟气成分计算」")
            Assert.True(abs (number "参数表 D26 α" - 1.6732) < 0.001, sprintf "参数表 α 应为活链接 1.6732，实际 %s" (value "参数表 D26 α"))
            Assert.True(abs (number "参数表 D27 L0" - 4.2394) < 0.001, sprintf "L0 应约 4.2394，实际 %s" (value "参数表 D27 L0"))
            Assert.True(abs (number "参数表 D28 湿烟气" - 7.7760) < 0.001, sprintf "湿烟气应约 7.7760，实际 %s" (value "参数表 D28 湿烟气"))
            Assert.True(abs (number "吸热量 C4 H2O" - 15.6349) < 0.01, sprintf "湿基 H2O 应约 15.635，实际 %s" (value "吸热量 C4 H2O"))
            Assert.True(abs (number "吸热量 C5 O2" - 7.5929) < 0.01, sprintf "湿基 O2 应约 7.593，实际 %s" (value "吸热量 C5 O2"))
            Assert.True(abs (number "吸热量 C6 N2" - 71.8211) < 0.01, sprintf "湿基 N2 应约 71.821，实际 %s" (value "吸热量 C6 N2"))
            Assert.True(abs (number "吸热量 C7 CO2" - 4.9511) < 0.01, sprintf "湿基 CO2 应约 4.951，实际 %s" (value "吸热量 C7 CO2"))
            Assert.True(number "吸热量 C12 ash" = 0.0, "ash 仍应按 0")
            Assert.True(readBack |> List.forall (fun (_, v) -> not (v.Contains "错误码")), "读回值里不应有 Excel 错误码")
            Assert.True(abs (number "吸热量 C38 计算/给定" - number "吸热量 C36 万大卡/hr" / 530.0) < 0.01, "计算值/给定值应仍等于 C36/530")

            Report.conclusion output "完成：旧成分块已删，吸热量表改吃「烟气成分计算」的湿基/干基，成分表移到吸热量表之前；α 1.6732、湿基 H2O 15.63% 读回自洽"
