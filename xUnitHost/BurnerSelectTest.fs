/// 在「焦炉煤气燃烧计算书.xlsx」里新建「燃烧器选型」表。
///
/// 参照模板：河北浩锐陶瓷「天然气燃烧器」那一页的选型骨架 ——
///   燃料温度 / 理论空气量 / α / 热值 / 额定功率 → 额定耗量 → 标态到常温修正
///   → 煤气管道直径与流速 / 助燃风管道与套管环隙流速。
///
/// 与模板的差别（都是为了跟补燃表口径不打架）：
///   1. 燃料侧参数全部活链接「补热燃烧计算」（LHV、α、理论空气量、净放热、吸热量），
///      成分或 α 一改，本页自动跟着变，不另立第二套口径。
///   2. 模板用「热风炉热效率 0.95」把 LHV 折成有效放热；本表 C84 已经是
///      「LHV − 产物在终温的显热」，本身不含炉体散热，不再乘一次效率
///      （乘了会与补燃表 C86 的煤气耗量对不上）。裕量统一由 C28 选型系数覆盖。
///   3. 额定热负荷按「燃料输入热功率（LHV 计）」定义，因为低热值煤气选燃烧器
///      看的是燃料输入热量；净放热另列一行，方便对照补燃表。
///
/// 表结构（A 名称 / B 单位 / C 数值 / D 备注），六节：
///   一、工况与燃料物性   二、设计煤气量与助燃空气量   三、燃烧器额定参数
///   四、标态→工况换算   五、管径与流速校核          六、选型结论与制造要求
///
/// 改的是用户正开着的那本真表，所以：
///   磁盘备份 → 按 PID 挂回 → 有旧表就先删 → 新建到末尾 → 逐格落笔
///   → 重算 → 逐格读回断言 → Save。不新建实例、不 Quit、不碰别的工作簿。
///
/// 跑法：
///   XUNITHOST_EXCEL_PID=11468 XUNITHOST_BURNER_SELECT=1 \
///     xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*燃烧器选型表*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal SelectBoard =

    let targetPath = @"C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx"
    let sheetName = "燃烧器选型"
    let source = "'补热燃烧计算'"

    let txt (v: obj) =
        if isNull v then "" else Convert.ToString(v, CultureInfo.InvariantCulture)

    let num (v: obj) : float option =
        match v with
        | null -> None
        | :? double as d when not (Double.IsNaN d) -> Some d
        | :? int as i when i > -2000000000 && i < 2000000000 -> Some(float i)
        | :? decimal as m -> Some(float m)
        | :? single as s -> Some(float s)
        | _ -> None

    /// 相对差
    let relDiff (a: float) (b: float) =
        if a = 0.0 then abs b else abs (a - b) / abs a

    /// 圆管内径 → 流通面积 m²
    let bore (mm: float) = Math.PI / 4.0 * Math.Pow(mm / 1000.0, 2)

    /// 套管环隙流通面积 m²
    let annulus (outer: float) (inner: float) =
        Math.PI / 4.0 * (Math.Pow(outer / 1000.0, 2) - Math.Pow(inner / 1000.0, 2))

    /// 体积流量 m³/h（工况）÷ 面积 → 流速 m/s
    let velocity (flow: float) (area: float) = flow / 3600.0 / area

    /// 标态→工况温度修正
    let korr (t: float) = (273.15 + t) / 273.15

/// C 列的内容：公式 / 数值 / 纯文本
type internal Cell =
    | Formula of string
    | Number of float
    | Text of string

/// 一行：地址前缀由行号拼，A 名称、B 单位、C 内容、D 备注
module internal SelectRows =

    let private f (s: string) = Formula s
    let private v (x: float) = Number x
    let private t (s: string) = Text s

    /// (行号, A 名称, B 单位, C 内容, D 备注)
    let all: (int * string * string * Cell * string) list =
        [ // ---------- 表头 ----------
          1, "名称", "单位", t "数值", "备注"
          // ---------- 一、工况与燃料物性 ----------
          2, "一、工况与燃料物性（带引号的一律活链接「补热燃烧计算」，改源头自动跟着变）", "", t "", ""
          3, "大气温度", "℃", f "='补热燃烧计算'!C4", "标态→工况换算用；与补燃表同一个输入格"
          4, "焦炉煤气温度", "℃", v 20.0, "进燃烧器前的燃料温度，可改"
          5, "焦炉煤气低位热值", "kcal/Nm³", f "='补热燃烧计算'!C21", "活链接「焦炉煤气」!E3 设计值"
          6, "焦炉煤气低位热值", "kJ/Nm³", f "=C5*4.1868", "1 kcal = 4.1868 kJ"
          7, "焦炉煤气平均分子量", "g/mol", f "=('补热燃烧计算'!C12*28+'补热燃烧计算'!C13*2+'补热燃烧计算'!C14*16+'补热燃烧计算'!C15*28+'补热燃烧计算'!C16*44+'补热燃烧计算'!C17*32+'补热燃烧计算'!C18*28)/100", "按成分加权：CO 28 / H2 2 / CH4 16 / C2H4 28 / CO2 44 / O2 32 / N2 28"
          8, "焦炉煤气密度", "kg/Nm³", f "=C7/22.414", "标态密度，供厂家核对喷嘴动量"
          9, "理论空气量", "Nm³/Nm³煤气", f "='补热燃烧计算'!C38*22.414/100", "由煤气成分算出（湿空气，含 H2O）；成分一改自动变"
          10, "过量空气系数 α", "—", f "='补热燃烧计算'!C23", "补燃自身助燃空气的过量系数，可改补燃表 C23"
          11, "实际空气量系数", "Nm³/Nm³煤气", f "=C9*C10", "即理论空气量乘 α"
          12, "补燃烟气吸热量", "kJ/h", f "='补热燃烧计算'!C26", "被加热烟气 175→500℃ 的焓差，即补燃负荷"
          13, "补燃烟气吸热量", "万大卡/h", f "=C12/4.1868/10000", "应等于补燃表 C25 的 637.45"
          14, "每 Nm³ 煤气净放热", "kJ/Nm³", f "='补热燃烧计算'!C84", "LHV 减产物在终温 500℃ 的显热"
          15, "每 Nm³ 煤气净放热", "万大卡/Nm³", f "=C14/4.1868/10000", "对照用：约 0.3046"
          // ---------- 二、设计量 ----------
          17, "二、设计煤气量与助燃空气量（应分别等于补燃表 C86、C89）", "", t "", ""
          18, "设计煤气耗量", "Nm³/h", f "=C12/C14", "净放热口径：吸热量 ÷ 每 Nm³ 净放热"
          19, "与补燃表 C86 偏差", "%", f "=(C18-'补热燃烧计算'!C86)/'补热燃烧计算'!C86*100", "应 ≈ 0；两条路径同源，用来互相校核"
          20, "设计助燃空气量", "Nm³/h", f "=C18*C11", "α=1.15、湿空气"
          21, "与补燃表 C89 偏差", "%", f "=(C20-'补热燃烧计算'!C89)/'补热燃烧计算'!C89*100", "应 ≈ 0"
          22, "设计煤气热量（按 LHV）", "kcal/h", f "=C18*C5", "燃料输入热功率，燃烧器选型看的是这一行"
          23, "设计煤气热量", "万大卡/h", f "=C22/10000", "约 837"
          24, "设计煤气热量", "kW", f "=C22*4.1868/3600", "1 kcal = 4.1868 kJ，1 h = 3600 s"
          25, "设计煤气热量", "MW", f "=C24/1000", "约 9.74"
          // ---------- 三、额定参数 ----------
          27, "三、燃烧器额定参数（选型）", "", t "", ""
          28, "选型裕量系数", "—", v 1.1, "建议 1.10~1.20；炉体散热、积灰、煤气热值波动一并覆盖，不另乘效率"
          29, "燃烧器额定热负荷", "kcal/h", f "=C22*C28", "设计热功率 × 裕量"
          30, "燃烧器额定热负荷", "万大卡/h", f "=C29/10000", "向上取整到厂家成品规格（950 或 1000）"
          31, "燃烧器额定热负荷", "kJ/h", f "=C29*4.1868", ""
          32, "燃烧器额定热负荷", "kW", f "=C31/3600", ""
          33, "燃烧器额定热负荷", "MW", f "=C32/1000", "对外报价与比价用这一行"
          34, "额定负荷率（设计/额定）", "%", f "=C22/C29*100", "设计点落在 80~95% 为宜"
          35, "额定煤气耗量", "Nm³/h", f "=C29/C5", "额定热负荷 ÷ LHV"
          36, "额定助燃空气量", "Nm³/h", f "=C35*C11", "α=1.15 时"
          // ---------- 四、标态→工况 ----------
          38, "四、标态→工况换算", "", t "", ""
          39, "煤气温度修正系数", "—", f "=(273.15+C4)/273.15", "(273.15+t)/273.15"
          40, "助燃风温度修正系数", "—", f "=(273.15+C3)/273.15", "助燃风取大气温度"
          41, "额定工况煤气量", "m³/h", f "=C35*C39", "选管径用这个值，不是标态值"
          42, "额定工况助燃风量", "m³/h", f "=C36*C40", "同上"
          // ---------- 五、管径与流速 ----------
          44, "五、管径与流速校核（C45/C47/C49/C50 是可改输入格，看右边流速是否落在经验区间）", "", t "", ""
          45, "煤气支管内径", "mm", v 260.0, "DN250（φ273×6.5），从煤气总管到燃烧器"
          46, "煤气支管流速", "m/s", f "=C41/3600/(PI()/4*(C45/1000)^2)", "供气支管 10~15 m/s；焦炉煤气体积流量大，管径不能照天然气缩"
          47, "助燃风总管内径", "mm", v 600.0, "DN600（φ610×6）"
          48, "助燃风总管流速", "m/s", f "=C42/3600/(PI()/4*(C47/1000)^2)", "10~15 m/s"
          49, "燃烧器外管内径", "mm", v 393.7, "DN400（φ406.4×6.35），套管式燃烧器外壳"
          50, "燃烧器内管外径（煤气）", "mm", v 273.0, "φ273，即 DN250 的外径"
          51, "套管环隙流通面积", "m²", f "=PI()/4*((C49/1000)^2-(C50/1000)^2)", "外管内径围的面积减掉内管外径"
          52, "套管环隙流速", "m/s", f "=C42/3600/C51", "环隙 40~60 m/s，兼作助燃风出口速度，防回火与保证混合"
          // ---------- 六、结论与要求 ----------
          54, "六、选型结论与制造要求（文本，不参与计算）", "", t "", ""
          55, "选型结论", "", t "", "单台套管式焦炉煤气燃烧器（内管煤气、外管环隙助燃风），配变频助燃风机；热负荷、耗量、管径见 C29~C52，任一输入改动本页自动更新"
          56, "燃烧器形式", "", t "", "套管式；若改分体式，助燃风按 C36 提给厂家，环隙流速那条可作废"
          57, "燃料侧要点", "", t "", "焦炉煤气 H2 约 57%，火焰传播速度快，喷口出口速度必须高于回火临界速度；建议多孔分级喷头或旋流稳焰，并设回火防止器"
          58, "煤气净化", "", t "", "含萘、焦油与粉尘，入口设过滤器并蒸汽伴热，防萘结晶堵塞喷孔；管道低点设排凝"
          59, "空气侧要点", "", t "", "α 可调 1.05~1.25；低负荷时上调 α 防回火；风量按额定值的 110% 选风机"
          60, "调节比", "", t "", "建议 1:5（20%~100% 额定）；低负荷段保不住喷口速度时，掺烟气再循环或改用值班烧嘴"
          61, "材质", "", t "", "喷头与火焰筒 310S/2520 耐热钢；外壳 Q235B + 耐火浇注料；500℃ 混合段按 600℃ 选材留裕量"
          62, "接口", "", t "", "煤气接管 DN250（PN16，法兰 HG/T 20592）；助燃风接管 DN600；点火与火焰检测按厂家配套"
          63, "口径说明", "", t "", "本表不含炉体散热与积灰修正，已在 C28 选型裕量中一并覆盖；不再单乘热效率，避免与补燃表 C86 的耗量口径重复计损"
          64, "输入格", "", t "", "本页只有 C4（煤气温度）、C28（裕量）、C45/C47/C49/C50（四个管径）可改；其余全是公式或活链接" ]

    /// 67 行起：助燃风机选型。这是用户手搭的骨架（理论流量 → 流量裕量 → 实际流量 → 压升/效率
    /// → 轴功率 → 电机功率 → 安全裕量），只在「表不存在要新建」时才写。
    /// C76 电机功率接的是 XLL 数组函数 `_xll.电机额定功率(C75)`（按轴功率取标准电机档），
    /// 重建时故意留空、C77 也留空，避免把用户的 XLL 接线冲掉；表已存在时这两格原样保留。
    let skeleton: (int * string * string * Cell * string) list =
        [ 67, "风机选型计算", "", t "", "助燃风机：流量按额定助燃风量计夏季温度与裕量；压升按系统阻力分项累加后取用"
          68, "理论流量", "m3/h", f "=C36*(273.15+C71)/273.15", "标态助燃风量（C36，α=1.15）乘夏季温度修正；这一格原来是空的，整条链断在除零上"
          69, "实际流量", "m3/h", f "=ROUNDUP(E69,-2)", "选型风量 = 理论 × 流量裕量，向上取整到百位；要锁定某台样本风量，把这一格直接改成数字"
          70, "流量裕量", "", f "=C69/C68", "回算出来的实际裕量（实际/理论）"
          71, "进风温度", "℃", v 35.0, "仙桃夏季通风极端温度；温度越高风量越大，按夏季选才够"
          72, "实际流量", "m3/s", f "=C69/3600", "功率公式用 SI 单位"
          73, "压升", "kPa", v 4.5, "取用值；下方 91 行给分项估算合计，92 行自动对差"
          74, "效率", "%", v 60.0, "全压效率。60% 偏保守，厂家样本给 70~78% 时改这里，轴功率与电机功率一起降"
          75, "轴功率", "kW", f "=C72*C73/C74*100", "Q(m3/s) × ΔP(kPa) ÷ 效率"
          76, "电机功率", "kW", t "", "原格是 XLL 数组函数「电机额定功率(C75)」，按轴功率取标准电机档；重建表后需重新接上或手填"
          77, "电机安全裕量", "", t "", "= C76 / C75，1.1~1.2 为正常" ]

    /// 78 行起：风压分项与结论。每次跑都重写。
    let fan: (int * string * string * Cell * string) list =
        [ 79, "一、风压分项（按选型风量 C69 与总管 C47 估算，20℃、1.2 kg/m³ 冷态当量）", "", t "", "与风机样本同口径；高温时系统阻力与风机全压同比例下降，不必再乘密度系数"
          80, "管内流速", "m/s", f "=C69/3600/(PI()/4*(C47/1000)^2)", "按选型风量算，比第五节的 11.69 高——那是额定工况值，这里是含裕量的上限"
          81, "动压", "Pa", f "=0.5*1.2*C80^2", "½ρv²，ρ 取 1.2"
          82, "管道长度", "m", v 30.0, "输入：风机房到燃烧器的展开长度"
          83, "摩擦阻力系数 λ", "—", v 0.02, "钢管、0.1 mm 粗糙度、Re>4e5 的常见值"
          84, "沿程阻力", "Pa", f "=C83/(C47/1000)*C81*C82", "λ/d × 动压 × 管长"
          85, "局部阻力系数合计", "—", v 8.0, "弯头、阀门、变径、进出口之和，按管路图填"
          86, "局部阻力", "Pa", f "=C85*C81", "Σζ × 动压"
          87, "燃烧器本体阻力", "Pa", v 2500.0, "必须由厂家给；套管式燃气燃烧器常见 2~4 kPa，2500 只是初估"
          88, "烟道内背压", "Pa", v 200.0, "补燃点烟道一般微负压，此处按 +200 保守取"
          89, "其他（消音器/滤网）", "Pa", v 300.0, "视是否设置而定"
          90, "分项合计", "Pa", f "=C84+C86+C87+C88+C89", "分项累加"
          91, "折算", "kPa", f "=C90/1000", "与 C73 对差"
          92, "与 C73 之差", "Pa", f "=C90-C73*1000", "绝对值在 10% 以内即可；超出就核分项，或按实际改 C73"
          94, "二、选型结论与待确认（文本，不参与计算）", "", t "", ""
          95, "风机形式", "", t "", "高压离心通风机（9-19 / 9-26 系列），直联 2900 r/min；万级风量配 4.5 kPa 正是这一档的典型参数"
          96, "选型参数", "", t "", "按本页 C69 风量、C73 压升、C71 进风温度、C76 电机功率提给厂家"
          97, "机型与转速", "", t "", "可对应 9-19 系列 No.8D 一档（叶轮 φ800、2900 r/min）；机号与性能点须按厂家样本核对，同风量下全压不得低于 C73"
          98, "调节方式", "", t "", "变频调速，按 α（1.05~1.25）与负荷（20%~100%）调；30~50 Hz，节能按转速立方"
          99, "接口与附件", "", t "", "出口与助燃风总管同径（见 C47）；入口设滤网与消音器；室外布置时电机防护 IP55"
          100, "噪声", "", t "", "距风机 1 m 处 ≤85 dB(A)；不达标加消音器与减振基础"
          101, "待确认项", "", t "", "C87 燃烧器本体阻力必须由厂家给（2500 是初估）；C82 管长、C83 λ、C85 Σζ 按管路图核；C74 效率按样本改"
          102, "口径说明", "", t "", "流量裕量 1.3 偏保守（常见 1.1~1.2）；若厂家样本风量已含裕量，把 E70 降到 1.1。C80 的 16 m/s 是选型上限，长期满负荷运行应把总管升到 DN700" ]

    /// 按行定义落笔：A 名称 / B 单位 / C 内容 / D 备注
    let write (sheet: Worksheet) (rows: (int * string * string * Cell * string) list) =
        for (r, label, unit, cell, note) in rows do
            if label <> "" then sheet.Range(sprintf "A%d" r).Value2 <- box label
            if unit <> "" then sheet.Range(sprintf "B%d" r).Value2 <- box unit

            match cell with
            | Formula s -> sheet.Range(sprintf "C%d" r).Formula2 <- box s
            | Number x -> sheet.Range(sprintf "C%d" r).Value2 <- box x
            | Text s -> if s <> "" then sheet.Range(sprintf "C%d" r).Value2 <- box s

            if note <> "" then sheet.Range(sprintf "D%d" r).Value2 <- box note

type BurnerSelectTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``燃烧器选型表 按补燃负荷新建并接线``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_BURNER_SELECT" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_BURNER_SELECT=1 才改真表（会新建工作表并 Save）"
        | _ ->
            let pidText = Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID"

            if String.IsNullOrWhiteSpace pidText then
                Assert.Skip "没设 XUNITHOST_EXCEL_PID：这条要挂在活着的 Excel 上"

            let pid = int pidText
            let target = SelectBoard.targetPath
            Assert.True(File.Exists target, sprintf "表不在：%s" target)

            Report.title output "燃烧器选型　按补燃负荷新建一张表并接到「补热燃烧计算」"

            // ---------- 1. 磁盘备份 ----------
            let projectRoot =
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))

            let stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            let backupDir = Path.Combine(projectRoot, "TestResults")
            Directory.CreateDirectory backupDir |> ignore

            let backup =
                Path.Combine(backupDir, sprintf "焦炉煤气燃烧计算书[建选型表前]%s.xlsx" stamp)

            File.Copy(target, backup, false)
            Report.line output "磁盘备份" (sprintf "%s（%d 字节）" backup (FileInfo backup).Length)
            Report.line output "挂回" (sprintf "PID %d（只附着，不新建、不 Quit）" pid)

            let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await
            let back = STA.run (fun () -> ExcelHost.Host.pidOf app)
            Should.equal (Some pid) back

            let result =
                STA.run (fun () ->
                    let book =
                        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                        |> List.map (fun i -> app.Workbooks.Item(i))
                        |> List.tryFind (fun b ->
                            let full = try b.FullName with _ -> ""
                            String.Equals(full, target, StringComparison.OrdinalIgnoreCase))

                    let book =
                        match book with
                        | Some b -> b
                        | None -> failwithf "这只 Excel 里没打开「%s」" target

                    let names () =
                        [ 1 .. (try book.Worksheets.Count with _ -> 0) ]
                        |> List.map (fun i -> (book.Worksheets.Item(i) :?> Worksheet).Name)

                    let before = names ()

                    // 源表的三个关键量先读一遍：接线接得对不对，拿它们当基准
                    let srcSheet (n: string) =
                        [ 1 .. book.Worksheets.Count ]
                        |> List.map (fun i -> book.Worksheets.Item(i) :?> Worksheet)
                        |> List.find (fun s -> s.Name = n)

                    let src = srcSheet "补热燃烧计算"
                    let srcNum (a: string) = SelectBoard.num (src.Range(a).Value2)

                    let srcGas = srcNum "C86"
                    let srcAir = srcNum "C89"
                    let srcHeat = srcNum "C26"

                    let srcGas =
                        match srcGas with
                        | Some v -> v
                        | None -> failwith "补热表 C86（煤气耗量）不是数，现场不对"

                    let srcAir =
                        match srcAir with
                        | Some v -> v
                        | None -> failwith "补热表 C89（助燃空气量）不是数"

                    let srcHeat =
                        match srcHeat with
                        | Some v -> v
                        | None -> failwith "补热表 C26（吸热量）不是数"

                    // ---------- 2. 表已在就就地写（绝不重建：会把用户改过的输入格和 C76 的 XLL 接线冲掉） ----------
                    let alertsWas = app.DisplayAlerts
                    app.DisplayAlerts <- false

                    try
                        let existing =
                            [ 1 .. (try book.Worksheets.Count with _ -> 0) ]
                            |> List.map (fun i -> book.Worksheets.Item(i) :?> Worksheet)
                            |> List.tryFind (fun s -> s.Name = SelectBoard.sheetName)

                        let created, built =
                            match existing with
                            | Some ws -> ws, false
                            | None ->
                                let last =
                                    book.Worksheets.Item(try book.Worksheets.Count with _ -> 1) :?> Worksheet

                                let ws = book.Worksheets.Add() :?> Worksheet
                                ws.Move(Type.Missing, box last) |> ignore
                                ws.Name <- SelectBoard.sheetName
                                ws, true

                        // ---------- 3. 逐格落笔 ----------
                        // 只有表是新建的才写 1~64 行与风机骨架；表已存在就一个都不碰
                        if built then
                            SelectRows.write created SelectRows.all
                            SelectRows.write created SelectRows.skeleton
                            created.Range("E69").Formula2 <- box "=C68*E70"
                            created.Range("F69").Formula2 <- box "=E69-C69"
                            created.Range("E70").Value2 <- box 1.3
                        else
                            // 已有表的骨架只刷三处：理论流量（原为空）、选型流量（原为手填 9600）、进风温度
                            SelectRows.write
                                created
                                (SelectRows.skeleton
                                 |> List.filter (fun (r, _, _, _, _) -> r = 68 || r = 69 || r = 71))

                        // 78 行起每次都重写（分项与结论，纯派生内容）
                        created.Range("A78:F200").ClearContents() |> ignore
                        SelectRows.write created SelectRows.fan

                        app.CalculateFull()

                        // ---------- 4. 读回 ----------
                        let nv (a: string) = SelectBoard.num (created.Range(a).Value2)

                        let need (a: string) (what: string) =
                            match nv a with
                            | Some v -> v
                            | None -> failwithf "%s（%s）算出来不是数" a what

                        let gas = need "C18" "设计煤气耗量"
                        let air = need "C20" "设计助燃空气量"
                        let heatKcal = need "C22" "设计煤气热量"
                        let ratedKcal = need "C29" "额定热负荷"
                        let rate = need "C34" "额定负荷率"
                        let ratedGas = need "C35" "额定煤气耗量"
                        let operGas = need "C41" "额定工况煤气量"
                        let operAir = need "C42" "额定工况助燃风量"
                        let gasBore = need "C45" "煤气支管内径"
                        let gasVel = need "C46" "煤气支管流速"
                        let airBore = need "C47" "助燃风总管内径"
                        let airVel = need "C48" "助燃风流速"
                        let ringArea = need "C51" "环隙流通面积"
                        let ringVel = need "C52" "环隙流速"
                        let lhv = need "C5" "LHV kcal"
                        let density = need "C8" "密度"
                        let devGas = need "C19" "与 C86 偏差"
                        let devAir = need "C21" "与 C89 偏差"
                        let margin = need "C28" "裕量"

                        // 独立复算（不经 Excel）：三条主链必须对得上
                        let checkGas = SelectBoard.relDiff srcGas gas <= 1e-9
                        let checkAir = SelectBoard.relDiff srcAir air <= 1e-9
                        // 注意：设计热功率用的是 kcal 口径的 C5，不是补热表 C22 的 kJ 口径
                        let checkHeat = SelectBoard.relDiff (gas * lhv) heatKcal <= 1e-9
                        let checkRated = SelectBoard.relDiff (heatKcal * margin) ratedKcal <= 1e-9
                        let checkRate = SelectBoard.relDiff (100.0 / margin) rate <= 1e-9
                        let checkDens = SelectBoard.relDiff (10.662 / 22.414) density <= 1e-3

                        // 流速校核：三处都要落在经验区间
                        let vGasOk = gasVel >= 10.0 && gasVel <= 15.0
                        let vAirOk = airVel >= 10.0 && airVel <= 15.0
                        let vRingOk = ringVel >= 40.0 && ringVel <= 60.0

                        // 独立复算流速，证明公式里的 PI()/4 没写错
                        let calcGasVel =
                            SelectBoard.velocity operGas (SelectBoard.bore gasBore)

                        let calcAirVel =
                            SelectBoard.velocity operAir (SelectBoard.bore airBore)

                        let calcRingVel =
                            SelectBoard.velocity operAir (SelectBoard.annulus 393.7 273.0)

                        let velAgree =
                            SelectBoard.relDiff calcGasVel gasVel <= 1e-6
                            && SelectBoard.relDiff calcAirVel airVel <= 1e-6
                            && SelectBoard.relDiff calcRingVel ringVel <= 1e-6
                            && SelectBoard.relDiff (SelectBoard.annulus 393.7 273.0) ringArea <= 1e-9

                        // ---------- 风机段（用户手搭的骨架 + 下方风压分项） ----------
                        let ratedAir = need "C36" "额定助燃空气量"
                        let fanTheory = need "C68" "风机理论流量"
                        let fanFlow = need "C69" "风机选型流量"
                        let fanMargin = need "C70" "流量裕量"
                        let inletT = need "C71" "进风温度"
                        let fanFlowSec = need "C72" "风机流量 m3/s"
                        let fanDp = need "C73" "压升 kPa"
                        let fanEff = need "C74" "风机效率"
                        let shaft = need "C75" "轴功率"
                        let motorKw = nv "C76"
                        let motorMargin = nv "C77"
                        let fanVel = need "C80" "计算风量下管内流速"
                        let dpSum = need "C90" "风压分项合计"
                        let dpSumKpa = need "C91" "分项折算 kPa"
                        let dpDiff = need "C92" "与 C73 之差"
                        let auxMargin = need "E70" "E70 流量裕量"

                        let checkTheory =
                            SelectBoard.relDiff (ratedAir * SelectBoard.korr inletT) fanTheory <= 1e-9

                        let needFlow = fanTheory * auxMargin
                        let checkFlow = fanFlow >= needFlow - 1e-9 && fanFlow < needFlow + 100.0

                        let checkShaft =
                            SelectBoard.relDiff (fanFlowSec * fanDp / (fanEff / 100.0)) shaft <= 1e-9

                        let checkDpSum = SelectBoard.relDiff dpSumKpa (dpSum / 1000.0) <= 1e-12
                        let dpAgree = abs dpDiff <= 0.15 * (fanDp * 1000.0)

                        // 电机功率那一格是 XLL 数组函数（按轴功率取标准档），拿不到就不硬判
                        let checkMotor =
                            match motorKw, motorMargin with
                            | Some m, Some r -> m >= shaft * 0.999 && r >= 1.0 && r <= 1.6
                            | _ -> true

                        // 表结构与位置
                        let after = names ()

                        // 兜底：整张表不能有任何错误值（备注误写成 = 开头会变 #NAME?）
                        let errCells =
                            [ for i in 1 .. 102 do
                                  for j in 1 .. 4 do
                                      let addr = sprintf "%c%d" (char (64 + j)) i
                                      let s = try SelectBoard.txt (created.Range(addr).Text) with _ -> ""

                                      if s.StartsWith "#" then yield sprintf "%s=%s" addr s ]

                        let placedLast =
                            after.Length > 0
                            && after.[after.Length - 1] = SelectBoard.sheetName

                        let srcIntact =
                            let srcGas2 = SelectBoard.num (src.Range("C86").Value2)
                            let srcHeat2 = SelectBoard.num (src.Range("C26").Value2)

                            match srcGas2, srcHeat2 with
                            | Some a, Some b ->
                                SelectBoard.relDiff srcGas a <= 1e-12
                                && SelectBoard.relDiff srcHeat b <= 1e-12
                            | _ -> false

                        book.Save()

                        {| Before = before
                           After = after
                           PlacedLast = placedLast
                           SrcIntact = srcIntact
                           Gas = gas
                           Air = air
                           HeatKcal = heatKcal
                           Rate = rate
                           RatedKcal = ratedKcal
                           RatedGas = ratedGas
                           OperGas = operGas
                           OperAir = operAir
                           GasVel = gasVel
                           AirVel = airVel
                           RingVel = ringVel
                           RingArea = ringArea
                           Lhv = lhv
                           Density = density
                           DevGas = devGas
                           DevAir = devAir
                           Margin = margin
                           CheckGas = checkGas
                           CheckAir = checkAir
                           CheckHeat = checkHeat
                           CheckRated = checkRated
                           CheckRate = checkRate
                           CheckDens = checkDens
                           VelAgree = velAgree
                           ErrCells = errCells
                           VGasOk = vGasOk
                           VAirOk = vAirOk
                           VRingOk = vRingOk
                           RatedAir = ratedAir
                           FanTheory = fanTheory
                           FanFlow = fanFlow
                           FanMargin = fanMargin
                           InletT = inletT
                           FanFlowSec = fanFlowSec
                           FanDp = fanDp
                           FanEff = fanEff
                           Shaft = shaft
                           MotorKw = motorKw
                           MotorMargin = motorMargin
                           FanVel = fanVel
                           DpSum = dpSum
                           DpSumKpa = dpSumKpa
                           DpDiff = dpDiff
                           CheckTheory = checkTheory
                           CheckFlow = checkFlow
                           CheckShaft = checkShaft
                           CheckDpSum = checkDpSum
                           DpAgree = dpAgree
                           CheckMotor = checkMotor
                           Saved = (try book.Saved with _ -> true)
                           Books = (try app.Workbooks.Count with _ -> -1) |}
                    finally
                        app.DisplayAlerts <- alertsWas)

            // ---------- 5. 汇报 ----------
            Report.line output "原有工作表" (String.concat " / " result.Before)
            Report.line output "现有工作表" (String.concat " / " result.After)
            Report.line output "新表在末尾" (sprintf "%b" result.PlacedLast)
            Report.line output "补燃表未被改动" (sprintf "%b（C86 与 C26 前后一致）" result.SrcIntact)

            Report.line output "" ""
            Report.line output "LHV" (sprintf "%.4f kcal/Nm³" result.Lhv)
            Report.line output "煤气密度" (sprintf "%.4f kg/Nm³（独立复算 0.4757）" result.Density)

            Report.line
                output
                "设计煤气耗量"
                (sprintf "%.4f Nm³/h（与补燃表 C86 偏差 %.3e %%，%b）" result.Gas result.DevGas result.CheckGas)

            Report.line
                output
                "设计助燃空气量"
                (sprintf "%.4f Nm³/h（与补燃表 C89 偏差 %.3e %%，%b）" result.Air result.DevAir result.CheckAir)

            Report.line
                output
                "设计热功率"
                (sprintf "%.4f kcal/h = %.2f 万大卡/h = %.1f kW" result.HeatKcal (result.HeatKcal / 10000.0) (result.HeatKcal * 4.1868 / 3600.0))

            Report.line
                output
                "额定热负荷"
                (sprintf
                    "%.4f kcal/h = %.2f 万大卡/h = %.2f MW（裕量 %.2f）"
                    result.RatedKcal
                    (result.RatedKcal / 10000.0)
                    (result.RatedKcal * 4.1868 / 3600.0 / 1000.0)
                    result.Margin)

            Report.line
                output
                "额定煤气耗量"
                (sprintf "%.4f Nm³/h；工况 %.1f m³/h" result.RatedGas result.OperGas)

            Report.line output "额定负荷率" (sprintf "%.2f %%" result.Rate)

            Report.line
                output
                "煤气支管"
                (sprintf "DN250（内径 260）→ %.2f m/s　%s" result.GasVel (if result.VGasOk then "落在 10~15" else "超出 10~15"))

            Report.line
                output
                "助燃风总管"
                (sprintf "DN600 → %.2f m/s　%s" result.AirVel (if result.VAirOk then "落在 10~15" else "超出 10~15"))

            Report.line
                output
                "套管环隙"
                (sprintf "φ393.7 / φ273 → 面积 %.4f m²，流速 %.2f m/s　%s" result.RingArea result.RingVel (if result.VRingOk then "落在 40~60" else "超出 40~60"))

            Report.line output "流速独立复算" (sprintf "三处与公式值一致：%b" result.VelAgree)

            Report.line
                output
                "全表错误值"
                (if result.ErrCells.IsEmpty then "无（#NAME? / #REF! / #VALUE! 一个都没有）" else String.concat "、" result.ErrCells)

            Report.line output "不变式 设计耗量=C86" (sprintf "%b" result.CheckGas)
            Report.line output "不变式 助燃风=C89" (sprintf "%b" result.CheckAir)
            Report.line output "不变式 热功率=耗量×LHV" (sprintf "%b" result.CheckHeat)
            Report.line output "不变式 额定=设计×裕量" (sprintf "%b" result.CheckRated)
            Report.line output "不变式 负荷率=1/裕量" (sprintf "%b" result.CheckRate)
            Report.line output "不变式 密度=分子量/22.414" (sprintf "%b" result.CheckDens)

            Report.line output "" ""
            Report.line output "—— 助燃风机 ——" ""
            Report.line output "额定助燃风量（标态）" (sprintf "%.2f Nm³/h（α=1.15）" result.RatedAir)
            Report.line output "进风温度" (sprintf "%.1f ℃" result.InletT)

            Report.line
                output
                "理论流量"
                (sprintf "%.2f m³/h（= 标态 × %.4f 温度修正）" result.FanTheory (SelectBoard.korr result.InletT))

            Report.line output "流量裕量" (sprintf "%.4f（由 C69/C68 回算）" result.FanMargin)
            Report.line output "选型风量" (sprintf "%.2f m³/h —— 提给厂家的风量" result.FanFlow)
            Report.line output "计算流速" (sprintf "%.2f m/s（按选型风量走 DN600）" result.FanVel)
            Report.line output "取用压升" (sprintf "%.2f kPa" result.FanDp)
            Report.line output "效率" (sprintf "%.1f %%" result.FanEff)
            Report.line output "轴功率" (sprintf "%.2f kW" result.Shaft)

            Report.line
                output
                "电机功率"
                (match result.MotorKw, result.MotorMargin with
                 | Some m, Some r -> sprintf "%.1f kW（XLL 取标准档，安全裕量 %.3f）" m r
                 | Some m, None -> sprintf "%.1f kW" m
                 | None, _ -> "未接（C76 空着）")

            Report.line output "" ""
            Report.line output "风压分项估算" (sprintf "%.0f Pa = %.3f kPa；与 C73 取用值差 %.0f Pa" result.DpSum result.DpSumKpa result.DpDiff)
            Report.line output "不变式 理论流量=标态×温度修正" (sprintf "%b" result.CheckTheory)
            Report.line output "不变式 选型风量=理论×裕量取整" (sprintf "%b" result.CheckFlow)
            Report.line output "不变式 轴功率=Q×ΔP/效率" (sprintf "%b" result.CheckShaft)
            Report.line output "不变式 分项合计=各项和" (sprintf "%b" result.CheckDpSum)
            Report.line output "分项与 C73 对差 ≤15%" (sprintf "%b" result.DpAgree)
            Report.line output "电机功率 ≥ 轴功率且裕量 1.0~1.6" (sprintf "%b" result.CheckMotor)

            Report.line output "工作簿" (sprintf "Saved=%b；工作簿数 %d" result.Saved result.Books)

            Should.equal true result.PlacedLast
            Should.equal true result.SrcIntact
            Should.equal true result.CheckGas
            Should.equal true result.CheckAir
            Should.equal true result.CheckHeat
            Should.equal true result.CheckRated
            Should.equal true result.CheckRate
            Should.equal true result.VelAgree
            Should.equal true result.ErrCells.IsEmpty
            Should.equal true result.VGasOk
            Should.equal true result.VAirOk
            Should.equal true result.VRingOk
            Should.equal true result.CheckTheory
            Should.equal true result.CheckFlow
            Should.equal true result.CheckShaft
            Should.equal true result.CheckDpSum
            Should.equal true result.DpAgree
            Should.equal true result.CheckMotor
            Should.equal true result.Saved

            Report.conclusion output "通过：燃烧器选型表在末尾，燃料侧全部活链接补燃表；助燃风机按理论流量×裕量定选型风量，风压分项与取用值吻合，已 Save"
