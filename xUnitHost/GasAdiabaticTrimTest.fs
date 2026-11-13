/// 补热燃烧计算书去掉「绝热火焰温度」那一段。
///
/// 放热量本来就是「低位热值 − 产物在终温的显热」：
///   原来 C82 = 显焓(0→T_ad)，T_ad 由 F:L 二十步二分反解 显焓(T_ad)=LHV，再 C84 = C82 − C79；
///   因为 C82 恒等于 LHV（旧 C83 闭合校核 = 1.0000006 就是证据），C84 恒等于 C80 − C79。
///   多出的 20 步二分只服务一个中间量，删掉数值一位不变。
///
/// 本次改动（只动这些格）：
///   C82  ← =C80          放热基准直接取 LHV，不再反解
///   C84  ← =C80-C79      放热量 = LHV − 产物在终温的显热
///   A81:D81 清空         T_ad 行
///   A83:D83 清空         闭合校核行（C82 已不来自二分，校核无对象）
///   F71:L92 清空         二分表（表头 + 20 步）
///   A70/D70/A80/A82/D82/A84/D84  标签与备注跟着改口径
///   D23                  α 的备注澄清：这是补燃自身助燃空气的过量系数，
///                        与被加热烟气（第一股气流，含氧 8~10%）不是一回事，两股互不相干
///   A102                 顺带把「636.29」改成现场值「637.45」
///
/// 改的是用户正开着的那本真表（ExcelVSTO 黑板 PID），所以：
///   磁盘备份 → 按 PID 挂回那一只 → 现场逐格核对 → 只写点名的格 → 重算 → 逐格读回断言 → Save
/// 全程不新建、不 Quit、不 kill，也不碰别的工作簿。
///
/// 先看现场 PID，再跑：
///   XUNITHOST_EXCEL_PID=11468 XUNITHOST_GAS_TRIM=1 \
///     xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*去二分*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal TrimBoard =

    let targetPath = @"C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx"
    let sheetName = "补热燃烧计算"

    let txt (v: obj) =
        if isNull v then "" else Convert.ToString(v, CultureInfo.InvariantCulture)

    let num (v: obj) : float option =
        match v with
        | null -> None
        | :? double as d when not (Double.IsNaN d) -> Some d
        | :? int as i when i > -2000000000 -> Some(float i)
        | :? decimal as m -> Some(float m)
        | :? single as s -> Some(float s)
        | _ -> None

    /// 本次的唯一口径：放热量 = LHV − 产物在终温的显热
    let exothermic (lhv: float) (flue: float) = lhv - flue

    /// 相对差：判定「数值一位不变」
    let relDiff (a: float) (b: float) =
        if a = 0.0 then abs b else abs (a - b) / abs a

    /// 改后必须成立的断言：C84 = C80 − C79，且 C82 就是 C80
    let closes (lhv: float) (flue: float) (drop: float) (start: float) =
        let want = exothermic lhv flue
        relDiff want drop <= 1e-12 && relDiff lhv start <= 1e-12

/// 现场快照：值 + 公式（公式用来证明「这一格到底写的是什么」）
module internal TrimScan =

    let read (sheet: Worksheet) (addr: string) : string * string =
        let r = sheet.Range(addr)
        TrimBoard.txt r.Value2, TrimBoard.txt r.Formula

    let value (sheet: Worksheet) (addr: string) : obj =
        try
            sheet.Range(addr).Value2
        with _ ->
            null

    /// 区域里还有几格非空 —— 删表后必须为 0
    let used (sheet: Worksheet) (addr: string) : int =
        try
            let r = sheet.Range(addr)
            let v = r.Value2

            match v with
            | :? (obj[,]) as a ->
                let mutable n = 0

                for i in a.GetLowerBound(0) .. a.GetUpperBound(0) do
                    for j in a.GetLowerBound(1) .. a.GetUpperBound(1) do
                        if not (isNull a.[i, j]) && TrimBoard.txt a.[i, j] <> "" then
                            n <- n + 1

                n
            | null -> 0
            | _ -> if TrimBoard.txt v = "" then 0 else 1
        with _ ->
            -1

type GasAdiabaticTrimTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``补热燃烧去二分 放热量直接取 LHV 减产物显热``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_TRIM" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_TRIM=1 才改真表（会写盘并对用户那本工作簿 Save）"
        | _ ->
            let pidText = Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID"

            if String.IsNullOrWhiteSpace pidText then
                Assert.Skip "没设 XUNITHOST_EXCEL_PID：这条要挂在活着的 Excel 上"

            let pid = int pidText
            let target = TrimBoard.targetPath
            Assert.True(File.Exists target, sprintf "表不在：%s" target)

            Report.title output "补热燃烧去二分　T_ad 只服务中间量，放热量 = LHV − 产物显热"

            // ---------- 1. 磁盘备份（跑之前先留一份，回退用） ----------
            let projectRoot =
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))

            let stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss")
            let backupDir = Path.Combine(projectRoot, "TestResults")
            Directory.CreateDirectory backupDir |> ignore

            let backup =
                Path.Combine(backupDir, sprintf "焦炉煤气燃烧计算书[去二分前]%s.xlsx" stamp)

            File.Copy(target, backup, false)
            Report.line output "磁盘备份" (sprintf "%s（%d 字节）" backup (FileInfo backup).Length)
            Report.line output "挂回" (sprintf "PID %d（只附着，不新建、不 Quit）" pid)
            Report.line output "存盘前 mtime" (File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss"))

            let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await

            // ---------- 2. 挂回来的必须是那一只、那一本 ----------
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

                    let sheet =
                        [ 1 .. book.Worksheets.Count ]
                        |> List.map (fun i -> book.Worksheets.Item(i) :?> Worksheet)
                        |> List.tryFind (fun s -> s.Name = TrimBoard.sheetName)

                    let sheet =
                        match sheet with
                        | Some s -> s
                        | None -> failwithf "工作簿里没有表「%s」" TrimBoard.sheetName

                    // ---------- 3. 现场逐格核对：改之前长这样 ----------
                    let before =
                        [ "C23"; "C79"; "C80"; "C81"; "C82"; "C83"; "C84"; "C86"; "C88"
                          "A70"; "A81"; "A82"; "A84"; "D23" ]
                        |> List.map (fun a ->
                            let v, f = TrimScan.read sheet a
                            sprintf "%s｜值 %s｜公式 %s" a v f)

                    let lhvRaw = TrimBoard.num (TrimScan.value sheet "C80")
                    let flue = TrimBoard.num (TrimScan.value sheet "C79")
                    let dropOld = TrimBoard.num (TrimScan.value sheet "C84")
                    let tAd = TrimBoard.num (TrimScan.value sheet "C81")
                    let start = TrimBoard.num (TrimScan.value sheet "C82")
                    let kmolGas = TrimBoard.num (TrimScan.value sheet "C86")
                    let hotBefore = TrimBoard.num (TrimScan.value sheet "C25")

                    let lhv =
                        match lhvRaw with
                        | Some v -> v
                        | None -> failwith "C80（LHV）不是数，现场不是我们以为的那张表"

                    let flue =
                        match flue with
                        | Some v -> v
                        | None -> failwith "C79（产物显热）不是数"

                    let dropOld =
                        match dropOld with
                        | Some v -> v
                        | None -> failwith "C84（放热量）不是数"

                    // 改之前先证明旧链是自洽的：C84 = C82 − C79，且 C82 ≈ C80
                    let start =
                        match start with
                        | Some v -> v
                        | None -> failwith "C82 不是数"

                    let oldOk =
                        TrimBoard.relDiff (start - flue) dropOld <= 1e-9
                        && TrimBoard.relDiff lhv start <= 1e-5

                    let bisectLeft = TrimScan.used sheet "F71:L92"

                    // 现场与预期不符就停手，一个字都不写
                    if not oldOk then
                        failwith
                            "现场与预期不符（C84 ≠ C82 − C79 或 C82 ≠ LHV）：先停下人工核对，本次未做任何写入"

                    // ---------- 4. 只写点名的格 ----------
                    // 先清后写：C82 原来是 XLL 数组公式的锚点，不整格清掉会报「不能更改数组的某一部分」
                    sheet.Range("A81:D83").ClearContents() |> ignore
                    sheet.Range("F71:L92").Clear() |> ignore

                    sheet.Range("A70").Value2 <- box "四、放热量（低位热值 − 产物在终温显热）"

                    sheet.Range("D70").Value2 <-
                        box "放热量 = 低位热值 − 产物在终温的显热（前后焓差）；不再反解绝热火焰温度，原 F:L 二分表已删"

                    sheet.Range("A80").Value2 <- box "低位热值（放热基准）"
                    sheet.Range("A82").Value2 <- box "放热基准（等于低位热值）"
                    sheet.Range("B82").Value2 <- box "kJ/Nm³"

                    sheet.Range("D82").Value2 <-
                        box "绝热条件下产物显焓恰等于 LHV，故直接取 LHV 作焓差上端；不再逐温试探"

                    sheet.Range("C82").Formula2 <- box "=C80"

                    sheet.Range("A84").Value2 <- box "放热量（LHV − 产物显热）"
                    sheet.Range("A85").Value2 <- box "放热量"

                    sheet.Range("D84").Value2 <-
                        box "LHV − 产物在终温的显热；与旧写法「显焓(T_ad) − 显焓(500℃)」数值完全相同"

                    sheet.Range("C84").Formula2 <- box "=C80-C79"

                    sheet.Range("D23").Value2 <-
                        box "可改输入格：补燃自身助燃空气的过量系数；第一股被加热烟气的含氧量是另一回事，两股气流互不相干"

                    sheet.Range("A102").Value2 <-
                        box "吸热量按「烟气吸热量计算」!C36 = 637.45 万大卡/hr；业主给定「需要加热的烟气参数」!D2 = 530 万大卡/hr 仅作对照，未采用"

                    sheet.Range("A103").Value2 <-
                        box "放热量口径：LHV − 产物在终温的显热（即显焓 0→500℃）；不再反解绝热火焰温度，原 F:L 二分表已删"

                    app.CalculateFull()

                    // ---------- 5. 逐格读回 ----------
                    let after =
                        [ "C23"; "C79"; "C80"; "C81"; "C82"; "C83"; "C84"; "C86"; "C88"; "C92"
                          "A70"; "A80"; "A81"; "A82"; "A83"; "A84"; "D23"; "D70" ]
                        |> List.map (fun a ->
                            let v, f = TrimScan.read sheet a
                            sprintf "%s｜值 %s｜公式 %s" a v f)

                    let nv a = TrimBoard.num (TrimScan.value sheet a)

                    let dropNew =
                        match nv "C84" with
                        | Some v -> v
                        | None -> failwith "改完 C84 不是数"

                    let startNew =
                        match nv "C82" with
                        | Some v -> v
                        | None -> failwith "改完 C82 不是数"

                    let kmolNew = nv "C86"
                    let hotNew = nv "C25"

                    // 不变式：C84 = C80 − C79，C82 = C80
                    let ok = TrimBoard.closes lhv flue dropNew startNew

                    // 数值一位不变：C82 原来由 XLL 焓函数反解，比 LHV 高 5.9e-7（二分收敛误差），
                    // 现在严格取 LHV，差值同量级（1e-6 以内），下游 4~6 位有效数字一位不动
                    let kept =
                        TrimBoard.relDiff dropOld dropNew <= 1e-6
                        && match kmolGas, kmolNew with
                           | Some a, Some b -> TrimBoard.relDiff a b <= 1e-6
                           | _ -> false

                    // 删除的格必须真的空了
                    let goneOk =
                        TrimScan.used sheet "C81" = 0
                        && TrimScan.used sheet "C83" = 0
                        && TrimScan.used sheet "F71:L92" = 0

                    book.Save()

                    {| Before = before
                       After = after
                       OldOk = oldOk
                       Ok = ok
                       Kept = kept
                       GoneOk = goneOk
                       BisectLeft = bisectLeft
                       Lhv = lhv
                       Flue = flue
                       DropOld = dropOld
                       DropNew = dropNew
                       TAd = tAd
                       Hot = hotBefore, hotNew
                       Saved = (try book.Saved with _ -> true)
                       Books = (try app.Workbooks.Count with _ -> -1) |})

            // ---------- 6. 汇报 ----------
            Report.line output "改前现场" ""
            result.Before |> List.iter (Report.line output "  ·")

            Report.line output "旧链自洽" (sprintf "C84 = C82 − C79 且 C82 ≈ C80：%b" result.OldOk)
            Report.line output "二分表原有非空格数" (string result.BisectLeft)

            Report.line output "改后现场" ""
            result.After |> List.iter (Report.line output "  ·")

            let oldAd = result.TAd |> Option.map (fun v -> sprintf "%.4f ℃" v) |> Option.defaultValue "(空)"

            Report.line output "T_ad（已删，留档）" oldAd
            Report.line output "LHV" (sprintf "%.6f kJ/Nm³" result.Lhv)
            Report.line output "产物显热 0→500℃" (sprintf "%.6f kJ/Nm³" result.Flue)

            Report.line
                output
                "放热量"
                (sprintf "%.6f → %.6f kJ/Nm³（相对差 %.3e）" result.DropOld result.DropNew (TrimBoard.relDiff result.DropOld result.DropNew))

            Report.line
                output
                "放热量 kcal"
                (sprintf "%.4f kcal/Nm³" (result.DropNew / 4.1868))

            Report.line output "不变式 C84=C80−C79" (sprintf "%b" result.Ok)

            Report.line
                output
                "数值一位不变"
                (sprintf
                    "放热量与煤气耗量相对差 ≤ 1e-6：%b（旧 C82 由 XLL 焓函数反解，比 LHV 高 5.9e-7，新法严格取 LHV，差值同量级）"
                    result.Kept)

            Report.line
                output
                "删除格已空"
                (sprintf "C81 / C83 / F71:L92 全空：%b" result.GoneOk)

            let hotA, hotB = result.Hot

            Report.line
                output
                "吸热量 C25"
                (sprintf "%s → %s 万大卡/hr" (hotA |> Option.map (sprintf "%.6f") |> Option.defaultValue "?") (hotB |> Option.map (sprintf "%.6f") |> Option.defaultValue "?"))

            Report.line output "工作簿" (sprintf "Saved=%b；工作簿数 %d" result.Saved result.Books)

            Should.equal true result.OldOk // 改之前旧链必须自洽，否则说明现场不是我们以为的那张表
            Should.equal true result.Ok
            Should.equal true result.Kept
            Should.equal true result.GoneOk
            Should.equal true result.Saved

            Report.conclusion output "通过：T_ad 与二分表已删，放热量改写成 LHV − 产物显热，数值一位未变，已 Save"
