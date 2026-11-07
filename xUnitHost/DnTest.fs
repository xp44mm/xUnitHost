/// 管道公称直径：从三本「N#酸洗线设计计算」的【板式纤维除雾器】C6 读工况烟气量，
/// 按 15 m/s 上限反推最小常用 DN，回填【3条酸洗线设备清单】三张报价表的「入口电动挡板门」规格。
///
/// 纯规则（Dn）与真机动作（DnReal）分开：先跑规格 Fact 证明算得对，
/// 再跑「清点」Fact（只读），最后跑「回填」Fact（写盘 + 读回核对）。
/// 真表两条 Fact 都靠 XUNITHOST_DN_DIR 叫起来，不设就跳过，正常跑测试不受影响。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel

/// 纯规则：15 m/s 是上限，不是设计流速；流速必须严格小于它。
module internal Dn =

    /// 常用公称直径（mm）。不取 450 / 550 / 750 这类不常用的档，
    /// 与既有设计「按500来 / 按800来」的取法一致。
    let series =
        [ 100; 150; 200; 250; 300; 350; 400; 500; 600; 700
          800; 900; 1000; 1200; 1400; 1600; 1800; 2000 ]

    [<Literal>]
    let Limit = 15.0

    /// 工况烟气量 m³/h → 流速 v (m/s) 下需要的管内径 (mm)
    let requiredMm (q: float) (v: float) : float =
        sqrt (4.0 * q / (3600.0 * v * Math.PI)) * 1000.0

    /// 向上取第一个「严格大于所需内径」的常用 DN
    let pick (q: float) : int =
        let need = requiredMm q Limit

        match series |> List.tryFind(fun d -> float d > need) with
        | Some d -> d
        | None -> failwithf "烟气量 %.1f m³/h 需要 %.1f mm，超出 DN2000，要人工选档" q need

    /// 选定 DN 后的实际流速 (m/s)
    let velocity (q: float) (dn: int) : float =
        let d = float dn / 1000.0
        let area = Math.PI * d * d / 4.0
        q / (3600.0 * area)

    let spec (dn: int) : string = sprintf "DN%d" dn

/// 真表那一侧：路径、取表、读格、写盘、读回。不硬编码目录，靠环境变量叫起来。
module internal DnReal =

    let dir () =
        match Environment.GetEnvironmentVariable "XUNITHOST_DN_DIR" with
        | null
        | "" -> None
        | value -> Some(Path.GetFullPath value)

    let calcPath (dir: string) (n: int) = Path.Combine(dir, sprintf "%d#酸洗线设计计算-1104.xlsx" n)
    let quotePath (dir: string) = Path.Combine(dir, "3条酸洗线设备清单-1104.xlsx")
    let sheetName (n: int) = sprintf "%d#酸洗线脱硝设备报价" n

    /// 按表名找表，不假定它是第几张
    let findSheet (book: Workbook) (name: string) : Worksheet =
        let found =
            [ 1..book.Worksheets.Count ]
            |> List.tryPick(fun i ->
                let sheet = book.Worksheets.Item(i) :?> Worksheet
                if sheet.Name = name then Some sheet else None)

        match found with
        | Some sheet -> sheet
        | None -> failwithf "工作簿里没有工作表「%s」" name

    let cellText (sheet: Worksheet) (addr: string) : string =
        let value = sheet.Range(addr).Value2
        if isNull value then "" else Convert.ToString(value, CultureInfo.InvariantCulture)

    /// 只读打开一本设计计算表，取【板式纤维除雾器】C6（工况烟气量 m³/h）
    let readQ (xl: Application) (dir: string) (n: int) : float =
        let path = calcPath dir n
        Assert.True(File.Exists path, sprintf "找不到设计计算表：%s" path)
        let book = ExcelHost.BookIO.openBook xl path true

        try
            let sheet = findSheet book "板式纤维除雾器"
            Convert.ToDouble(sheet.Range("C6").Value2, CultureInfo.InvariantCulture)
        finally
            ExcelHost.BookIO.closeBook book

    type Row =
        { N: int
          Q: float
          NeedMm: float
          Dn: int
          V: float
          Spec: string }

    type Outcome =
        { Rows: Row list
          Before: string list
          After: string list
          Neighbor: string
          Backup: string
          BackedUp: bool
          Identity: string
          SavedAt: string }

    /// 原稿留一份：文件名带落盘时刻，重复跑不会覆盖已有备份
    let backupOf (dir: string) (target: string) : string =
        let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmm")
        Path.Combine(dir, sprintf "3条酸洗线设备清单-1104[原稿]%s.xlsx" stamp)

    /// 一段 Excel 命令里做完：读三本计算表 → 备份 → 回填三张报价表 → 存盘 → 只读读回核对
    let write (dir: string) : Outcome =
        Host.execOnceQuery Kind.Excel (fun app ->
            match app with
            | ExcelApp xl ->
                xl.AskToUpdateLinks <- false

                let rows =
                    [ 1..3 ]
                    |> List.map(fun n ->
                        let q = readQ xl dir n
                        let dn = Dn.pick q

                        { N = n
                          Q = q
                          NeedMm = Dn.requiredMm q Dn.Limit
                          Dn = dn
                          V = Dn.velocity q dn
                          Spec = Dn.spec dn })

                let target = quotePath dir
                Assert.True(File.Exists target, sprintf "找不到设备清单：%s" target)

                let backup = backupOf dir target
                let backedUp = not (File.Exists backup)
                if backedUp then File.Copy(target, backup, false)

                // 这是用户点名的现有工作簿（不是测试临时文档）：动笔前先认人，把身份贴进报告
                let book = ExcelHost.BookIO.openBook xl target false
                let identity = Doc.ofCom (fun () -> book.FullName) (fun () -> book.Path) |> Doc.describe

                // 2# 表没有挡板门行，C4 是除雾器规格；先记下它的 B4，插行后核对原行没丢
                let before =
                    rows
                    |> List.map(fun row ->
                        let sheet = findSheet book (sheetName row.N)
                        cellText sheet (if row.N = 2 then "B4" else "C4"))

                for row in rows do
                    let sheet = findSheet book (sheetName row.N)

                    if row.N = 2 then
                        // 照 1# / 3# 的样子插一行：整行插入，Excel 自己把下方公式的行号顺延
                        sheet.Range("4:4").Insert() |> ignore
                        // 连格式带文字从 1# 的挡板门行拷过来，只要改规格那一格
                        let src = findSheet book (sheetName 1)
                        src.Range("4:4").Copy(box (sheet.Range("4:4"))) |> ignore

                    sheet.Range("C4").Value2 <- row.Spec

                book.Save()
                let savedAt = File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss")
                ExcelHost.BookIO.closeBook book

                // 存完只读重新开一遍，逐格读回来
                let check = ExcelHost.BookIO.openBook xl target true

                let after =
                    rows
                    |> List.map(fun row ->
                        let sheet = findSheet check (sheetName row.N)
                        cellText sheet "C4")

                let s2 = findSheet check (sheetName 2)

                let neighbor =
                    sprintf
                        "B4=「%s」　B5=「%s」　C5 含「%s」"
                        (cellText s2 "B4")
                        (cellText s2 "B5")
                        (if (cellText s2 "C5").Contains("玻璃钢纤维除雾塔") then "玻璃钢纤维除雾塔" else "！原规格丢了")

                ExcelHost.BookIO.closeBook check

                { Rows = rows
                  Before = before
                  After = after
                  Neighbor = neighbor
                  Backup = backup
                  BackedUp = backedUp
                  Identity = identity
                  SavedAt = savedAt }
            | _ -> failwith "kind 不对：这条命令只走 Excel")
        |> Report.ok

/// 规格 Fact：纯逻辑，不起 Excel
type DnSpecTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``管径：按15m/s上限反推常用DN，三条线流速都小于15``() =
        let qs = [ 6439.56043956044; 8586.080586080587; 22538.46153846154 ]
        Report.title output "管径：15 m/s 上限反推常用 DN，三条线流速都小于 15"

        let picked =
            qs
            |> List.mapi(fun i q ->
                let dn = Dn.pick q
                let v = Dn.velocity q dn

                Report.line2
                    output
                    (sprintf "%d#" (i + 1))
                    "烟气量"
                    (sprintf
                        "%10.1f m³/h → 需内径 %7.1f mm → %s，实际 %6.2f m/s"
                        q
                        (Dn.requiredMm q Dn.Limit)
                        (Dn.spec dn)
                        v)

                dn, v)

        Should.equal [ 400; 500; 800 ] (picked |> List.map fst)
        Assert.True(picked |> List.forall(fun (_, v) -> v < Dn.Limit))
        Report.conclusion output "通过：按上限反推、向上取常用 DN，三条线的流速都严格小于 15 m/s"

/// 真表两条：先清点（只读），再回填（写盘）
type DnFolderTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``真表：清点三本计算表的工况烟气量与应填DN（一个文件都不动）``() =
        match DnReal.dir () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_DN_DIR=<目录> 才清点")
        | Some dir ->
            Report.requireClean "EXCEL"
            Assert.True(Directory.Exists dir, sprintf "目录不存在：%s" dir)

            let qs =
                Host.execOnceQuery Kind.Excel (fun app ->
                    match app with
                    | ExcelApp xl -> [ 1..3 ] |> List.map(DnReal.readQ xl dir)
                    | _ -> failwith "kind 不对：这条命令只走 Excel")
                |> Report.ok

            Report.title output (sprintf "清点 %s（只读，一个文件都不动）" dir)

            for i, q in qs |> List.indexed do
                let dn = Dn.pick q

                Report.line2
                    output
                    (sprintf "%d#" (i + 1))
                    "C6"
                    (sprintf
                        "%10.1f m³/h → 需内径 %7.1f mm → %s，实际 %6.2f m/s"
                        q
                        (Dn.requiredMm q Dn.Limit)
                        (Dn.spec dn)
                        (Dn.velocity q dn))

            Should.equal [ "DN400"; "DN500"; "DN800" ] (qs |> List.map(Dn.pick >> Dn.spec))
            Report.conclusion output "只清点：三本计算表只读打开，设备清单一个字没动"

    [<Fact>]
    member this.``真表：回填报价表三张表的入口电动挡板门DN``() =
        match DnReal.dir () with
        | None -> Assert.Skip("不硬编码路径：设 XUNITHOST_DN_DIR=<目录> 才回填")
        | Some dir ->
            Report.requireClean "EXCEL"
            Assert.True(Directory.Exists dir, sprintf "目录不存在：%s" dir)

            let result = DnReal.write dir
            Report.title output "真表：三本设计计算 → 报价表「入口电动挡板门」DN"
            Report.line output "认人" result.Identity

            Report.line
                output
                "备份"
                (if result.BackedUp then result.Backup else "已有备份，未重复生成：" + result.Backup)

            for (row, before, after) in List.zip3 result.Rows result.Before result.After do
                Report.line2
                    output
                    (sprintf "%d#" row.N)
                    "烟气量"
                    (sprintf
                        "%10.1f m³/h → 需内径 %7.1f mm → %s，实际 %6.2f m/s"
                        row.Q
                        row.NeedMm
                        row.Spec
                        row.V)

                Report.line2
                    output
                    (sprintf "%d#" row.N)
                    "回填"
                    (sprintf "原「%s」→ 写「%s」→ 读回「%s」" before row.Spec after)

            Report.line output "2# 插行" result.Neighbor
            Report.line output "保存" result.SavedAt

            Should.equal [ "DN400"; "DN500"; "DN800" ] (result.Rows |> List.map(fun r -> r.Spec))
            Should.equal [ "DN400"; "DN500"; "DN800" ] result.After
            Assert.True(result.Rows |> List.forall(fun r -> r.V < Dn.Limit))
            Assert.Contains("入口电动挡板门", result.Neighbor)
            Assert.Contains("玻璃钢纤维除雾塔", result.Neighbor)
            Report.conclusion output "通过：三格 DN 写进去又读回来一致，2# 插行后原除雾器规格还在，原稿已备份"
