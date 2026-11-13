/// 【黑板挂回】焦炉煤气燃烧计算书.xlsx /「需要加热的烟气参数」!B2:B10：
/// 把九行「名称，单位」拆开 —— 名称留在 B 原位，单位写到右边 C 列。
///
/// 与 AttachedConvertTest 同一条通道：ExcelVSTO 抄现场 → 按 PID 挂回现成实例，
/// 绝不新建 / 绝不 Quit / 绝不 kill；所有 COM 读写排在 OsHost.Sta.run 那条固定 STA 线程上。
/// 不 requireClean：常驻实例在跑是前提，不是障碍。
/// 这是用户点名的真表（不是测试自己新建的临时文档），所以按 PriceTest / DnTest 的口径：
/// 先备份到项目目录外 → 写之前逐行核对现场 → 存盘 → 逐格读回核对。
namespace XUnitHost

open System
open System.Globalization
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

/// 纯规则：第一个逗号（全角 / 半角）左边是名称、右边是单位；拆不出就 None，绝不乱写。
module internal UnitSplit =
    let private marks = [| '，'; ',' |]

    let split (raw: string) : (string * string) option =
        if String.IsNullOrWhiteSpace raw then
            None
        else
            let i = raw.IndexOfAny marks

            if i <= 0 || i >= raw.Length - 1 then
                None
            else
                let name = raw.Substring(0, i).Trim()
                let unit = raw.Substring(i + 1).Trim()

                if String.IsNullOrWhiteSpace name || String.IsNullOrWhiteSpace unit then
                    None
                else
                    Some(name, unit)

module internal GasParam =

    /// 插件 2026-11-10 14:04:54 抄下来的现场，原文留在代码里当记录。
    /// 展示文本的解析只走 BlackboardText（测试项目里的参考实现），OsHost 不认文本。
    let boardText =
        [ @"[ExcelVSTO 当前 Excel 实例] 采集时间 2026-11-10 14:04:54"
          @"PID : 28884"
          @"ROT : !ExcelVSTO.Instance.28884"
          @"HWND : 0x000A04D6"
          @"EXCEL.EXE : C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
          @"工作簿 : 焦炉煤气燃烧计算书.xlsx"
          @"路径 : C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx" ]
        |> String.concat "\n"

    let info () : BoardApp = BlackboardText.ofText boardText |> Report.ok
    let doc () : BoardDoc = BlackboardText.docOfText boardText

    let sheetName = "需要加热的烟气参数"

    /// 用户贴出来的 B2:B10 原文。写之前逐行核对，现场变了就停手。
    let expected =
        [ "热交换量，万大卡/hr"
          "需加热烟气量，Nm3/h"
          "烟气原始温度，℃"
          "加热后烟气温度，℃"
          "烟气压力，kPa"
          "直接燃烧器尺寸，mm"
          "助燃空气温度，℃"
          "烟气含氧量，%"
          "主燃料及低位热值，kcal/Nm³" ]

    let pairs () : (string * string) list =
        expected
        |> List.map(fun raw ->
            match UnitSplit.split raw with
            | Some pair -> pair
            | None -> failwithf "拆分规则算不出来：%s" raw)

    let private bookPath () : string =
        match (doc ()).Path with
        | Some path -> path
        | None -> failwith "黑板里没有落盘路径，不往下走"

    /// 只认已经开着的那一本，绝不 Workbooks.Open 再开一份
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
        | None ->
            let opened =
                [ 1..count ]
                |> List.map(fun i ->
                    try app.Workbooks.Item(i).FullName with _ -> "(读不出名字)")

            failwithf "这只 Excel 里没开着「%s」；现有工作簿：%s" path (String.concat "、 " opened)

    let private findSheet (book: Workbook) : Worksheet =
        let found =
            [ 1..book.Worksheets.Count ]
            |> List.tryPick(fun i ->
                let sheet = book.Worksheets.Item(i) :?> Worksheet
                if sheet.Name = sheetName then Some sheet else None)

        match found with
        | Some sheet -> sheet
        | None ->
            let names =
                [ 1..book.Worksheets.Count ]
                |> List.map(fun i -> (book.Worksheets.Item(i) :?> Worksheet).Name)

            failwithf "工作簿里没有表「%s」；现有：%s" sheetName (String.concat "、" names)

    let cellText (sheet: Worksheet) (addr: string) : string =
        let value = sheet.Range(addr).Value2
        if isNull value then "" else Convert.ToString(value, CultureInfo.InvariantCulture)

    let column (sheet: Worksheet) (col: string) : string list =
        [ for row in 2..10 -> cellText sheet (sprintf "%s%d" col row) ]

    let private merged (sheet: Worksheet) (addr: string) : bool =
        try
            let value = sheet.Range(addr).MergeCells
            not (isNull value) && Convert.ToBoolean(value, CultureInfo.InvariantCulture)
        with _ ->
            false

    type Survey =
        { BookName: string
          BookPath: string
          SheetName: string
          UsedRange: string
          A: string list
          B: string list
          C: string list
          D: string list
          E: string list
          MergedB: bool
          MergedC: bool
          BookSaved: bool
          DiskAt: string }

    /// 只读现场：A2:E10 全读出来，右边那几列有没有内容一目了然（不写、不存盘）
    let survey (app: Application) : Survey =
        let path = bookPath ()
        let book = findBook app path
        let sheet = findSheet book

        { BookName = book.Name
          BookPath = book.FullName
          SheetName = sheet.Name
          UsedRange = (try sheet.UsedRange.Address() with _ -> "(读不出)")
          A = column sheet "A"
          B = column sheet "B"
          C = column sheet "C"
          D = column sheet "D"
          E = column sheet "E"
          MergedB = merged sheet "B2"
          MergedC = merged sheet "C2"
          BookSaved = (try book.Saved with _ -> true)
          DiskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss") }

    /// 写盘：现场复核 → 九行逐格拆 → Save() → 同一次 STA 调用里逐格读回
    let write (app: Application) : string list * string list * bool * string =
        let path = bookPath ()
        let book = findBook app path
        let sheet = findSheet book

        // 写之前再验一次：现场跟黑板上贴的九行一模一样，C 列全空，B/C 都不是合并格
        Should.equal expected (column sheet "B")
        Should.equal (List.replicate 9 "") (column sheet "C")
        Assert.True(not (merged sheet "B2"), "B2 是合并单元格，拆开前要先人工拆并")
        Assert.True(not (merged sheet "C2"), "C2 是合并单元格，拆开前要先人工拆并")

        for row, (name, unit) in List.zip [ 2..10 ] (pairs ()) do
            sheet.Range(sprintf "B%d" row).Value2 <- box name
            sheet.Range(sprintf "C%d" row).Value2 <- box unit

        book.Save()

        let bAfter = column sheet "B"
        let cAfter = column sheet "C"
        let saved = (try book.Saved with _ -> true)
        let diskAt = File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss")
        bAfter, cAfter, saved, diskAt

/// 三条 Fact：纯规则 → 只读清点 → 真写盘（写盘那条要显式设环境变量）
type GasParamTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``拆分规则：九行名称单位都能拆开``() =
        Report.title output "拆分规则：第一个逗号左边名称、右边单位"

        for raw, pair in List.zip GasParam.expected (GasParam.pairs ()) do
            Report.line output raw (sprintf "名称「%s」＋单位「%s」" (fst pair) (snd pair))

        Should.equal 9 (GasParam.pairs ()).Length
        Should.equal "热交换量" (GasParam.pairs () |> List.head |> fst)
        Should.equal "kcal/Nm³" (GasParam.pairs () |> List.last |> snd)
        Report.line output "无逗号" (sprintf "%A" (UnitSplit.split "烟气压力kPa"))
        Report.line output "缺单位" (sprintf "%A" (UnitSplit.split "烟气压力，"))
        Assert.True((UnitSplit.split "烟气压力kPa").IsNone)
        Assert.True((UnitSplit.split "烟气压力，").IsNone)
        Report.conclusion output "通过：九行全部拆得开，缺逗号 / 缺单位一律不写；纯逻辑，不碰真机"

    [<Fact>]
    member this.``烟气参数清点：只读现场 A2:E10``() =
        let info = GasParam.info ()
        let doc = GasParam.doc ()
        let before = Report.procIds info.Image

        Report.title output "烟气参数清点：挂回常驻 Excel，只读 A2:E10"
        Report.line output "参数表" (info.describe())
        Report.line output "文档" (doc.describe())
        Report.line output "现场" (sprintf "%s 进程 %A" info.Image (Set.toList before))

        let rawType =
            STA.run(fun () ->
                match Com.runningAt info.Rot with
                | Some raw -> raw.GetType().FullName
                | None -> "None")

        Report.line output "ROT 对象" rawType

        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

        let snap = STA.run(fun () -> GasParam.survey app)
        Report.line output "工作簿" (sprintf "「%s」 %s" snap.BookName snap.BookPath)
        Report.line output "工作表" (sprintf "「%s」 UsedRange=%s" snap.SheetName snap.UsedRange)

        for i in 0..8 do
            Report.line
                output
                (sprintf "第%d行" (i + 2))
                (sprintf
                    "A=「%s」 B=「%s」 C=「%s」 D=「%s」 E=「%s」"
                    snap.A.[i]
                    snap.B.[i]
                    snap.C.[i]
                    snap.D.[i]
                    snap.E.[i])

        Report.line output "C 列" (sprintf "全空=%b；内容=%s" (snap.C |> List.forall(fun v -> v = "")) (String.concat "|" snap.C))
        Report.line output "合并格" (sprintf "B2=%b C2=%b" snap.MergedB snap.MergedC)
        Report.line output "工作簿 Saved" (string snap.BookSaved)
        Report.line output "磁盘最后写入" snap.DiskAt

        Should.equal GasParam.expected snap.B
        Should.equal true (Set.contains info.Pid before)

        let after = Report.procIds info.Image
        Report.line output "收尾" (sprintf "%s 进程 %A" info.Image (Set.toList after))
        Should.equal before after
        Report.conclusion output "通过：B2:B10 与黑板贴的九行一致；只读，磁盘一个字节没动"

    [<Fact>]
    member this.``烟气参数拆分：名称留 B 单位进 C``() =
        match Environment.GetEnvironmentVariable "XUNITHOST_GAS_SPLIT" with
        | null
        | "" -> Assert.Skip "设 XUNITHOST_GAS_SPLIT=1 才真的写盘"
        | _ ->
            let info = GasParam.info ()
            let doc = GasParam.doc ()
            let before = Report.procIds info.Image
            let target =
                match doc.Path with
                | Some path -> path
                | None -> failwith "黑板里没有落盘路径，不往下走"

            Assert.True(File.Exists target, sprintf "文件不在：%s" target)

            let projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
            let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmmss")
            let backup = Path.Combine(projectRoot, "TestResults", sprintf "焦炉煤气燃烧计算书[拆分前]%s.xlsx" stamp)
            let backedUp = not (File.Exists backup)
            if backedUp then File.Copy(target, backup, false)

            Report.title output "烟气参数拆分：名称留在 B2:B10，单位写到 C2:C10"
            Report.line output "参数表" (info.describe())
            Report.line output "文档" (doc.describe())
            Report.line output "备份" (if backedUp then backup else "已有备份，未重复生成：" + backup)
            Report.line output "备份字节" (sprintf "%d B" (FileInfo backup).Length)

            let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
            Report.line output "挂回" (sprintf "Ok pid=%d（只附着，没新建、没退出）" info.Pid)

            let bBefore, cBefore, diskBefore =
                STA.run(fun () ->
                    let snap = GasParam.survey app
                    snap.B, snap.C, snap.DiskAt)

            Report.line output "存盘前" diskBefore
            Should.equal GasParam.expected bBefore
            Should.equal (List.replicate 9 "") cBefore

            let bAfter, cAfter, saved, diskAfter = STA.run(fun () -> GasParam.write app)

            for i in 0..8 do
                Report.line
                    output
                    (sprintf "第%d行" (i + 2))
                    (sprintf "B「%s」→「%s」　C「%s」→「%s」" bBefore.[i] bAfter.[i] cBefore.[i] cAfter.[i])

            Report.line output "工作簿 Saved" (string saved)
            Report.line output "存盘后" diskAfter
            Should.equal (GasParam.pairs () |> List.map fst) bAfter
            Should.equal (GasParam.pairs () |> List.map snd) cAfter
            Should.equal true saved
            Should.equal true (File.Exists backup)

            let after = Report.procIds info.Image
            Report.line output "现场" (sprintf "%s 进程 干活前 %A / 收尾后 %A" info.Image (Set.toList before) (Set.toList after))
            Should.equal before after
            Should.equal true (Set.contains info.Pid after)
            Report.conclusion output "完成：九行名称留 B、单位进 C，已存盘并逐格读回一致；原件备份在项目目录外，实例没新建也没退出"
