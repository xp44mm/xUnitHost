/// 【黑板挂回 + 临时工作簿】验证 Excel 365 两种公式写入语义：
///   Range.Formula   走旧语法，碰到可能返回多值的自定义函数会自动补隐式交集 @；
///   Range.Formula2  走动态数组语法，原样存，不加 @。
///
/// 用户问的是「写公式时」的差别，所以同一条 UDF 公式分别用两种属性写、再读回
/// Formula / Formula2 两个视图和计算值；还加同格反向覆盖，确认差别来自写入方式
/// 而不是函数本身。对照组装了内置 SUM（标量）和 UNIQUE（动态数组）。
///
/// 只在一本「新建、测完 Close(SaveChanges=false)」的临时工作簿里写：
/// 用户开着的「焦炉煤气燃烧计算书.xlsx」「天然气燃烧计算.xlsx」一个字节不动，
/// 不新建 Excel 实例、不 Quit、不 Save。COM 全在 OsHost.Sta.run 固定 STA 线程。
namespace XUnitHost

open System
open System.Globalization
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal Formula2Probe =

    type Case =
        { Label: string
          Address: string
          Way: string
          Written: string
          Formula: string
          Formula2: string
          Value: string
          Error: string }

    type Outcome =
        { Cases: Case list
          Spill: string list
          Existing: (string * string * string) list
          TempBook: string
          CountBefore: int
          CountAfter: int
          ActiveBefore: string
          ActiveAfter: string
          SavedBefore: bool
          SavedAfter: bool }

    let private txt (v: obj) =
        if isNull v then "" else Convert.ToString(v, CultureInfo.InvariantCulture)

    let private grab (f: unit -> obj) =
        try
            f () |> txt
        with ex ->
            "抛错：" + ex.Message

    /// 写一条公式，立刻局部重算，再读回两种公式视图与计算值。
    /// 备注：以 = 开头的文本一律当公式，别把说明文字交给它。
    let private one (sheet: Worksheet) (label: string) (addr: string) (way: string) (formula: string) : Case =
        let r = sheet.Range(addr)

        let error =
            try
                if way = "Formula" then
                    r.Formula <- box formula
                else
                    r.Formula2 <- box formula

                ""
            with ex ->
                "抛错：" + ex.Message

        sheet.Calculate()

        { Label = label
          Address = addr
          Way = way
          Written = formula
          Formula = grab (fun () -> r.Formula)
          Formula2 = grab (fun () -> r.Formula2)
          Value = grab (fun () -> r.Value2)
          Error = error }

    /// 只读观察上一轮用 Range.Formula 写进用户表的 XLL 公式：
    /// 磁盘 xlsx 里存的是旧语法（没有 @），真机两个视图分别长什么样，一并摆出来。
    let private existingViews (app: Application) : (string * string * string) list =
        match (FlueBoard.doc ()).Path with
        | None -> []
        | Some path ->
            let book =
                [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
                |> List.tryPick (fun i ->
                    let b = app.Workbooks.Item(i)
                    let full = try b.FullName with _ -> ""

                    if String.Equals(full, path, StringComparison.OrdinalIgnoreCase) then
                        Some b
                    else
                        None)

            match book with
            | None -> []
            | Some b ->
                let sheet =
                    [ 1 .. b.Worksheets.Count ]
                    |> List.tryPick (fun i ->
                        let s = b.Worksheets.Item(i) :?> Worksheet
                        if s.Name = "烟气吸热量计算" then Some s else None)

                match sheet with
                | None -> []
                | Some s ->
                    [ for a in [ "C28"; "C31" ] do
                          let r = s.Range(a)
                          yield a, grab (fun () -> r.Formula), grab (fun () -> r.Formula2) ]

    /// 新建临时工作簿。Excel 偶发 0x800AC472（对象模型暂时不可用 / 用户正在编辑单元格），
    /// 这个码不在 Common.Busy 的通用忙码里，就在这里局部兜一下，不改公用件。
    let private addBook (app: Application) : Workbook =
        let transient = [ 0x800AC472u; 0x80010001u; 0x8001010Au; 0x80010105u ]
        let clock = System.Diagnostics.Stopwatch.StartNew()
        let mutable book: Workbook = null
        let mutable last: exn = null

        while isNull (box book) && clock.ElapsedMilliseconds < 12000L do
            try
                book <- app.Workbooks.Add()
            with ex ->
                last <- ex

                if List.contains (uint32 ex.HResult) transient then
                    System.Threading.Thread.Sleep 400
                else
                    raise ex

        if isNull (box book) then
            raise last

        book

    let run (app: Application) : Outcome =
        let countBefore = try app.Workbooks.Count with _ -> 0
        let activeBefore = try app.ActiveWorkbook.Name with _ -> ""
        let savedBefore = try app.ActiveWorkbook.Saved with _ -> true

        let book = addBook app
        let sheet = book.Worksheets.Item(1) :?> Worksheet

        // 参考表「体积法scr烟气吸热」的组分与 kmol/h：让 UDF 真能算出 75260897.83
        let inputs =
            [ "H2O", 1039.1492236994736
              "O2", 157.08794503435354
              "N2", 4515.0316900151693
              "CO2", 522.37975372535027
              "SO2", 0.0 ]

        inputs
        |> List.iteri (fun i (name, kmol) ->
            let r = i + 2
            sheet.Range(sprintf "A%d" r).Value2 <- box name
            sheet.Range(sprintf "B%d" r).Value2 <- box kmol)

        let udf = "=gasMoleEnthalpy($A$2:$A$6,$B$2:$B$6,130)"
        let sumFormula = "=SUM($B$2:$B$6)"
        let uniq = "=UNIQUE($A$2:$A$6)"

        let cases = ResizeArray<Case>()
        let emit label addr way formula = one sheet label addr way formula |> cases.Add

        emit "UDF/F" "D2" "Formula" udf
        emit "UDF/F2" "D3" "Formula2" udf
        emit "SUM/F" "D4" "Formula" sumFormula
        emit "SUM/F2" "D5" "Formula2" sumFormula
        emit "UNIQ/F" "D7" "Formula" uniq
        emit "UNIQ/F2" "D9" "Formula2" uniq
        // 同一格反向覆盖：D2 改 Formula2、D3 改 Formula，看 @ 随写入方式走
        emit "覆盖/F2" "D2" "Formula2" udf
        emit "覆盖/F" "D3" "Formula" udf

        let spill =
            [ for r in 9..13 -> grab (fun () -> sheet.Range(sprintf "D%d" r).Value2) ]

        // 先把临时簿里的观测值读成纯数据，再关掉它（不保存、不弹框）
        let tempName = try book.Name with _ -> ""
        book.Close(box false)

        if activeBefore <> "" then
            (try
                app.Workbooks.Item(box activeBefore).Activate()
             with _ ->
                ())

        // 只读观察用户表现存公式放在最后：别在新建临时簿之前动 Excel 的对象模型
        { Cases = List.ofSeq cases
          Spill = spill
          Existing = existingViews app
          TempBook = tempName
          CountBefore = countBefore
          CountAfter = (try app.Workbooks.Count with _ -> 0)
          ActiveBefore = activeBefore
          ActiveAfter = (try app.ActiveWorkbook.Name with _ -> "")
          SavedBefore = savedBefore
          SavedAfter = (try app.ActiveWorkbook.Saved with _ -> true) }

type ExcelFormula2Test(output: ITestOutputHelper) =

    [<Fact>]
    member this.``公式写法对比：Formula 自动加隐式交集而 Formula2 不加``() =
        Report.title output "公式写法对比：Range.Formula 自动补 @，Range.Formula2 原样存"
        let info = FlueBoard.info()
        let before = Report.procIds info.Image
        Report.line output "黑板" (info.describe())
        Report.line output "挂回" (sprintf "pid=%d rot=%s（只附着，没新建、没退出）" info.Pid info.Rot)

        let app = ExcelHost.ExcelTask.GetApplicationAsync info.Pid |> Report.await
        let result = STA.run (fun () -> Formula2Probe.run app)

        Report.line output "临时簿" (sprintf "「%s」　写完 Close(SaveChanges=false)，不落盘" result.TempBook)
        Report.line output "输入" "A2:B6 = 参考表五组分 kmol/h；UDF 第三参数 130℃"
        Report.line output "UDF" "=gasMoleEnthalpy($A$2:$A$6,$B$2:$B$6,130)"
        Report.line output "对照" "=SUM($B$2:$B$6)（标量内置） / =UNIQUE($A$2:$A$6)（动态数组）"
        Report.line output "两个视图" "Formula 是旧兼容视图（隐藏 @）；Formula2 是动态数组视图（照实显示 @）"

        for c in result.Cases do
            Report.line
                output
                c.Label
                (sprintf
                    "%s 写 %s → Formula「%s」 Formula2「%s」 值「%s」%s"
                    c.Way
                    c.Address
                    c.Formula
                    c.Formula2
                    c.Value
                    (if c.Error = "" then "" else "　" + c.Error))

        Report.line output "UNIQUE 溢出" (sprintf "D9:D13 = %s" (String.concat " / " result.Spill))

        for (addr, f, f2) in result.Existing do
            Report.line output (sprintf "现表 %s" addr) (sprintf "Formula「%s」　Formula2「%s」" f f2)

        Report.line output "工作簿数" (sprintf "前 %d → 后 %d" result.CountBefore result.CountAfter)
        Report.line output "活动簿" (sprintf "「%s」 → 「%s」" result.ActiveBefore result.ActiveAfter)
        Report.line output "原簿 Saved" (sprintf "%b → %b" result.SavedBefore result.SavedAfter)
        let after = Report.procIds info.Image
        Report.line output "进程" (sprintf "EXCEL 干活前 %A / 收尾后 %A" (Set.toList before) (Set.toList after))

        // 断言放在报告之后：万一实测与预期相反，上面已经把真实文本全留下了
        let byLabel label = result.Cases |> List.find (fun c -> c.Label = label)
        let udfF = byLabel "UDF/F"
        let udfF2 = byLabel "UDF/F2"

        // 实测口径：@ 存在动态数组视图（Formula2 / 公式栏）里；旧兼容视图 Formula 读回会隐藏它
        Assert.True(udfF.Formula2.Contains("@"), sprintf "Formula 写入 UDF 后应被补 @：%s" udfF.Formula2)
        Assert.False(udfF.Formula.Contains("@"), sprintf "Formula 旧视图读回会隐藏 @：%s" udfF.Formula)
        Assert.False(udfF2.Formula2.Contains("@"), sprintf "Formula2 写入 UDF 后不该带 @：%s" udfF2.Formula2)
        Assert.True(udfF.Formula2 <> udfF2.Formula2, "两种写入方式存下来的公式文本应当不同")
        Should.equal udfF.Value udfF2.Value  // 标量 UDF：@ 不改计算结果

        Assert.False((byLabel "覆盖/F2").Formula2.Contains("@"), "同格改用 Formula2 覆盖后 @ 应消失")
        Assert.True((byLabel "覆盖/F").Formula2.Contains("@"), "同格改用 Formula 覆盖后 @ 应回来")

        Should.equal false ((byLabel "SUM/F").Formula2.Contains("@"))
        Should.equal false ((byLabel "SUM/F2").Formula2.Contains("@"))
        Assert.True((byLabel "UNIQ/F").Formula2.Contains("@"), "动态数组函数用 Formula 写同样会被补 @")
        Assert.False((byLabel "UNIQ/F2").Formula2.Contains("@"), "动态数组函数用 Formula2 写不该被补 @")
        Should.equal 5 (result.Spill |> List.filter (fun s -> s <> "") |> List.length)  // Formula2 写的真溢出五行

        Should.equal result.CountBefore result.CountAfter
        Should.equal result.ActiveBefore result.ActiveAfter
        Should.equal result.SavedBefore result.SavedAfter
        Should.equal before after

        Report.conclusion output "实测：Formula 写会补 @（只在 Formula2/公式栏可见，Formula 读回隐藏），Formula2 原样存；临时簿已关，现场没变"
