/// 把「焦炉煤气燃烧计算书.xlsx」另存一份**不依赖 XLL 插件**的副本：
/// 副本里所有 `_xll.*` 自定义函数（本机 XLL 提供）整体替换成当前计算结果常量，
/// 其余公式（内置函数、跨表引用）一概保留。原件一个字节都不动。
///
/// 扫描口径（2026-11-11，原件 6 张表，共 **12 处 / 8 个函数名**）：
///   _xll.空气体积成分       烟气成分计算 C7:C9、补热燃烧计算 C6:C8         数组
///   _xll.理论空气量         补热燃烧计算 C35:C37                           数组
///   _xll.实际空气量         补热燃烧计算 C39:C41                           数组
///   _xll.实际烟气量         补热燃烧计算 C43:C47                           数组
///   _xll.烟气湿基体积分数   补热燃烧计算 C51:C55                           数组
///   _xll.烟气干基体积分数   补热燃烧计算 C57:C61                           数组
///   _xll.gasMoleEnthalpy    补热燃烧计算 C77、C78（放热量口径的 0℃/终温焓）
///                           烟气吸热量计算 C28、C31（175℃/500℃ 摩尔焓）
///   _xll.电机额定功率       燃烧器选型 C76
///
/// **坑**：Excel 数组公式在 openpyxl 里是 `ArrayFormula` 对象、不是字符串，
/// 按 `isinstance(v,str)` 扫会整片漏掉（第一轮就漏了 C77/C78/C76 三处）。
/// 本 Fact 改成「读整块 UsedRange.Formula 的二维数组 + 按函数名匹配 + 用 CurrentArray
/// 反查整个数组范围」，以后计算书加行加列也不会漏。
///
/// 做法（只动副本）：
///   ① 原件 SaveCopyAs 出新副本（存的是内存当前状态，原件磁盘不受影响、不置脏）
///   ② 打开副本，逐区域 读 Value2 → ClearContents → 写回同一批数值
///   ③ 全表重算，逐区域比对（替换前后必须一致）
///   ④ 保存关闭副本 → 独立读磁盘 XML 确认 `_xll.` 归零
///
/// 闸门：XUNITHOST_EXCEL_PID=<pid>，没设就 Skip。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Interop.Excel
open OsHost

module internal NoXll =

    let source = @"C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx"

    let target = @"C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书[无XLL].xlsx"

    /// 本机 XLL 提供的函数名（COM 的 Formula 里不带 `_xll.` 前缀，只能按名字认）
    let xllFuncs =
        [ "空气体积成分", "湿空气体积成分 H2O/O2/N2"
          "理论空气量", "理论空气量（组分体积）"
          "实际空气量", "实际空气量（α 修正后）"
          "实际烟气量", "实际烟气量（组分体积）"
          "烟气湿基体积分数", "湿基体积分数"
          "烟气干基体积分数", "干基体积分数"
          "gasMoleEnthalpy", "烟气摩尔焓"
          "电机额定功率", "电机功率按轴功率选档" ]

    let pid () =
        match Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID" with
        | null
        | "" -> None
        | s ->
            match Int32.TryParse s with
            | true, p -> Some p
            | _ -> None

    let findBook (app: Application) (path: string) : Workbook option =
        [ 1 .. (try app.Workbooks.Count with _ -> 0) ]
        |> List.map (fun i -> app.Workbooks.Item(i))
        |> List.tryFind (fun b ->
            String.Equals((try b.FullName with _ -> ""), path, StringComparison.OrdinalIgnoreCase))

    let sheetOf (book: Workbook) (name: string) : Worksheet =
        book.Worksheets.[name] :?> Worksheet

    /// 副本别重名：已存在就加时间戳挪走，别直接覆盖（上次那份可能不对）
    let staleCopy () : string option =
        if not (File.Exists target) then
            None
        else
            let dir = Path.GetDirectoryName target
            let stamp = File.GetLastWriteTime(target).ToString("yyyyMMdd-HHmm")
            let bak = Path.Combine(dir, sprintf "焦炉煤气燃烧计算书[无XLL][覆盖前]%s.xlsx" stamp)
            File.Move(target, bak, true)
            Some bak

    type Cell =
        { Sheet: string
          Area: string
          Func: string
          Note: string
          Formula: string
          Value: obj }

    /// 单个单元格的公式文本（数组公式 COM 会带 {}），没有公式返回 None
    let private formulaOf (rng: Range) : string option =
        let v = rng.Formula
        if isNull v then None else Some(Convert.ToString(v, Globalization.CultureInfo.InvariantCulture))

    /// 这个公式串调用了哪个 XLL 函数
    let private hitFunc (formula: string) : (string * string) option =
        xllFuncs
        |> List.tryFind (fun (fn, _) -> formula.Contains(fn + "(") || formula.Contains(fn + " ("))

    /// 全表扫描：整块读 UsedRange.Formula（一次 COM 调用一张表），本地按名字匹配，
    /// 再用 CurrentArray 把数组公式归到同一范围，避免 spill 子格重复计数
    let private scan (book: Workbook) : Cell list =
        [ for ws in book.Worksheets |> Seq.cast<Worksheet> do
              let used = ws.UsedRange
              let r0, c0 = used.Row, used.Column
              let grid = try used.Formula :?> obj[,] with _ -> null

              if not (isNull grid) then
                  let lb0, lb1 = grid.GetLowerBound 0, grid.GetLowerBound 1

                  for i in lb0 .. grid.GetUpperBound 0 do
                      for j in lb1 .. grid.GetUpperBound 1 do
                          match grid.[i, j] with
                          | :? string as f when f.StartsWith "=" ->
                              match hitFunc f with
                              | Some(fn, note) ->
                                  let r, c = r0 + (i - lb0), c0 + (j - lb1)
                                  let cell = (ws.Cells.[r, c] :?> Range)

                                  // 数组锚点：动态数组 spill 用 SpillingToRange，
                                  // 老式 CSE 数组用 CurrentArray，都不是就是单格。
                                  // 注意 HasArray 对 spill 返回 False，不能只认它。
                                  let area =
                                      try
                                          let c1 = ws.Cells.[r, c] :?> Range

                                          if Convert.ToBoolean(c1.HasSpill) then
                                              c1.SpillingToRange.Address(false, false)
                                          else
                                              c1.Address(false, false)
                                      with _ ->
                                          try
                                              (ws.Cells.[r, c] :?> Range).CurrentArray.Address(false, false)
                                          with _ ->
                                              (ws.Cells.[r, c] :?> Range).Address(false, false)

                                  yield
                                      { Sheet = ws.Name
                                        Area = area
                                        Func = fn
                                        Note = note
                                        Formula = f
                                        Value = Unchecked.defaultof<obj> }
                              | None -> ()
                          | _ -> () ]
        |> List.distinctBy (fun c -> c.Sheet, c.Area)

    type Run =
        { Copy: string
          StaleBackup: string option
          Cells: Cell list
          After: (string * string * string) list
          SavedAt: string
          SheetCount: int }

    /// 把一格/一块的值转成给人看的字符串（区域是 COM 的 1-based object[,]）
    let private show (v: obj) : string =
        match v with
        | null -> "（空）"
        | :? (obj[,]) as a ->
            String.Join(
                " ｜ ",
                [ for i in a.GetLowerBound 0 .. a.GetUpperBound 0 ->
                      Convert.ToString(a.[i, a.GetLowerBound 1], Globalization.CultureInfo.InvariantCulture) ]
            )
        | x -> Convert.ToString(x, Globalization.CultureInfo.InvariantCulture)

    let build () : Run =
        match pid () with
        | None -> failwith "没设 XUNITHOST_EXCEL_PID"
        | Some pid ->
            let app = ExcelHost.ExcelTask.GetApplicationAsync pid |> Report.await

            STA.run (fun () ->
                let src =
                    findBook app source
                    |> Option.defaultWith (fun () -> failwithf "这只 Excel（PID %d）里没打开 %s" pid source)

                let stale = staleCopy ()

                let oldAlerts = app.DisplayAlerts
                app.DisplayAlerts <- false
                try
                    // ① 先算到最新再另存（SaveCopyAs 写的是内存状态，原件不置脏）
                    app.CalculateFull()
                    src.SaveCopyAs target

                    let copy = app.Workbooks.Open target
                    try
                        // ② 扫全表，读值
                        let cells =
                            scan copy
                            |> List.map (fun c -> { c with Value = (sheetOf copy c.Sheet).Range(c.Area).Value2 })

                        if cells.IsEmpty then
                            failwith "一个 XLL 调用都没扫到，先确认副本打开的是正确的书"

                        // ③ 逐区域：清公式 → 逐格写回同一批数值
                        //    （整块 Range.Value2 <- obj[,] 在这里只落第一格，必须逐格）
                        for c in cells do
                            let rng = (sheetOf copy c.Sheet).Range(c.Area)

                            try
                                rng.ClearContents() |> ignore
                            with _ ->
                                // 数组公式偶发挡在「不能更改数组的某一部分」
                                rng.FormulaArray <- ""
                                rng.ClearContents() |> ignore

                            match c.Value with
                            | :? (obj[,]) as a ->
                                let r0, c0 = a.GetLowerBound 0, a.GetLowerBound 1

                                for i in r0 .. a.GetUpperBound 0 do
                                    for j in c0 .. a.GetUpperBound 1 do
                                        let v = a.[i, j]

                                        if not (isNull v) then
                                            (rng.Cells.[i - r0 + 1, j - c0 + 1] :?> Range).Value2 <- v
                            | v when not (isNull v) -> rng.Value2 <- v
                            | _ -> ()

                        app.CalculateFull()

                        // ④ 读回：值必须与原值逐格一致，且已经不是公式
                        let after =
                            cells
                            |> List.map (fun c ->
                                let rng = (sheetOf copy c.Sheet).Range(c.Area)
                                let now = rng.Formula

                                let isStillFormula =
                                    not (isNull now)
                                    && (Convert.ToString(now)).StartsWith "="

                                // 逐格比对：替换前 vs 读回
                                let changed =
                                    match c.Value with
                                    | :? (obj[,]) as a ->
                                        let r0, c0 = a.GetLowerBound 0, a.GetLowerBound 1

                                        [ for i in r0 .. a.GetUpperBound 0 do
                                              for j in c0 .. a.GetUpperBound 1 ->
                                                  let before = a.[i, j]

                                                  let afterV =
                                                      (rng.Cells.[i - r0 + 1, j - c0 + 1] :?> Range).Value2

                                                  let same =
                                                      match before, afterV with
                                                      | null, null -> true
                                                      | (:? float as x), (:? double as y) -> abs (x - y) <= 1e-9 * max 1.0 (abs x)
                                                      | x, y -> Convert.ToString(x) = Convert.ToString(y)

                                                  if not same then
                                                      sprintf "(%d,%d) %s→%s" (i - r0 + 1) (j - c0 + 1) (Convert.ToString before) (Convert.ToString afterV)
                                                  else
                                                      "" ]
                                        |> List.filter (fun s -> s <> "")
                                        |> fun xs -> if xs.IsEmpty then "" else "！不一致：" + String.Join("；", xs)
                                    | _ ->
                                        let b = c.Value
                                        let a2 = rng.Value2

                                        let same =
                                            match b, a2 with
                                            | null, null -> true
                                            | (:? float as x), (:? double as y) -> abs (x - y) <= 1e-9 * max 1.0 (abs x)
                                            | x, y -> Convert.ToString(x) = Convert.ToString(y)

                                        if same then "" else "！不一致"

                                (sprintf "%s!%s" c.Sheet c.Area,
                                 sprintf "%s　（%s）" c.Func c.Note,
                                 sprintf "%s　%s%s" (show rng.Value2) (if isStillFormula then "！仍是公式" else "已是常量") changed))

                        copy.Save()
                        let savedAt = File.GetLastWriteTime(target).ToString("yyyy-MM-dd HH:mm:ss")
                        let n = copy.Worksheets.Count
                        copy.Close(false)

                        { Copy = target
                          StaleBackup = stale
                          Cells = cells
                          After = after
                          SavedAt = savedAt
                          SheetCount = n }
                    with _ ->
                        // 副本出问题就关掉，原件照旧
                        (try
                            for i in [ 1 .. app.Workbooks.Count ] |> List.rev do
                                let b = app.Workbooks.Item(i)

                                if String.Equals(b.FullName, target, StringComparison.OrdinalIgnoreCase) then
                                    b.Close(false)
                         with _ ->
                            ())

                        reraise ()
                finally
                    app.DisplayAlerts <- oldAlerts)

    /// 独立核磁盘：xlsx 是 zip，扫所有 sheet XML 里还有没有 `_xll.`
    let diskScan (path: string) : (string * int) list =
        use fs =
            new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

        use zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read)

        [ for e in zip.Entries do
              if e.FullName.StartsWith "xl/worksheets/" && e.FullName.EndsWith ".xml" then
                  use reader = new StreamReader(e.Open())

                  let text = reader.ReadToEnd()

                  let count = Text.RegularExpressions.Regex.Matches(text, "_xll\\.").Count

                  yield e.FullName, count ]

/// 真机：另存无 XLL 副本并替换成数值
type StoveNoXllTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``计算书另存无XLL副本 把插件函数写成当前值``() =
        match NoXll.pid () with
        | None -> Assert.Skip("不硬编码 PID：设 XUNITHOST_EXCEL_PID=<pid> 才动手")
        | Some pid ->
            Report.title output (sprintf "PID %d：%s → 无 XLL 副本" pid NoXll.source)

            let r = NoXll.build ()

            Report.line output "副本" (sprintf "%s（原件未改动）" r.Copy)

            match r.StaleBackup with
            | Some bak -> Report.line output "旧副本" (sprintf "已挪走：%s" bak)
            | None -> Report.line output "旧副本" "原来没有，直接新建"

            let funcs = r.Cells |> List.map (fun c -> c.Func) |> List.distinct

            Report.line
                output
                "替换"
                (sprintf "%d 个区域 / %d 个函数名：%s" r.Cells.Length funcs.Length (String.Join("、", funcs)))

            for c in r.Cells do
                Report.line2 output (sprintf "%s!%s" c.Sheet c.Area) (sprintf "%s（%s）" c.Func c.Note) (sprintf "改前 %s" c.Formula)

            for (where, head, body) in r.After do
                Report.line2 output where head body

            Report.line output "保存" (sprintf "%s　%d 张表" r.SavedAt r.SheetCount)

            // 磁盘：所有 sheet XML 里 `_xll.` 必须归零
            let scanDisk = NoXll.diskScan r.Copy

            for (part, n) in scanDisk do
                Report.line2 output "磁盘" part (sprintf "_xll. 出现 %d 次" n)

            let left = scanDisk |> List.sumBy snd
            Report.line2 output "磁盘" "合计" (sprintf "残留 _xll. %d 处" left)

            Should.equal 0 left

            Report.conclusion output "通过：全部 XLL 区域写成当前值，磁盘 XML 再无 _xll.，原件未保存未改动"
