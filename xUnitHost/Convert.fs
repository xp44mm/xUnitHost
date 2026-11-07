/// 老格式转新格式的**入口**：只做分派、批量和事后核对，不含任何转换细节。
///
/// 六条转换各自独立，住在自己的 host 库末尾，两边各有一个统一入口：
///   · WordHost/DocUpgrade.fs   `WordHost.DocToDocx`   .doc 无宏 → .docx  （正文逐字核对）
///                              `WordHost.DocToDocm`   .doc 有宏 → .docm  （正文 + 复核 VBA）
///                              `WordHost.DocUpgrade`  统一入口：判宏后选路
///   · ExcelHost/XlsUpgrade.fs  `ExcelHost.XlsToXlsx`  .xls 无宏 → .xlsx  （格子值指纹）
///                              `ExcelHost.XlsToXlsm`  .xls 有宏 → .xlsm  （格子值 + VBA）
///                              `ExcelHost.XlsUpgrade` 统一入口：判宏后选路
///   · PowerPointHost/PptUpgrade.fs  `PptToPptx`  .ppt 无宏 → .pptx （幻灯片文字）
///                                   `PptToPptm`  .ppt 有宏 → .pptm （文字 + VBA）
///                                   `PptUpgrade.pptUpgrade` 统一入口：判宏后选路
/// host 库只回传 (目标, 状态, 详情)，这里组装成 Outcome——依赖方向不倒过来。
///
/// 两条硬规矩不变：
///   · **转完立刻核对内容**——正文/格子值对不上就当失败，原件保留；
///   · **只删核对通过的**（`canDelete`），别的原件一个字都不动。
namespace XUnitHost

module Convert =

    /// host 库回传的三元组 → 测试项目的结果记录
    let private toOutcome (item: Item) (target: string, status: string, detail: string) : Outcome =
        { Source = item.Source
          Target = target
          Status = status
          Detail = detail }

    /// 正文比较口径（转发给 WordHost，保留老入口）
    let normalize (text: string) : string = WordHost.DocIO.normalize text

    /// 一个 Word 进程里转一批 .doc：判宏、选路都在 WordHost.DocUpgrade 里
    let words (items: Item list) : Outcome list =
        if items.IsEmpty then
            []
        else
            match WordHost.Host.execQuery(fun app ->
                items |> List.map(fun i -> WordHost.DocUpgrade.convert app i.Source |> toOutcome i)
            ) with
            | Ok outcomes -> outcomes
            | Error message ->
                items
                |> List.map(fun i ->
                    { Source = i.Source
                      Target = i.Target
                      Status = "失败"
                      Detail = "起 Word 失败：" + message })

    /// 一个 Excel 进程里转一批 .xls：判宏、选路都在 ExcelHost.XlsUpgrade 里
    let books (items: Item list) : Outcome list =
        if items.IsEmpty then
            []
        else
            match ExcelHost.Host.execQuery(fun app ->
                let results =
                    items |> List.map(fun i -> ExcelHost.XlsUpgrade.convert app i.Source |> toOutcome i)
                // 中途被判为「跳过」的那本还开着，收尾前先统统关掉（QuitWithoutSaving 也会再兜一次）
                ExcelHost.Host.CloseAllWorkbooksWithoutSaving app |> ignore
                results
            ) with
            | Ok outcomes -> outcomes
            | Error message ->
                items
                |> List.map(fun i ->
                    { Source = i.Source
                      Target = i.Target
                      Status = "失败"
                      Detail = "起 Excel 失败：" + message })

    /// 常驻 Excel（黑板挂回式）：app 是 ExcelHost.Host.GetApplication 交回的**现成实例**。
    /// 只逐份开/关自己那本工作簿 —— 绝不 CloseAllWorkbooksWithoutSaving、绝不 Quit、绝不新建。
    /// 调用方必须把这段活包在 OsHost.Sta.run 里（与附着同一条固定 STA 线程）。
    let booksAttached
        (app: Microsoft.Office.Interop.Excel.Application)
        (items: Item list)
        : Outcome list =
        items |> List.map(fun i -> ExcelHost.XlsUpgrade.convert app i.Source |> toOutcome i)

    /// 一个 PowerPoint 进程里转一批 .ppt：判宏、选路都在 PowerPointHost.PptUpgrade 里
    let shows (items: Item list) : Outcome list =
        if items.IsEmpty then
            []
        else
            match PowerPointHost.Host.execQuery(fun app ->
                let results =
                    items |> List.map(fun i -> PowerPointHost.PptUpgrade.convert app i.Source |> toOutcome i)
                PowerPointHost.Host.CloseAllPresentations app |> ignore
                results
            ) with
            | Ok outcomes -> outcomes
            | Error message ->
                items
                |> List.map(fun i ->
                    { Source = i.Source
                      Target = i.Target
                      Status = "失败"
                      Detail = "起 PowerPoint 失败：" + message })

    /// 整批：.doc 走一个 Word 进程，.xls 走一个 Excel 进程，.ppt 走一个 PowerPoint 进程
    let all (items: Item list) : Outcome list =
        let docs = items |> List.filter(fun i -> Paths.isWord i.Source)
        let sheets = items |> List.filter(fun i -> Paths.isExcel i.Source)
        let decks = items |> List.filter(fun i -> Paths.isPpt i.Source)
        words docs @ books sheets @ shows decks

    // ---------------- 独立的核对（测试和事后检查都用） ----------------

    /// 读一份 Word 文档的正文（去掉段落标记）
    let readWordText (path: string) : Result<string, string> =
        WordHost.Host.execQuery(fun app -> WordHost.DocIO.readText app path)

    /// 读一份工作簿的内容指纹
    let readBookPrint (path: string) : Result<string, string> =
        ExcelHost.Host.execQuery(fun app -> ExcelHost.BookIO.readPrint app path)

    /// 读一份演示文稿的幻灯片文字指纹
    let readPptPrint (path: string) : Result<string, string> =
        PowerPointHost.Host.execQuery(fun app -> PowerPointHost.PptIO.readPrint app path)

    /// 是不是「已经核对通过、可以删原件」的状态
    let canDelete (outcome: Outcome) =
        outcome.Status = "转了" || outcome.Status = "已有·一致"
