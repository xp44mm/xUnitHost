/// 转换在测试项目这一侧的底座：数据类型 + 路径规则。
/// Word / Excel 的开、关、指纹、判宏等 COM 细节已经归到 `WordHost` / `ExcelHost` 里，
/// 这里**不含任何 Office 调用**——host 库只回传 (目标, 状态, 详情)，由 `Convert.fs` 组装 Outcome。
namespace XUnitHost

open System
open System.IO

/// 一份待办：原文件 → 目标文件。Target 是扫描时估的，真正落盘由各条转换自己定。
type Item = { Source: string; Target: string }

/// 一份的结果
type Outcome =
    { Source: string
      Target: string
      /// 转了 / 已有·一致 / 跳过 / 失败
      Status: string
      Detail: string }

module Paths =

    let private isCandidate (path: string) =
        let name = Path.GetFileName path

        // ~$ 开头的是 Office 的临时锁文件，不是文档；.docx/.xlsx/.pptx 不会被后缀判断误收
        not (name.StartsWith "~$")
        && (name.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith(".ppt", StringComparison.OrdinalIgnoreCase))

    /// 跳过缓存目录和隐藏目录：`.idx_cache` 里放的是索引缓存，转它只会把缓存搞坏
    let private skipDir (path: string) =
        let name = Path.GetFileName path
        name.StartsWith "." || name.Equals(".idx_cache", StringComparison.OrdinalIgnoreCase)

    let isWord (path: string) = path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)
    let isExcel (path: string) = path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
    let isPpt (path: string) = path.EndsWith(".ppt", StringComparison.OrdinalIgnoreCase)

    /// 按类型分轮跑：XUNITHOST_CONVERT_KIND=doc|xls|ppt，不设 / 不认得 = 全跑。
    /// 某一类被弹窗或坏文件卡住，不拖累另外两类；失败归因也清楚。
    type Kind =
        | Doc
        | Xls
        | Ppt
        | All

    let parseKind (text: string) : Kind =
        match (defaultArg (Option.ofObj text) "").Trim().ToLowerInvariant() with
        | "doc" | ".doc" | "word" -> Doc
        | "xls" | ".xls" | "excel" -> Xls
        | "ppt" | ".ppt" | "powerpoint" -> Ppt
        | _ -> All

    let kindFromEnv () : Kind =
        parseKind (Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_KIND")

    let kindLabel (kind: Kind) : string =
        match kind with
        | Doc -> "doc"
        | Xls -> "xls"
        | Ppt -> "ppt"
        | All -> "全部"

    let filterByKind (kind: Kind) (items: Item list) : Item list =
        match kind with
        | All -> items
        | Doc -> items |> List.filter(fun i -> isWord i.Source)
        | Xls -> items |> List.filter(fun i -> isExcel i.Source)
        | Ppt -> items |> List.filter(fun i -> isPpt i.Source)

    /// 本轮上限（翻倍法用）：XUNITHOST_CONVERT_LIMIT=N；不设 / 非法 = 不限，行为跟以前一样。
    let limitFromEnv () : int option =
        let text =
            defaultArg (Option.ofObj (Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_LIMIT")) ""

        match Int32.TryParse text with
        | true, n when n > 0 -> Some n
        | _ -> None

    /// 已知失败的旧文件清单：每行一个相对路径，# 开头是注释。
    /// 失败件原件保留，会永远排在队首；翻倍推进时把它们排除，否则后面的轮次全被它们挡住。
    let skipFromEnv (root: string) : Set<string> =
        match Option.ofObj (Environment.GetEnvironmentVariable "XUNITHOST_CONVERT_SKIP") with
        | None
        | Some "" -> Set.empty
        | Some file when not (File.Exists file) -> Set.empty
        | Some file ->
            File.ReadAllLines file
            |> Array.map(fun l -> l.Trim())
            |> Array.filter(fun l -> l <> "" && not (l.StartsWith "#"))
            |> Array.map(fun l -> Path.GetFullPath(Path.Combine(root, l)))
            |> Set.ofArray

    /// 稳定排序：Directory.GetFiles 不保证顺序，翻倍法取「前 N 份」必须可复现。
    let sortItems (items: Item list) : Item list =
        items |> List.sortBy(fun i -> i.Source.ToLowerInvariant())

    /// 本轮真正要动的：按类型 → 排除已知失败 → 稳定排序 → 取前 N 份。
    let selectForRound (root: string) (kind: Kind) (items: Item list) : Item list =
        let skipped = skipFromEnv root

        let ordered =
            items
            |> filterByKind kind
            |> List.filter(fun i -> not (skipped.Contains i.Source))
            |> sortItems

        match limitFromEnv () with
        | Some n -> ordered |> List.truncate n
        | None -> ordered

    /// 目标文件：没宏 → .docx/.xlsx/.pptx、有宏 → .docm/.xlsm/.pptm。
    /// 扫描阶段还没打开文件、不知道有没有宏，所以一律先按「没宏」占位；
    /// 真转的时候由 host 库里那六条转换各自定后缀。
    let targetOf (source: string) =
        if isWord source then Path.ChangeExtension(source, ".docx")
        elif isExcel source then Path.ChangeExtension(source, ".xlsx")
        else Path.ChangeExtension(source, ".pptx")

    /// 这一份旧文件对应的一对新格式后缀（宏版在前）
    let private newPair (source: string) =
        if isWord source then (".docm", ".docx")
        elif isExcel source then (".xlsm", ".xlsx")
        else (".pptm", ".pptx")

    /// 删原件时用这个：新文件到底落在哪个后缀上（有宏的是 .docm / .xlsm / .pptm）。
    /// 两个后缀都不在时按「没宏」的那个算，删原件前 `File.Exists` 还会再挡一道。
    let existingTargetOf (source: string) =
        let macro, plain = newPair source
        let m = Path.ChangeExtension(source, macro)
        if File.Exists m then m else Path.ChangeExtension(source, plain)

    /// 实际存在的同名新格式文件（可能两个都在）。改名让位时要一个不漏地全改。
    let existingTargetsOf (source: string) : string list =
        let macro, plain = newPair source
        [ Path.ChangeExtension(source, macro); Path.ChangeExtension(source, plain) ]
        |> List.filter File.Exists

    /// 重名让位：现有新文件改名成 `X-bak<戳>.docx`，扩展名不变、随时能改回来。
    [<Literal>]
    let RenameStamp = "202611061938"

    let renameOf (path: string) =
        Path.Combine(
            Path.GetDirectoryName path,
            Path.GetFileNameWithoutExtension path + "-bak" + RenameStamp + Path.GetExtension path
        )

    /// 递归扫一遍：谁要转、转到哪
    let scan (root: string) : Item list =
        let rec walk (dir: string) : Item list =
            let here =
                try
                    Directory.GetFiles dir
                    |> Array.filter isCandidate
                    |> Array.map(fun f -> { Source = f; Target = targetOf f })
                    |> Array.toList
                with _ ->
                    []

            let deeper =
                try
                    Directory.GetDirectories dir
                    |> Array.filter(skipDir >> not)
                    |> Array.map walk
                    |> List.concat
                with _ ->
                    []

            here @ deeper

        walk root
