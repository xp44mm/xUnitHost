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
