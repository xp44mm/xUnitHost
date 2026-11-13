/// 文件：文档路径与「最后一次实质性保存」都从磁盘上问，不听内存里的标记。
[<RequireQualifiedAccess>]
module OsHost.Paths

open System
open System.Collections.Concurrent
open System.Diagnostics
open System.IO
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open System.Threading
open System.Threading.Tasks


/// 文档路径规整成绝对路径；新文档（没存过盘）返回 None。
/// Word 的新文档 FullName 是 ""，Excel 的是 "Book1"，AutoCAD 的是 "Drawing1.dwg"。
let canonical (raw: string) : string option =
    if String.IsNullOrWhiteSpace raw then
        None
    else
        try
            if Path.IsPathRooted raw then Some(Path.GetFullPath raw) else None
        with _ ->
            None

/// 「最后一次实质性保存」的判据：文档文件在磁盘上的最后写入时间。
/// 不用 Document.Saved 之类的内存标记 —— 那个只说明「有没有未保存改动」，
/// 说不出「上次保存是什么时候」；而磁盘时间正是保存真的落了盘的证据。
let savedMs (path: string option) : int64 =
    match path with
    | Some p when not (String.IsNullOrWhiteSpace p) ->
        try
            let file = FileInfo p
            if file.Exists then Clock.unixMs file.LastWriteTimeUtc else 0L
        with _ ->
            0L
    | _ -> 0L

