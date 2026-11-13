namespace OsHost

open System
open System.Globalization

[<RequireQualifiedAccess>]
module Rot =
    /// 注册名 → `(前缀, 进程号)`：`!WordVSTO.Instance.27232` → `Some("!WordVSTO.Instance.", 27232)`。
    /// 没有点、最后一段不是正整数、空串，一律 None —— 不猜。
    /// 前缀含末尾那个点，`build` 直接拼回去。
    let parse (rotName: string) : (string * int) option =
        if String.IsNullOrWhiteSpace rotName then
            None
        else
            let dot = rotName.LastIndexOf '.'
            if dot < 0 then
                None
            else
                match
                    Int32.TryParse(rotName.Substring(dot + 1), NumberStyles.Integer, CultureInfo.InvariantCulture)
                with
                | true, number when number > 0 -> Some(rotName.Substring(0, dot + 1), number)
                | _ -> None

    /// `(前缀, 进程号)` → 注册名：`"!WordVSTO.Instance."` 27232 → `Some "!WordVSTO.Instance.27232"`。
    /// 前缀空、pid ≤ 0 就是 None（不造一个假的出来）。
    let build (prefix: string) (pid: int) : string option =
        if String.IsNullOrWhiteSpace prefix || pid <= 0 then
            None
        else
            Some(prefix + string pid)
