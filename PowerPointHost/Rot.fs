/// PowerPoint 的 ROT 注册名：PowerPointVSTO 插件把当前实例登记成 `!PowerPointVSTO.Instance.<pid>`。
///
/// 前缀只在这一家成立，所以固化在这儿；通用的切 / 拼在 `OsHost.Rot`（纯字符串、不碰 COM）。
/// 用途：`Host.findOnSta` 扫 ROT 时先看注册名里的 pid，再看 HWND → pid。
namespace PowerPointHost

[<RequireQualifiedAccess>]
module Rot =

    /// 名字前缀（含末尾的点）。要改就改这一处。
    [<Literal>]
    let Prefix = "!PowerPointVSTO.Instance."

    /// 注册名 → 进程号：`!PowerPointVSTO.Instance.27232` → `Some 27232`；不是这个前缀就 None。
    let pidOf (rotName: string) : int option =
        OsHost.Rot.parse rotName
        |> Option.filter (fun (prefix, _) -> prefix = Prefix)
        |> Option.map snd

    /// 进程号 → 注册名：27232 → `Some "!PowerPointVSTO.Instance.27232"`。
    let nameOf (pid: int) : string option = OsHost.Rot.build Prefix pid
