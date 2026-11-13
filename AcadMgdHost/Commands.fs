namespace AcadMgdHost

open System
open System.Reflection
open System.Runtime.Versioning
open Autodesk.AutoCAD.Runtime
open Autodesk.AutoCAD.DatabaseServices

/// 加载即执行兜底。
/// batchplot-plugin 的实测：控制台里 netload 能跑 Initialize，但 CommandMethod 可能不注册。
/// 所以 Initialize 先无条件报一行「确实加载了」，再按环境变量决定是否直接干活。
type Entry() =
    interface IExtensionApplication with
        member _.Initialize() =
            Probe.line "loaded" "IExtensionApplication.Initialize 被调用"

            if Environment.GetEnvironmentVariable "ACADMGDHOST_AUTORUN" = "1" then
                Probe.line "autorun" "Initialize 直接跑（不经过命令注册）"
                Probe.banner()
                Probe.database()

        member _.Terminate() = ()

/// 两个自定义命令：脚本按名字调用，能不能调起来就是「命令注册」这件事的判据。
type Commands() =

    [<CommandMethod("XH_HELLO")>]
    member _.Hello() =
        Probe.banner()

    [<CommandMethod("XH_DB")>]
    member _.Db() =
        Probe.database()

    /// 画一个五角星并另存到桌面：清空种子模型空间 → 闭合十顶点多段线 → SaveAs → 读回核对。
    [<CommandMethod("XH_STAR")>]
    member _.Star() =
        Star.run()
