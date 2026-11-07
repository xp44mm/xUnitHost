namespace AcadMgdHost

open System
open System.Reflection
open System.Runtime.Versioning
open Autodesk.AutoCAD.Runtime
open Autodesk.AutoCAD.DatabaseServices

/// 只给 Core Console / accoreconsole.exe 用：引用面只有 acdbmgd.dll + accoremgd.dll。
/// 输出一律走 Console.WriteLine —— accoreconsole 是控制台进程，
/// Editor.WriteMessage 和任何 UI 类型都不可用（acmgd.dll 也绝不能引）。
module internal Probe =

    let private tag = "[AcadMgdHost]"

    /// accoreconsole 会把 Console.Out 重定向到 %TEMP%\accc* 临时文件，外层 stdout 里看不到。
    /// 所以额外往 ACADMGDHOST_LOG 指定的文件写一份，脚本 / Fact 才有稳定可读的回信。
    let private logPath =
        match Environment.GetEnvironmentVariable "ACADMGDHOST_LOG" with
        | null | "" -> None
        | path -> Some path

    let line (label: string) (value: string) =
        let text = sprintf "%s %-10s %s" tag label value
        Console.WriteLine text

        match logPath with
        | Some path ->
            try
                System.IO.File.AppendAllText(path, text + Environment.NewLine, System.Text.Encoding.Default)
            with _ ->
                ()
        | None -> ()

    /// 这几行是「程序集真的被 .NET Framework 进程加载了」的证据：
    /// 程序集名 / 版本 / 编译时的目标框架 / 宿主进程 / CLR 版本。
    let banner () =
        let asm = Assembly.GetExecutingAssembly()

        let tfm =
            match asm.GetCustomAttributes(typeof<TargetFrameworkAttribute>, false) with
            | [| :? TargetFrameworkAttribute as attr |] -> attr.FrameworkName
            | _ -> "(unknown)"

        let name = asm.GetName()
        line "assembly" (name.Name + " " + string name.Version)
        line "framework" tfm
        line "process" (Diagnostics.Process.GetCurrentProcess().ProcessName)
        line "clr" (string Environment.Version)

    /// 证明 acdbmgd 的数据库 API 在无头进程里可用：当前图纸 + 模型空间实体数。
    /// 取不到就照实打印异常，不吞掉，也不让命令整个崩掉。
    let database () =
        try
            let db = HostApplicationServices.WorkingDatabase

            if isNull db then
                line "database" "WorkingDatabase = null"
            else
                line "dwg" (if String.IsNullOrEmpty db.Filename then "(未命名)" else db.Filename)

                use tr = db.TransactionManager.StartTransaction()
                let bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) :?> BlockTable
                let ms = tr.GetObject(bt.[BlockTableRecord.ModelSpace], OpenMode.ForRead) :?> BlockTableRecord
                let mutable count = 0

                for _ in ms do
                    count <- count + 1

                tr.Commit()
                line "modelspace" (sprintf "%d 个实体" count)
        with ex ->
            line "database" ("读不了：" + ex.GetType().Name + " " + ex.Message)

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
