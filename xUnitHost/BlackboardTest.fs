/// 黑板模式：插件把「当前实例」抄到黑板上，WorkBuddy 只读写、不负责黑板的创建与销毁。
/// 2026-11-13 起没有「通用参数表」了：挂回只认 PID，展示文本的参考解析在 BlackboardText.fs（只供调用方参考）。
/// 这里验证黑板的底层：ROT 按名字精确附着（多实例机器上唯一不会认错人的挂法）。
/// 文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求推迟（2026-11-07 13:51），暂不做。
/// 旧路 AppKind / AppInfo 已于 14:11 删除；11-08 又删掉 GeneralIdentity，四家只在 Host 上留一个 SelfImage。
/// 全部只读本机进程与 ROT，绝不新建 / 关闭 Word / Excel / PowerPoint / AutoCAD。
namespace XUnitHost

open System
open System.IO
open System.Runtime.InteropServices
open Xunit
open FSharp.xUnit
open OsHost

/// ROT 自测探针：本机进程里登记一个对象，不需要 Office 也能验「按名字取回」
[<ComVisible(true)>]
type BlackboardProbe() =
    member _.Tag = "黑板探针"

type BlackboardTest(output: ITestOutputHelper) =

    // 文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求推迟（2026-11-07 13:51）：
    // 这条 Fact 与两路库里的 checkDoc 一起撤掉了；要找回旧实现看 git 历史（6be1725 之前）。

    [<Fact>]
    member this.``ROT 按名字取回黑板登记的对象，名字不在就报 None``() =
        Report.title output "ROT 按名字：登记一个本地对象，按 moniker 名字精确取回"

        let name = sprintf "!XUnitHost.Blackboard.%d" Environment.ProcessId
        let probe = BlackboardProbe()

        let outcome =
            STA.here(fun () ->
                match Com.register name probe with
                | Error message -> Error message
                | Ok cookie ->
                    try
                        Ok(Com.runningAt name, Com.runningAt "!XUnitHost.Blackboard.不存在")
                    finally
                        Com.revoke cookie)

        match outcome with
        | Error message ->
            Report.line output "登记" message
            Assert.Skip("本机不允许在此进程登记 ROT：" + message)
        | Ok(found, missing) ->
            Report.line
                output
                "按名取回"
                (match found with
                 | Some value -> value.GetType().FullName
                 | None -> "None")

            Report.line
                output
                "不存在的名字"
                (match missing with
                 | Some value -> value.GetType().FullName
                 | None -> "None")

            Should.equal true found.IsSome
            Should.equal true missing.IsNone
            Report.conclusion output "通过：按 moniker 名字精确取回；名字不在就 None，绝不新建"

