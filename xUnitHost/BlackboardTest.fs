/// 黑板模式：插件把「当前实例」抄到黑板上，WorkBuddy 只读写、不负责黑板的创建与销毁。
/// 生产库只认结构化参数表 GeneralAppInfo；把展示文本填成参数表的参考实现在 BlackboardText.fs（只供调用方参考）。
/// 这里验证四家 GetApplication(info) 的入口：参数表、认人、按 ROT 名字精确附着、拒接过时的信息。
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

    [<Fact>]
    member this.``黑板模式 结构化参数表：参考解析把插件文本填成 GeneralAppInfo``() =
        Report.title output "黑板模式：库只认 GeneralAppInfo 参数表，文本解析是测试项目的参考实现"

        let wordText =
            "[WordVSTO 当前 Word 实例] 采集时间 2026-11-07 08:39:53
PID : 27232
ROT : !WordVSTO.Instance.27232
HWND : 0x002304FA
WINWORD.EXE : C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE
文档 : 文档1
路径 : (未保存)"

        let excelText =
            "[ExcelVSTO 当前 Excel 实例] 采集时间 2026-11-07 09:10:11
PID : 18860
ROT : !ExcelVSTO.Instance.18860
HWND : 0x001A0B2C
EXCEL.EXE : C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE
工作簿 : Book1
路径 : (未保存)"

        let pptText =
            "[黑板 当前实例] 采集时间 2026-11-07 09:20:12
KIND : ppt
PID : 4321
ROT : !PptVSTO.Instance.4321
HWND : 0x0000ABCD
POWERPNT.EXE : C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE
演示文稿 : 演示文稿1
路径 : (未保存)"

        let acadText =
            "[黑板 当前实例] 采集时间 2026-11-07 09:30:13
KIND : autocad
PID : 6550
ROT : !AcadVSTO.Instance.6550
HWND : 0x0001EF01
ACAD.EXE : C:\\Program Files\\Autodesk\\AutoCAD 2020\\acad.exe
图纸 : Drawing1.dwg
路径 : (未保存)"

        let parse (label: string) (text: string) =
            match BlackboardText.ofText text with
            | Ok info ->
                Report.line output label (info.describe())
                info
            | Error message -> failwith message

        let word = parse "Word" wordText
        let wordDoc = BlackboardText.docOfText wordText
        Report.line output "Word 标题" (BlackboardText.caption wordText)
        Report.line output "Word 文档" (wordDoc.describe())
        Should.equal WordHost.Host.SelfImage word.Image
        Should.equal 27232 word.Pid
        Should.equal "!WordVSTO.Instance.27232" word.Rot
        Should.equal (Some "文档1") wordDoc.Name
        Should.equal None wordDoc.Path

        // 库入口真正要的是这张参数表：手工按字段填报，与参考解析出来的那条应完全相同。
        let filled =
            GeneralAppInfo.create
                (WordHost.Host.SelfImage,
                 27232,
                 "!WordVSTO.Instance.27232",
                 @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")

        Report.line output "手工填报" (filled.describe())
        Should.equal word filled

        let excel = parse "Excel" excelText
        Should.equal ExcelHost.Host.SelfImage excel.Image
        Should.equal 18860 excel.Pid
        Should.equal "!ExcelVSTO.Instance.18860" excel.Rot
        Should.equal None (BlackboardText.docOfText excelText).Path

        let ppt = parse "PPT" pptText
        Should.equal PowerPointHost.Host.SelfImage ppt.Image
        Should.equal 4321 ppt.Pid
        Should.equal "!PptVSTO.Instance.4321" ppt.Rot

        let acad = parse "CAD" acadText
        Should.equal AutoCADHost.Host.SelfImage acad.Image
        Should.equal 6550 acad.Pid
        Should.equal (Some "Drawing1.dwg") (BlackboardText.docOfText acadText).Name

        // 采集时间行不进表：把标题里的采集时间抽掉，解析出来的参数表应当一模一样
        let noClockText = wordText.Replace(" 采集时间 2026-11-07 08:39:53", "")
        let noClock = BlackboardText.ofText noClockText
        Report.line output "抽掉采集时间" (sprintf "%A" noClock)
        Should.equal (Ok word) noClock

        let incomplete = BlackboardText.ofText "KIND : word"
        Report.line output "缺 PID" (sprintf "%A" incomplete)
        Should.equal true (match incomplete with Error _ -> true | _ -> false)

        Report.conclusion output "通过：四家文本由参考实现填成 GeneralAppInfo（只含进程信息），文档另取 GeneralDocInfo；手工填报同一条记录字段完全相同"

    [<Fact>]
    member this.``黑板模式 四家 GetApplication 拒接过时的信息，绝不新建``() =
        Report.title output "黑板模式：过时信息 → 四家 GetApplication 一律拒接"

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        let dead =
            GeneralAppInfo.create
                (WordHost.Host.SelfImage, 999999, "!WordVSTO.Instance.999999", "")

        Report.line output "输入" (dead.describe())

        let word = WordHost.Host.GetApplication dead
        let excel = ExcelHost.Host.GetApplication dead
        let ppt = PowerPointHost.Host.GetApplication dead
        let acad = AutoCADHost.Host.GetApplication dead

        let show (label: string) (result: Result<'T, string>) =
            Report.line
                output
                label
                (match result with
                 | Ok _ -> "Ok（不该发生）"
                 | Error message -> message)

        show "Word" word
        show "Excel" excel
        show "PPT" ppt
        show "CAD" acad

        // 表里只有进程身份：四家入口都会先核现场，同一张死掉的表对谁都是「已经不在运行」
        let messageOf (result: Result<'T, string>) =
            match result with
            | Error message -> message
            | Ok _ -> ""

        for message in [ messageOf word; messageOf excel; messageOf ppt; messageOf acad ] do
            Should.equal true (message.Contains "已经不在运行" && message.Contains WordHost.Host.SelfImage)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：四家都没起进程，过时的表对哪个入口都是「已经不在运行」"

    // 文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求推迟（2026-11-07 13:51）：
    // 这条 Fact 与两路库里的 checkDoc 一起撤掉了；要找回旧实现看 git 历史（6be1725 之前）。

    [<Fact>]
    member this.``ROT 按名字取回黑板登记的对象，名字不在就报 None``() =
        Report.title output "ROT 按名字：登记一个本地对象，按 moniker 名字精确取回"

        let name = sprintf "!XUnitHost.Blackboard.%d" Environment.ProcessId
        let probe = BlackboardProbe()

        let outcome =
            Sta.here(fun () ->
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

    [<Fact>]
    member this.``黑板模式 按参考解析与参数表定位真机：活着就挂上，退了就明确报过时``() =
        Report.title output "黑板模式：展示文本 → 参数表 → 四家 GetApplication（只读，不新建不退出）"

        let file = Environment.GetEnvironmentVariable "XUNITHOST_BLACKBOARD_FILE"

        if String.IsNullOrWhiteSpace file then
            Assert.Skip "没设 XUNITHOST_BLACKBOARD_FILE：把插件复制的黑板文本存成文件再跑"

        let text = File.ReadAllText file
        Report.line output "来源" "XUNITHOST_BLACKBOARD_FILE 的展示文本，由测试项目的参考实现转成参数表"

        match BlackboardText.ofText text with
        | Error message ->
            Report.line output "参考解析" message
            failwith message
        | Ok info ->
            Report.line output "参数表" (info.describe())
            Report.line output "文档" (BlackboardText.docOfText(text).describe())

            let image = info.Image
            let before = Report.procIds image
            Report.line output "现场" (sprintf "%s 进程 %A" image (Set.toList before))

            let result =
                if image = WordHost.Host.SelfImage then
                    WordHost.Host.GetApplication info |> Result.map(fun app -> box app)
                elif image = ExcelHost.Host.SelfImage then
                    ExcelHost.Host.GetApplication info |> Result.map(fun app -> box app)
                elif image = PowerPointHost.Host.SelfImage then
                    PowerPointHost.Host.GetApplication info |> Result.map(fun app -> box app)
                elif image = AutoCADHost.Host.SelfImage then
                    AutoCADHost.Host.GetApplication info |> Result.map(fun app -> box app)
                else
                    Error(sprintf "认不出 image=%s 的入口" image)

            Report.line
                output
                "结果"
                (match result with
                 | Ok value -> sprintf "挂上：%s" (value.GetType().FullName)
                 | Error message -> sprintf "拒接：%s" message)

            let after = Report.procIds image
            Report.line output "收尾" (sprintf "%s 进程 %A" image (Set.toList after))
            Should.equal before after

            if not (Set.contains info.Pid before) then
                Should.equal true (match result with Error _ -> true | _ -> false)

            Report.conclusion output "通过：按黑板信息核对现场；实例已退就明确拒接，全程没起也没关任何进程"
