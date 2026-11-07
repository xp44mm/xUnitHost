/// 黑板模式：插件把「当前实例」抄到黑板上，WorkBuddy 只读写、不负责黑板的创建与销毁。
/// 这里验证四家 GetApplication(info) 的入口：解析、认人、按 ROT 名字精确附着、拒接过时的信息。
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
    member this.``黑板信息 解析：插件剪贴板文本认得四家``() =
        Report.title output "黑板信息解析：四家插件文本 → AppInfo"

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
            match AppInfo.ofText text with
            | Ok info ->
                Report.line output label (AppInfo.describe info)
                info
            | Error message -> failwith message

        let word = parse "Word" wordText
        Should.equal AppKind.Word word.Kind
        Should.equal 27232 word.Pid
        Should.equal "!WordVSTO.Instance.27232" word.Rot
        Should.equal (nativeint 0x002304FA) word.Hwnd
        Should.equal (Some "文档1") word.DocName
        Should.equal None word.DocPath

        let excel = parse "Excel" excelText
        Should.equal AppKind.Excel excel.Kind
        Should.equal 18860 excel.Pid
        Should.equal "!ExcelVSTO.Instance.18860" excel.Rot
        Should.equal None excel.DocPath

        let ppt = parse "PPT" pptText
        Should.equal AppKind.PowerPoint ppt.Kind
        Should.equal 4321 ppt.Pid
        Should.equal "!PptVSTO.Instance.4321" ppt.Rot

        let acad = parse "CAD" acadText
        Should.equal AppKind.AutoCAD acad.Kind
        Should.equal 6550 acad.Pid
        Should.equal (Some "Drawing1.dwg") acad.DocName

        Report.line output "Word 采集" (Report.clock word.CapturedMs)
        Should.equal true (word.CapturedMs > 0L)

        let incomplete = AppInfo.ofText "KIND : word\nPID : 123"
        Report.line output "缺采集时间" (sprintf "%A" incomplete)
        Should.equal true (match incomplete with Error _ -> true | _ -> false)

        Report.conclusion output "通过：四家文本各认各的，PID / ROT / HWND / 采集时间都读得出，缺采集时间就拒收"

    [<Fact>]
    member this.``黑板模式 四家 GetApplication 拒接过时的信息，绝不新建``() =
        Report.title output "黑板模式：过期 / 错配信息 → 四家 GetApplication 一律拒接"

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        let captured = Report.unixMs DateTime.Now
        let dead = AppInfo.ofParts AppKind.Word 999999 "!WordVSTO.Instance.999999" 0n "" captured
        Report.line output "输入" (AppInfo.describe dead)

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

        Should.equal true (match word with Error message -> message.Contains "已经不在运行" | _ -> false)
        Should.equal true (match excel with Error message -> message.Contains "不接" | _ -> false)
        Should.equal true (match ppt with Error message -> message.Contains "不接" | _ -> false)
        Should.equal true (match acad with Error message -> message.Contains "不接" | _ -> false)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：四家都没起进程，过时报过时、错配报不接"

    [<Fact>]
    member this.``黑板模式 文档级校验：换了文档或采集后被保存就拒接``() =
        Report.title output "黑板模式：附着后核对文档，保证数据完整不过时"

        let path =
            Path.Combine(Path.GetTempPath(), sprintf "xunithost-blackboard-%d.txt" Environment.ProcessId)

        File.WriteAllText(path, "黑板")

        try
            let captured = Report.unixMs DateTime.Now
            let baseInfo = AppInfo.ofParts AppKind.Word 1 "" 0n "" captured
            let info = { baseInfo with DocName = Some "文档1"; DocPath = Some path }

            // 文件最后写入时间早于采集时刻：没被别人动过
            File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddHours(-1.0))
            let aligned = AppInfo.checkDoc info (Some "文档1") (Some path)
            Report.line output "一致" (sprintf "%A" aligned)
            Should.equal true (match aligned with Ok _ -> true | _ -> false)

            let changed = AppInfo.checkDoc info (Some "另一份.docx") (Some path)
            Report.line output "换了文档" (sprintf "%A" changed)
            Should.equal true (match changed with Error _ -> true | _ -> false)

            let closed = AppInfo.checkDoc info None None
            Report.line output "文档关了" (sprintf "%A" closed)
            Should.equal true (match closed with Error _ -> true | _ -> false)

            // 采集之后又存了一次盘：黑板信息已经过时
            File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddHours(1.0))
            let saved = AppInfo.checkDoc info (Some "文档1") (Some path)
            Report.line output "存过盘" (sprintf "%A" saved)
            Should.equal true (match saved with Error _ -> true | _ -> false)

            Report.conclusion output "通过：名字 / 路径一致才放行，换了文档或采集后存过盘就拒接"
        finally
            File.Delete path

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
    member this.``黑板模式 按黑板文本定位真机：活着就挂上，退了就明确报过时``() =
        Report.title output "黑板模式：按剪贴板文本定位真机实例（只读，不新建不退出）"

        let file = Environment.GetEnvironmentVariable "XUNITHOST_BLACKBOARD_FILE"

        if String.IsNullOrWhiteSpace file then
            Assert.Skip "没设 XUNITHOST_BLACKBOARD_FILE：把插件复制的黑板文本存成文件再跑"

        let text = File.ReadAllText file

        match AppInfo.ofText text with
        | Error message ->
            Report.line output "解析" message
            failwith message
        | Ok info ->
            Report.line output "黑板" (AppInfo.describe info)
            Report.line output "文档" (sprintf "名字=%A 路径=%A" info.DocName info.DocPath)

            let image = AppInfo.image info.Kind
            let before = Report.procIds image
            Report.line output "现场" (sprintf "%s 进程 %A" image (Set.toList before))

            let result =
                match info.Kind with
                | AppKind.Word -> WordHost.Host.GetApplication info |> Result.map(fun app -> box app)
                | AppKind.Excel -> ExcelHost.Host.GetApplication info |> Result.map(fun app -> box app)
                | AppKind.PowerPoint -> PowerPointHost.Host.GetApplication info |> Result.map(fun app -> box app)
                | AppKind.AutoCAD -> AutoCADHost.Host.GetApplication info |> Result.map(fun app -> box app)

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
