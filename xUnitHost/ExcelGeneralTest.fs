/// ExcelHost.Host.GetApplication(info: GeneralAppInfo) 的验证：旧路收束后唯一的 Excel 黑板入口。
/// 身份常量 ExcelHost.Host.SelfIdentity 就在专用 Host 上；14:11 之前并存的 Host.General 子模块已撤掉。
/// 全部只读本机进程与 ROT，绝不新建 / 关闭 Word / Excel / PowerPoint / AutoCAD。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open OsHost

type ExcelGeneralTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Excel 通用参数表 GetApplication 错标记与过时信息都拒接``() =
        Report.title output "Excel 通用参数表 GetApplication：错标记 / 过时信息都拒接，不新建"

        let captured = Report.unixMs DateTime.Now
        let rot = "!ExcelVSTO.Instance.999999"

        let wrongTag =
            GeneralAppInfo.create (GeneralIdentity.create ("EXCEL", "Excel", "word"), 999999, rot, 0n, "", captured)

        let stale =
            GeneralAppInfo.create (
                GeneralIdentity.create ("EXCEL", "Excel", ExcelHost.Host.SelfIdentity.Tag),
                999999,
                rot,
                0n,
                "",
                captured
            )

        Report.line output "入口标记" ExcelHost.Host.SelfIdentity.Tag
        Report.line output "错标记表" (wrongTag.describe())
        Report.line output "过时参数表" (stale.describe())

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        let wrong = ExcelHost.Host.GetApplication wrongTag
        Report.line output "错标记" (sprintf "%A" wrong)
        Should.equal true (match wrong with Error message -> message.Contains "不接" | _ -> false)

        let old = ExcelHost.Host.GetApplication stale
        Report.line output "过时" (sprintf "%A" old)
        Should.equal true (match old with Error message -> message.Contains "绝不另起一个" | _ -> false)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：Tag 不对报不接、进程已退报过时；四家进程前后一致，没起也没关"

    [<Fact>]
    member this.``Excel 通用参数表 GetApplication 按 ROT 名取回现成对象再认强类型``() =
        Report.title output "Excel 通用参数表 GetApplication：按 ROT 名挂回，认错类型就明确报错"

        let selfImage = Path.GetFileNameWithoutExtension Environment.ProcessPath
        let rot = sprintf "!General.%d" Environment.ProcessId
        let captured = Report.unixMs DateTime.Now
        let identity = GeneralIdentity.create (selfImage, "测试进程", ExcelHost.Host.SelfIdentity.Tag)
        let info = GeneralAppInfo.create (identity, Environment.ProcessId, rot, 0n, "", captured)
        let probe = BlackboardProbe()

        Report.line output "参数表" (info.describe())
        Report.line output "探针类型" (probe.GetType().FullName)
        Report.line output "取法" "GeneralAppInfo.attach → Com.runningAt info.Rot；work = 强转成 Excel.Application + pid 核对"

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        // register / GetApplication / revoke 都排在同一条固定 STA 线程上：不跨公寓，也不嵌套 Sta.run
        let registered = Sta.run(fun () -> Com.register rot probe)

        match registered with
        | Error message ->
            Report.line output "登记" message
            Assert.Skip("本机不允许在此进程登记 ROT：" + message)
        | Ok cookie ->
            try
                let hit = ExcelHost.Host.GetApplication info
                let hitText =
                    match hit with
                    | Error message -> "Error " + message
                    | Ok _ -> "Ok (Excel.Application)"

                Report.line output "取回后认类型" hitText
                Should.equal true (match hit with Error message -> message.Length > 0 | _ -> false)
                // 取到了对象（不是「没拿到就绝不新建」），只是强类型认不出来
                Should.equal false (match hit with Error message -> message.Contains "绝不新建" | _ -> true)

                let missing = { info with Rot = "!General.不存在" }
                let absent = ExcelHost.Host.GetApplication missing
                Report.line output "名字不在" (sprintf "%A" absent)
                Should.equal true (match absent with Error message -> message.Contains "绝不新建" | _ -> false)

                let mismatch =
                    { info with Identity = GeneralIdentity.create (selfImage, "测试进程", "word") }

                let refused = ExcelHost.Host.GetApplication mismatch
                Report.line output "标记错配" (sprintf "%A" refused)
                Should.equal true (match refused with Error message -> message.Contains "不接" | _ -> false)
            finally
                Sta.run(fun () -> Com.revoke cookie)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：按 ROT 名取回的现成对象不是 Excel.Application 就明确报错，名字不在 / 标记错配都拒接，四家进程没动"
