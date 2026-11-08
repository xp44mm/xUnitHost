/// ExcelHost.Host.GetApplication(info: GeneralAppInfo) 的验证：旧路收束后唯一的 Excel 黑板入口。
/// 进程名常量 ExcelHost.Host.SelfImage 就在专用 Host 上（11-08 起没有种类标记，只剩进程名）。
/// 全部只读本机进程与 ROT，绝不新建 / 关闭 Word / Excel / PowerPoint / AutoCAD。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open OsHost

type ExcelGeneralTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Excel 通用参数表 GetApplication 过时信息拒接``() =
        Report.title output "Excel 通用参数表 GetApplication：过时信息拒接，不新建"

        let captured = Report.unixMs DateTime.Now
        let rot = "!ExcelVSTO.Instance.999999"

        let stale =
            GeneralAppInfo.create (ExcelHost.Host.SelfImage, 999999, rot, "", captured)

        Report.line output "入口进程名" ExcelHost.Host.SelfImage
        Report.line output "过时参数表" (stale.describe())

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        let old = ExcelHost.Host.GetApplication stale
        Report.line output "过时" (sprintf "%A" old)
        Should.equal true (match old with Error message -> message.Contains "绝不另起一个" | _ -> false)

        // 空的进程名比什么都拦得住：没有 image 就查不到进程，更不会去 ROT 里猜
        let noImage = { stale with Image = "" }
        let blank = ExcelHost.Host.GetApplication noImage
        Report.line output "缺进程名" (sprintf "%A" blank)
        Should.equal true (match blank with Error message -> message.Contains "Image" | _ -> false)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：进程已退报过时、缺进程名报缺 Image；四家进程前后一致，没起也没关"

    [<Fact>]
    member this.``Excel 通用参数表 GetApplication 按 ROT 名取回现成对象再认强类型``() =
        Report.title output "Excel 通用参数表 GetApplication：按 ROT 名挂回，认错类型就明确报错"

        let selfImage = Path.GetFileNameWithoutExtension Environment.ProcessPath
        let rot = sprintf "!General.%d" Environment.ProcessId
        let captured = Report.unixMs DateTime.Now
        let info = GeneralAppInfo.create (selfImage, Environment.ProcessId, rot, "", captured)
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

                let noImage = { info with Image = "" }
                let blank = ExcelHost.Host.GetApplication noImage
                Report.line output "缺进程名" (sprintf "%A" blank)
                Should.equal true (match blank with Error message -> message.Contains "Image" | _ -> false)
            finally
                Sta.run(fun () -> Com.revoke cookie)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：按 ROT 名取回的现成对象不是 Excel.Application 就明确报错；名字不在 / 缺进程名都拒接，四家进程没动"
