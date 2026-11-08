/// 通用参数表 GeneralAppInfo 的验证：身份由各家专用 Host 的参数给出，认人 / 校验 / 附着还灵不灵。
/// 旧路 AppInfo / AppKind 已删除；这里的三条 Fact 全部只读本机进程与 ROT，
/// 绝不新建 / 关闭 Word / Excel / PowerPoint / AutoCAD。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open OsHost

type GeneralAppInfoTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``通用参数表 进程名常量在四家专用 Host，第五家不用改库``() =
        Report.title output "通用参数表：进程名常量在专用 Host，第五家不用改库"

        // 原来 AppKind.image() 的四个分支，现在各归各家 Host 的 SelfImage；识别种类没有第二个字段
        for label, image in
            [ "Word", WordHost.Host.SelfImage
              "Excel", ExcelHost.Host.SelfImage
              "PPT", PowerPointHost.Host.SelfImage
              "CAD", AutoCADHost.Host.SelfImage
              "第五家", "et" ] do
            Report.line2 output "进程名" label image

        Should.equal "WINWORD" WordHost.Host.SelfImage
        Should.equal "EXCEL" ExcelHost.Host.SelfImage
        Should.equal "POWERPNT" PowerPointHost.Host.SelfImage
        Should.equal "acad" AutoCADHost.Host.SelfImage

        let captured = Report.unixMs DateTime.Now
        let info = GeneralAppInfo.create ("et", 1234, "!WpsVSTO.Instance.1234", "", captured)
        Report.line output "第五家参数表" (info.describe())
        Should.equal "et" info.Image
        // 表里只有「找到这只进程」要用的现场值：没有种类标记、没有显示名
        Report.line output "字段" "Image / Pid / Rot / ExePath / CapturedMs（没有窗口句柄、没有种类标记）"

        Report.conclusion output "通过：进程名是 Host 上的一个字符串，库里没有封闭种类、没有标记字段；第五家直接填 image 就能用"

    [<Fact>]
    member this.``通用参数表 获取函数占位：注入 acquire 挂到登记对象，不需要 Office``() =
        Report.title output "通用参数表：获取函数占位，注入 acquire 就能挂到现成对象"

        let selfImage = Path.GetFileNameWithoutExtension Environment.ProcessPath
        let name = sprintf "!General.%d" Environment.ProcessId
        let probe = BlackboardProbe()
        let captured = Report.unixMs DateTime.Now
        let info = GeneralAppInfo.create (selfImage, Environment.ProcessId, name, "", captured)

        Report.line output "参数表" (info.describe())
        Report.line output "获取函数" "acquire = Com.runningAt info.Rot；work = 取回对象的类型名"

        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        // register / acquire / revoke 都排到同一条固定 STA 线程上：避免跨公寓，也避免 Sta.run 嵌套
        let registered = Sta.run(fun () -> Com.register name probe)

        match registered with
        | Error message ->
            Report.line output "登记" message
            Assert.Skip("本机不允许在此进程登记 ROT：" + message)
        | Ok cookie ->
            try
                let acquire (i: GeneralAppInfo) : obj option = Com.runningAt i.Rot
                let work (raw: obj) : Result<string, string> = Ok(raw.GetType().FullName)

                let hit = GeneralAppInfo.attachWith info acquire work
                Report.line output "按名挂回" (sprintf "%A" hit)
                Should.equal (Ok "XUnitHost.BlackboardProbe") hit

                let missing = { info with Rot = "!General.不存在" }
                let absent = GeneralAppInfo.attachWith missing acquire work
                Report.line output "名字不在" (sprintf "%A" absent)
                Should.equal true (match absent with Error message -> message.Contains "绝不新建" | _ -> false)

                let noImage = GeneralAppInfo.attachWith { info with Image = "" } acquire work
                Report.line output "缺进程名" (sprintf "%A" noImage)
                Should.equal true (match noImage with Error message -> message.Contains "Image" | _ -> false)
            finally
                Sta.run(fun () -> Com.revoke cookie)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：acquire 由调用方注入，按 ROT 名正向挂回；取不到 / 缺进程名都明确拒接，四家进程没动"

    [<Fact>]
    member this.``通用参数表 四家 Host 各用自己的 SelfImage，过时信息一律拒接``() =
        Report.title output "通用参数表：四家 Host 各用自己的 SelfImage，过时信息一律拒接"

        let captured = Report.unixMs DateTime.Now
        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        let entries: (string * string * (GeneralAppInfo -> Result<obj, string>)) list =
            [ "Word", WordHost.Host.SelfImage,
              (fun i -> WordHost.Host.GetApplication i |> Result.map(fun app -> box app))
              "Excel", ExcelHost.Host.SelfImage,
              (fun i -> ExcelHost.Host.GetApplication i |> Result.map(fun app -> box app))
              "PPT", PowerPointHost.Host.SelfImage,
              (fun i -> PowerPointHost.Host.GetApplication i |> Result.map(fun app -> box app))
              "CAD", AutoCADHost.Host.SelfImage,
              (fun i -> AutoCADHost.Host.GetApplication i |> Result.map(fun app -> box app)) ]

        for label, image, call in entries do
            let info = GeneralAppInfo.create (image, 999999, "!self.Instance.999999", "", captured)

            Report.line output label (info.describe())

            let result = call info
            let text =
                match result with
                | Error message -> message
                | Ok _ -> "Ok（不该发生）"

            Report.line2 output "拒接" label text
            Should.equal true (
                match result with
                | Error message -> message.Contains image && message.Contains "绝不另起一个"
                | _ -> false
            )

        // 文档信息与认进程的表分开：GeneralAppInfo 里没有文档字段，文档另立一条
        let doc = GeneralDocInfo.create (Some "文档1", Some @"C:\tmp\文档1.docx")
        Report.line output "文档信息" (doc.describe())
        Should.equal (Some "文档1") doc.Name
        Should.equal (Some @"C:\tmp\文档1.docx") doc.Path
        Should.equal "名字=(无) 路径=(无)" (GeneralDocInfo.empty.describe())

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：四家进程名常量都来自专用 Host，过时信息四家入口都拒接；文档信息不在认进程的表里"
