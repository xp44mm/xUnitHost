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
    member this.``通用参数表 身份常量在四家专用 Host，第五家不用改库``() =
        Report.title output "通用参数表：四家身份常量在专用 Host，第五家不用改库"

        // 原来 AppKind.image() / kindName() 四个分支的值，现在各归各家 Host 的 SelfIdentity
        let word = WordHost.Host.SelfIdentity
        let excel = ExcelHost.Host.SelfIdentity
        let ppt = PowerPointHost.Host.SelfIdentity
        let acad = AutoCADHost.Host.SelfIdentity
        let fifth = GeneralIdentity.create ("et", "WPS 表格", "et")

        for identity in [ word; excel; ppt; acad; fifth ] do
            Report.line output identity.Tag (identity.describe())

        Should.equal "WINWORD" word.Image
        Should.equal "Word" word.Name
        Should.equal "EXCEL" excel.Image
        Should.equal "POWERPNT" ppt.Image
        Should.equal "acad" acad.Tag
        Should.equal "et" fifth.Image

        let captured = Report.unixMs DateTime.Now
        let info = GeneralAppInfo.create (fifth, 1234, "!WpsVSTO.Instance.1234", 0n, "", captured)
        Report.line output "第五家参数表" (info.describe())
        Should.equal "WPS 表格" info.Identity.Name

        let right = GeneralAppInfo.ensureTag "et" info
        let wrong = GeneralAppInfo.ensureTag "word" info
        Report.line output "同标记" (sprintf "%A" right)
        Report.line output "错标记" (sprintf "%A" wrong)
        Should.equal true (match right with Ok () -> true | _ -> false)
        Should.equal true (match wrong with Error message -> message.Contains "不接" | _ -> false)

        Report.conclusion output "通过：四家 + 第五家的身份都是 Host 上的三个参数，库里没有封闭种类；标记错配就拒接"

    [<Fact>]
    member this.``通用参数表 获取函数占位：注入 acquire 挂到登记对象，不需要 Office``() =
        Report.title output "通用参数表：获取函数占位，注入 acquire 就能挂到现成对象"

        let selfImage = Path.GetFileNameWithoutExtension Environment.ProcessPath
        let name = sprintf "!General.%d" Environment.ProcessId
        let probe = BlackboardProbe()
        let captured = Report.unixMs DateTime.Now
        let identity = GeneralIdentity.create (selfImage, "测试进程", "probe")
        let info = GeneralAppInfo.create (identity, Environment.ProcessId, name, 0n, "", captured)

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

                let hit = GeneralAppInfo.attachWith info "probe" acquire work
                Report.line output "按名挂回" (sprintf "%A" hit)
                Should.equal (Ok "XUnitHost.BlackboardProbe") hit

                let missing = { info with Rot = "!General.不存在" }
                let absent = GeneralAppInfo.attachWith missing "probe" acquire work
                Report.line output "名字不在" (sprintf "%A" absent)
                Should.equal true (match absent with Error message -> message.Contains "绝不新建" | _ -> false)

                let mismatch = GeneralAppInfo.attachWith info "word" acquire work
                Report.line output "标记错配" (sprintf "%A" mismatch)
                Should.equal true (match mismatch with Error message -> message.Contains "不接" | _ -> false)
            finally
                Sta.run(fun () -> Com.revoke cookie)

        for image, ids in before do
            let now = Report.procIds image
            Report.line2 output "进程" image (sprintf "%A → %A" (Set.toList ids) (Set.toList now))
            Should.equal ids now

        Report.conclusion output "通过：acquire 由调用方注入，按 ROT 名正向挂回；取不到 / 标记错配都明确拒接，四家进程没动"

    [<Fact>]
    member this.``通用参数表 四家 Host 各用自己的 SelfIdentity，过时信息一律拒接``() =
        Report.title output "通用参数表：四家 Host 各用自己的 SelfIdentity，过时信息一律拒接"

        let captured = Report.unixMs DateTime.Now
        let images = [ "WINWORD"; "EXCEL"; "POWERPNT"; "acad" ]
        let before = images |> List.map(fun image -> image, Report.procIds image)

        let entries: (string * GeneralIdentity * (GeneralAppInfo -> Result<obj, string>)) list =
            [ "Word", WordHost.Host.SelfIdentity,
              (fun i -> WordHost.Host.GetApplication i |> Result.map(fun app -> box app))
              "Excel", ExcelHost.Host.SelfIdentity,
              (fun i -> ExcelHost.Host.GetApplication i |> Result.map(fun app -> box app))
              "PPT", PowerPointHost.Host.SelfIdentity,
              (fun i -> PowerPointHost.Host.GetApplication i |> Result.map(fun app -> box app))
              "CAD", AutoCADHost.Host.SelfIdentity,
              (fun i -> AutoCADHost.Host.GetApplication i |> Result.map(fun app -> box app)) ]

        for label, identity, call in entries do
            let info =
                GeneralAppInfo.create
                    (identity, 999999, sprintf "!%s.Instance.999999" identity.Tag, 0n, "", captured)

            Report.line output label (info.describe())

            let result = call info
            let text =
                match result with
                | Error message -> message
                | Ok _ -> "Ok（不该发生）"

            Report.line2 output "拒接" label text
            Should.equal true (
                match result with
                | Error message -> message.Contains identity.Image && message.Contains "绝不另起一个"
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

        Report.conclusion output "通过：四家身份常量都来自专用 Host，过时信息四家入口都拒接、进程名各归各家；文档信息不在认进程的表里"
