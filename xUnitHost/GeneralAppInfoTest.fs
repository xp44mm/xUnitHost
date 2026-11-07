/// 通用参数表 GeneralAppInfo 的验证：与 AppKind 解耦之后，认人 / 校验 / 附着还灵不灵。
/// 这三条 Fact 全部只读本机进程与 ROT，绝不新建 / 关闭 Word / Excel / PowerPoint / AutoCAD。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open OsHost

/// 调用方适配示例（**不是库 API**）：把老参数表里的封闭种类翻成通用身份参数。
/// 生产层 GeneralAppInfo 不知道 AppKind 存在；要不要从旧表搬过来，是调用方自己的事。
module internal GeneralBridge =

    let identityOf (kind: AppKind) : GeneralIdentity =
        let tag =
            match kind with
            | AppKind.Word -> "word"
            | AppKind.Excel -> "excel"
            | AppKind.PowerPoint -> "ppt"
            | AppKind.AutoCAD -> "acad"

        GeneralIdentity.create (kind.image(), kind.kindName(), tag)

    let ofAppInfo (info: AppInfo) : GeneralAppInfo =
        { (GeneralAppInfo.create
            (identityOf info.Kind, info.Pid, info.Rot, info.Hwnd, info.ExePath, info.CapturedMs))
          with
              DocName = info.DocName
              DocPath = info.DocPath }

type GeneralAppInfoTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``通用参数表 身份由参数填，第五家不用改库``() =
        Report.title output "通用参数表：身份由参数填，不认 AppKind"

        let word = GeneralIdentity.create ("WINWORD", "Word", "word")
        let excel = GeneralIdentity.create ("EXCEL", "Excel", "excel")
        let ppt = GeneralIdentity.create ("POWERPNT", "PowerPoint", "ppt")
        let acad = GeneralIdentity.create ("acad", "AutoCAD", "acad")
        let fifth = GeneralIdentity.create ("et", "WPS 表格", "et")

        for identity in [ word; excel; ppt; acad; fifth ] do
            Report.line output identity.Tag (identity.describe())

        Should.equal "WINWORD" word.Image
        Should.equal "Word" word.Name
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

        Report.conclusion output "通过：四家 + 第五家的身份都是三个参数，库里没有 AppKind；标记错配就拒接"

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
    member this.``通用参数表 与 AppInfo 的现场校验与文档校验同口径``() =
        Report.title output "通用参数表：与 AppInfo 逐项对照，校验口径没有丢"

        let captured = Report.unixMs DateTime.Now
        let oldInfo = AppInfo.create (AppKind.Word, 999999, "!WordVSTO.Instance.999999", 0n, "", captured)
        let newInfo = GeneralBridge.ofAppInfo oldInfo

        Report.line output "旧参数表" (oldInfo.describe())
        Report.line output "新参数表" (newInfo.describe())
        Should.equal "WINWORD" newInfo.Identity.Image

        let oldVerify = AppInfo.verify oldInfo
        let newVerify = GeneralAppInfo.verify newInfo
        Report.line output "旧现场校验" (sprintf "%A" oldVerify)
        Report.line output "新现场校验" (sprintf "%A" newVerify)
        Should.equal oldVerify newVerify

        let path =
            Path.Combine(Path.GetTempPath(), sprintf "xunithost-general-%d.txt" Environment.ProcessId)

        File.WriteAllText(path, "通用参数表")

        try
            let old = { oldInfo with DocName = Some "文档1"; DocPath = Some path }
            let gen = GeneralBridge.ofAppInfo old

            let pair (label: string) (currentName: string option) (currentPath: string option) =
                let a = AppInfo.checkDoc old currentName currentPath
                let b = GeneralAppInfo.checkDoc gen currentName currentPath
                Report.line output label (sprintf "%A | %A" a b)
                Should.equal a b

            File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddHours(-1.0))
            pair "一致" (Some "文档1") (Some path)
            pair "换了文档" (Some "另一份.docx") (Some path)
            pair "文档关了" None None

            File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddHours(1.0))
            pair "存过盘" (Some "文档1") (Some path)

            Report.conclusion output "通过：同一条信息两张表，现场校验与四种文档情形逐项一致；旧表一行未改"
        finally
            File.Delete path
