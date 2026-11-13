/// ROT 前缀：通用层只做「切 / 拼」，前缀固化在各家项目里。
///
/// 离线跑，不碰 Office、不启任何进程（只把自己按 Word 的前缀登记一条再撤掉）：
///   xUnitHost.exe -method "*ROT 前缀*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.Diagnostics
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open OsHost

type RotTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``ROT 前缀：通用层切拼、各家固化前缀互不认错``() : Task =
        task {
            Report.title output "ROT　通用切 / 拼 + 各家前缀"

            // ① 通用层（OsHost.Rot）：只认「<前缀><pid>」这个形状，不碰 COM
            let cases =
                [ "!WordVSTO.Instance.27232", Some("!WordVSTO.Instance.", 27232)
                  "!ExcelVSTO.Instance.8", Some("!ExcelVSTO.Instance.", 8)
                  "!Foo", None
                  "!Foo.", None
                  "", None ]

            for name, expected in cases do
                let actual = OsHost.Rot.parse name
                Report.line2 output "parse" (sprintf "%A" name) (sprintf "%A" actual)
                Should.equal expected actual

            let built = OsHost.Rot.build "!WordVSTO.Instance." 27232
            Report.line2 output "build" "!WordVSTO.Instance. + 27232" (sprintf "%A" built)
            Should.equal (Some "!WordVSTO.Instance.27232") built
            Should.equal None (OsHost.Rot.build "!WordVSTO.Instance." 0)

            // ② 各家：前缀固化在项目里，别人的名字不认
            let word = WordHost.Rot.pidOf "!WordVSTO.Instance.27232"

            Report.line2 output "Word" "!WordVSTO.Instance.27232" (sprintf "%A" word)
            Should.equal (Some 27232) word
            Should.equal None (WordHost.Rot.pidOf "!ExcelVSTO.Instance.27232")
            Should.equal (Some "!WordVSTO.Instance.27232") (WordHost.Rot.nameOf 27232)

            let ppt = PowerPointHost.Rot.pidOf "!PowerPointVSTO.Instance.27232"

            Report.line2 output "PowerPoint" "!PowerPointVSTO.Instance.27232" (sprintf "%A" ppt)
            Should.equal (Some 27232) ppt
            Should.equal None (PowerPointHost.Rot.pidOf "!WordVSTO.Instance.27232")
            Should.equal (Some "!PowerPointVSTO.Instance.27232") (PowerPointHost.Rot.nameOf 27232)

            // ③ 端到端：把自己按 Word 的前缀登记到 ROT，再从枚举回来的名字认出 pid
            let me = Process.GetCurrentProcess().Id
            let wanted = WordHost.Rot.nameOf me

            Report.line2 output "本进程" (sprintf "pid %d" me) (sprintf "按 Word 前缀应叫 %A" wanted)
            Should.equal true wanted.IsSome

            let found =
                STA.here (fun () ->
                    let name = wanted.Value

                    match Com.register name (obj ()) with
                    | Error message ->
                        Report.line2 output "登记" name ("失败：" + message)
                        None
                    | Ok cookie ->
                        try
                            Report.line2 output "登记" name (sprintf "cookie=%d" cookie)

                            Com.runningEntries ()
                            |> List.tryFind(fun (entry, _) -> entry = name)
                            |> Option.bind(fun (entry, _) -> WordHost.Rot.pidOf entry)
                        finally
                            Com.revoke cookie)

            Report.line2 output "枚举回来" (sprintf "pid %d" me) (sprintf "%A" found)
            Should.equal (Some me) found

            Report.conclusion output "通过：通用层只切 / 拼，前缀由各家固化；别人的登记名不会被认成本家的"
        }
