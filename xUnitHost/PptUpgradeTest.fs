/// pptUpgrade 的验证：一条纯规则（后缀 / 格式码 / 选路），一条真机（造 .ppt → 升级 → 读回比对）。
/// 真机那条要自己从零起一个 PowerPoint；本机已经开着就跳过，绝不碰用户手上的演示文稿。
namespace XUnitHost

open System
open System.IO
open Xunit
open FSharp.xUnit
open Microsoft.Office.Core
open Microsoft.Office.Interop.PowerPoint
open OsHost

type PptUpgradeTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``pptUpgrade 规格：默认一律存 pptx，格式码与后缀各归其主``() =
        // 纯逻辑，不碰 PowerPoint：验证两条转换各自的规格与统一入口的选路
        let dir = Path.Combine(Path.GetTempPath(), "pptconvert-spec")
        let ppt = Path.Combine(dir, "样例.ppt")
        let pptx = PowerPointHost.PptToPptx.targetOf ppt
        let pptm = PowerPointHost.PptToPptm.targetOf ppt

        Report.title output "pptUpgrade 规格：默认一律 .pptx（宏不留）；XUNITHOST_PPT_KEEP_MACRO=1 才按宏分到 .pptm"

        Report.line output ".ppt 无宏 → pptx" (sprintf "%s　格式码 %d　PowerPointHost/PptToPptx" (Path.GetFileName pptx) (int PowerPointHost.PptToPptx.format))
        Report.line output ".ppt 有宏 → pptm" (sprintf "%s　格式码 %d　PowerPointHost/PptToPptm" (Path.GetFileName pptm) (int PowerPointHost.PptToPptm.format))
        Report.line output "统一入口" "pptUpgrade：默认直走 PptToPptx（宏只记录、不拦路）；只有 XUNITHOST_PPT_KEEP_MACRO=1 才判宏 → route 选路"

        Should.equal (Path.ChangeExtension(ppt, ".pptx")) pptx
        Should.equal (Path.ChangeExtension(ppt, ".pptm")) pptm
        Should.equal 24 (int PowerPointHost.PptToPptx.format)
        Should.equal 25 (int PowerPointHost.PptToPptm.format)
        Should.equal PowerPointHost.PptUpgrade.Pptx (PowerPointHost.PptUpgrade.route false)
        Should.equal PowerPointHost.PptUpgrade.Pptm (PowerPointHost.PptUpgrade.route true)

        Report.conclusion output "通过：两个格式码、两个后缀、两条选路规则都归在 PowerPointHost，互不串味"

    [<Fact>]
    member this.``pptUpgrade：一份 .ppt 升级成 .pptx，幻灯片文字一字不差，原件不动``() =
        Report.requireClean "POWERPNT"

        let beforeProc = Report.procIds "POWERPNT"
        let dir =
            Path.Combine(Path.GetTempPath(), "pptconvert-test-" + Guid.NewGuid().ToString("N").Substring(0, 8))

        Directory.CreateDirectory dir |> ignore

        let ppt = Path.Combine(dir, "样例.ppt")
        let body = "第一页：升级要一字不差\r第二页：含中文和数字 12345"

        try
            Report.title output "pptUpgrade：.ppt → .pptx，幻灯片文字一字不差，原件不动"

            let work (app: Application) =
                // ---- 造一份老格式 .ppt（一张空白版式 + 一个文本框）----
                let pres = app.Presentations.Add(MsoTriState.msoFalse)
                let slide = pres.Slides.Add(1, PpSlideLayout.ppLayoutBlank)
                let shape =
                    slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        50.0f,
                        50.0f,
                        400.0f,
                        100.0f
                    )

                shape.TextFrame.TextRange.Text <- body
                pres.SaveAs(ppt, PpSaveAsFileType.ppSaveAsPresentation, MsoTriState.msoTriStateMixed)
                let before = PowerPointHost.PptIO.slidePrint pres
                pres.Close()

                // ---- 升级：判宏 → 选路 → 转 → 里面已经读回比对过一次 ----
                let outcome = PowerPointHost.PptUpgrade.pptUpgrade app ppt
                let target, status, detail = outcome

                let after =
                    if status = "转了" || status = "已有·一致" then
                        PowerPointHost.PptIO.readPrint app target
                    else
                        ""

                // ---- 再来一次：已有同名 pptx，应该转到临时文件比完就删，不覆盖 ----
                let again = PowerPointHost.PptUpgrade.pptUpgrade app ppt

                // ---- 保宏那条路也真机走一遍（造带宏 .ppt 需要「信任 VBA 工程对象模型」，
                //      本机不一定开；这里至少证明格式码 25 与 .pptm 落盘、复核这条链是通的）----
                let macroOutcome = PowerPointHost.PptToPptm.convert app ppt
                let macroTarget, macroStatus, _ = macroOutcome

                let macroPrint =
                    if macroStatus = "转了" || macroStatus = "已有·一致" then
                        PowerPointHost.PptIO.readPrint app macroTarget
                    else
                        ""

                // ---- 负例：打不开的 .ppt（文件不存在）必须失败、原件保留 ----
                let missing = PowerPointHost.PptUpgrade.pptUpgrade app (Path.Combine(dir, "不存在.ppt"))

                ( before
                  , outcome
                  , after
                  , again
                  , macroOutcome
                  , macroPrint
                  , missing
                  , File.Exists ppt
                  , File.Exists(Path.ChangeExtension(ppt, ".pptx"))
                  , File.Exists(Path.ChangeExtension(ppt, ".pptm")) )

            let before, (target, status, detail), after, (_, againStatus, _), (macroTarget, macroStatus, macroDetail), macroPrint, (_, missingStatus, missingDetail), originalExists, pptxExists, pptmExists =
                PowerPointHost.Host.execQuery work |> Report.ok

            Report.line output "样例" (sprintf "%s（%d 字节）" (Path.GetFileName ppt) (FileInfo(ppt)).Length)
            Report.line output "升级" (sprintf "%s —— %s" status detail)
            Report.line output "落盘" (Path.GetFileName target)
            Report.line output "核对" (sprintf "幻灯片文字指纹 %d 字符，与原件逐字相同" after.Length)
            Report.line output "再来一次" (sprintf "%s（已有同名新文件，先转临时文件比完就删，不覆盖）" againStatus)
            Report.line output "保宏路" (sprintf "%s → %s　%s" (Path.GetFileName ppt) (Path.GetFileName macroTarget) macroDetail)
            Report.line output "打不开的" (sprintf "%s —— %s" missingStatus missingDetail)

            Should.equal "转了" status
            Assert.True(pptxExists, "没生成 pptx")
            Assert.True(originalExists, "转换不该动原件（删是另一条命令的事）")
            Should.equal before after
            Assert.Contains(PowerPointHost.PptIO.normalize body, after)
            Should.equal "已有·一致" againStatus
            Should.equal "转了" macroStatus
            Assert.True(pptmExists, "没生成 pptm")
            Should.equal before macroPrint
            Should.equal "失败" missingStatus
            Assert.True(missingDetail.Length > 0, "失败要给出原因")

            Report.conclusion output "通过：pptUpgrade 默认存 .pptx、转完读回比对、已有同名不覆盖、打不开就失败原件保留"
        finally
            // 兜底只清自己起的进程；用户手上的一只都不碰
            let leftovers = Set.difference (Report.procIds "POWERPNT") beforeProc
            if not leftovers.IsEmpty then Process.kill leftovers

            (try
                Directory.Delete(dir, true)
             with _ ->
                 ())
