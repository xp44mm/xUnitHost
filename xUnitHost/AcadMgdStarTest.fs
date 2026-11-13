namespace XUnitHost

open System
open System.Diagnostics
open System.IO
open System.Text
open Xunit
open FSharp.xUnit
open AcadMgdHost

/// 五角星：AcadMgdHost/XH_STAR 的纯几何自检 + accoreconsole 真机闭环。
/// 真机那条只起无头 accoreconsole，种子 DWG 复制到 %TEMP%，不碰用户开着的 acad.exe。
type AcadMgdStarTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``五角星 十顶点几何自检，不碰真机``() =
        Report.title output "五角星　十顶点几何（纯计算，不碰真机）"
        let points = StarGeometry.vertices 0.0 0.0 100.0

        Report.line output "顶点数" (string (List.length points))
        Should.equal 10 (List.length points)

        let (x0, y0) = List.head points
        Report.line output "第一点" (sprintf "(%.6f, %.6f)" x0 y0)
        Should.equal true (Math.Abs x0 < 1e-9 && Math.Abs(y0 - 100.0) < 1e-9)

        let radii = points |> List.mapi (fun i (x, y) -> i, Math.Sqrt(x * x + y * y))
        let outer = radii |> List.filter (fun (i, _) -> i % 2 = 0) |> List.map snd
        let inner = radii |> List.filter (fun (i, _) -> i % 2 = 1) |> List.map snd

        Report.line output "外半径" (sprintf "%.6f" (List.max outer))
        Report.line output "内半径" (sprintf "%.6f" (List.max inner))
        Report.line output "半径比" (sprintf "%.9f" StarGeometry.innerRatio)
        Should.equal true (outer |> List.forall (fun r -> Math.Abs(r - 100.0) < 1e-9))
        Should.equal true (inner |> List.forall (fun r -> Math.Abs(r - 100.0 * StarGeometry.innerRatio) < 1e-9))

        let perimeter = StarGeometry.perimeter points
        Report.line output "闭合周长" (sprintf "%.6f" perimeter)
        Should.equal true (perimeter > 0.0)
        Report.conclusion output "通过：10 点交替内外半径、首点在正上方、闭合周长可算"

    [<Fact>]
    member this.``五角星 画到桌面并读回核对``() =
        Report.title output "五角星　XH_STAR → 桌面 DWG → 读回核对"

        let root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."))
        let dll = Path.Combine(root, "AcadMgdHost", "bin", "Debug", "netstandard2.0", "AcadMgdHost.dll")
        let exe = @"C:\Program Files\Autodesk\AutoCAD 2020\accoreconsole.exe"
        let seedSource = @"C:\Program Files\Autodesk\AutoCAD 2020\Express\brkline.dwg"
        let target = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "五角星.dwg")

        Report.line output "插件" dll
        Report.line output "目标" target

        if not (File.Exists dll) then Assert.Skip("插件还没编译：" + dll)
        if not (File.Exists exe) then Assert.Skip("找不到 accoreconsole.exe：" + exe)
        if not (File.Exists seedSource) then Assert.Skip("找不到种子 DWG：" + seedSource)
        // 动笔前认人：桌面上已经有同名文件就停手，绝不覆盖已落盘的图纸
        if File.Exists target then Assert.Skip("桌面上已有「五角星.dwg」，不覆盖已落盘的文件")

        let work = Path.Combine(Path.GetTempPath(), "xUnitHost-star")
        Directory.CreateDirectory work |> ignore
        let scr = Path.Combine(work, "star.scr")
        let reportPath = Path.Combine(work, "star-report.txt")
        let logPath = Path.Combine(work, "star.log")
        let seed = Path.Combine(work, "seed.dwg")

        for path in [ scr; reportPath; logPath ] do
            if File.Exists path then File.Delete path

        File.Copy(seedSource, seed, true)
        File.SetAttributes(seed, FileAttributes.Normal)

        let dllPosix = dll.Replace('\\', '/')
        let script =
            [ "(setvar \"secureload\" 0)"
              "(setvar \"filedia\" 0)"
              sprintf "(command \"_.netload\" \"%s\")" dllPosix
              "XH_STAR"
              "_.quit _y" ]
            |> String.concat "\r\n"
        File.WriteAllText(scr, script + "\r\n", Encoding.ASCII)

        Report.line output "脚本" scr
        Report.line output "种子" seed

        let psi = ProcessStartInfo(exe)
        psi.Arguments <- sprintf "/i \"%s\" /s \"%s\"" seed scr
        psi.UseShellExecute <- false
        psi.RedirectStandardOutput <- true
        psi.RedirectStandardError <- true
        psi.StandardOutputEncoding <- Encoding.UTF8
        psi.StandardErrorEncoding <- Encoding.UTF8
        psi.EnvironmentVariables.["ACADMGDHOST_LOG"] <- logPath
        psi.EnvironmentVariables.["ACADMGDHOST_STAR_REPORT"] <- reportPath
        psi.EnvironmentVariables.["ACADMGDHOST_STAR_OUT"] <- target
        psi.EnvironmentVariables.["ACADMGDHOST_AUTORUN"] <- "0"

        let proc = Process.Start psi
        let outTask = proc.StandardOutput.ReadToEndAsync()
        let errTask = proc.StandardError.ReadToEndAsync()
        let exited = proc.WaitForExit(180000)

        if not exited then
            proc.Kill(true)

        let stdoutText = outTask.Result
        let stderrText = errTask.Result
        let exitCode = proc.ExitCode
        proc.Dispose()

        Report.line output "退出码" (string exitCode)
        Report.line output "stdout" (sprintf "%d 字符" stdoutText.Length)
        Report.line output "stderr" (if String.IsNullOrWhiteSpace stderrText then "(空)" else stderrText.Trim())
        Assert.True(exited, "accoreconsole 180 秒没退出")

        let rows =
            if File.Exists reportPath then
                File.ReadAllLines(reportPath, Encoding.UTF8)
                |> Array.choose (fun text ->
                    let at = text.IndexOf " = "
                    if at < 0 then None else Some(text.Substring(0, at).Trim(), text.Substring(at + 3).Trim()))
                |> Map.ofArray
            else
                Map.empty

        let value key =
            match Map.tryFind key rows with
            | Some v -> v
            | None -> "(缺失)"

        Report.line output "报告" reportPath
        Report.line output "status" (value "status")
        Report.line output "顶点" (value "vertices")
        Report.line output "闭合" (value "closed")
        Report.line output "周长" (value "perimeter")
        Report.line output "保存字节" (value "saved")
        Report.line output "DWG 头" (value "header")
        Report.line output "读回顶点" (value "readback_vertices")
        Report.line output "最大偏差" (value "readback_max_delta")
        Report.line output "错误" (value "error")

        Should.equal "OK" (value "status")
        Should.equal "10" (value "vertices")
        Should.equal "True" (value "closed")
        Should.equal "10" (value "readback_vertices")
        Should.equal "0" (value "readback_max_delta")

        Assert.True(File.Exists target, "桌面 DWG 没生成")
        let info = FileInfo target
        Report.line output "落盘" (sprintf "%s　%d 字节" info.FullName info.Length)
        Should.equal (Int64.Parse(value "saved")) info.Length

        use fs = File.OpenRead target
        let buffer = Array.zeroCreate<byte> 6
        fs.Read(buffer, 0, 6) |> ignore
        let header = Encoding.ASCII.GetString buffer
        Report.line output "文件头" header
        Should.equal (value "header") header

        let leftovers = Report.procIds "accoreconsole"
        Report.line output "现场" (if leftovers.IsEmpty then "无 accoreconsole 残留" else string leftovers)
        Report.conclusion output "通过：桌面 DWG 已生成，读回 10 顶点与画的时候逐点一致"
