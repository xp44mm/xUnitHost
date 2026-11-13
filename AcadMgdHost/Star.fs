namespace AcadMgdHost

open System
open System.IO
open System.Text
open Autodesk.AutoCAD.DatabaseServices
open Autodesk.AutoCAD.Geometry

/// 在 accoreconsole 的 WorkingDatabase 里画一个五角星，另存到桌面，
/// 再从磁盘重新打开、逐点读回核对（和格式转换「转一份、读回比一份」同一口径）。
module internal Star =

    let private centerX = 0.0
    let private centerY = 0.0
    let private radius = 100.0

    /// 默认落在桌面「五角星.dwg」；ACADMGDHOST_STAR_OUT 可覆盖。
    let outputPath () =
        match Environment.GetEnvironmentVariable "ACADMGDHOST_STAR_OUT" with
        | null | "" ->
            let desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Path.Combine(desktop, "五角星.dwg")
        | path -> path

    let private reportPath () =
        match Environment.GetEnvironmentVariable "ACADMGDHOST_STAR_REPORT" with
        | null | "" -> None
        | path -> Some path

    /// 单独写一份 UTF-8 报告：Probe.line 那份日志走系统 ANSI，外层读还要猜编码。
    let private writeReport (rows: (string * string) list) =
        match reportPath () with
        | None -> ()
        | Some path ->
            let text =
                rows
                |> List.map (fun (k, v) -> k + " = " + v)
                |> String.concat "\r\n"
            File.WriteAllText(path, text + "\r\n", UTF8Encoding(false))

    let private say (key: string) (text: string) = Probe.line key text

    let private headerOf (path: string) =
        use fs = File.OpenRead path
        let buffer = Array.zeroCreate<byte> 6
        let read = fs.Read(buffer, 0, 6)
        Encoding.ASCII.GetString(buffer, 0, read)

    /// 种子 DWG 只是 /i 的容器（副本在 %TEMP%）：先清空模型空间，再画一条闭合十顶点多段线。
    let private drawInto (db: Database) =
        use tr = db.TransactionManager.StartTransaction()
        let bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) :?> BlockTable
        let ms = tr.GetObject(bt.[BlockTableRecord.ModelSpace], OpenMode.ForWrite) :?> BlockTableRecord

        let existing = ms |> Seq.cast<ObjectId> |> Seq.toArray

        for oid in existing do
            let ent = tr.GetObject(oid, OpenMode.ForWrite) :?> Entity
            ent.Erase()

        let points = StarGeometry.vertices centerX centerY radius
        let pl = new Polyline()

        points
        |> List.iteri (fun i (x, y) -> pl.AddVertexAt(i, Point2d(x, y), 0.0, 0.0, 0.0))

        pl.Closed <- true
        pl.ColorIndex <- 1
        ms.AppendEntity(pl) |> ignore
        tr.AddNewlyCreatedDBObject(pl, true)
        tr.Commit()
        (existing.Length, points)

    /// 从磁盘另开一个 Database 读回，找那条闭合多段线，逐点比对坐标。
    let private readBack (target: string) (expected: (float * float) list) =
        use db = new Database(false, true)
        db.ReadDwgFile(target, FileOpenMode.OpenForReadAndAllShare, true, null)
        db.CloseInput(true)
        use tr = db.TransactionManager.StartTransaction()
        let bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) :?> BlockTable
        let ms = tr.GetObject(bt.[BlockTableRecord.ModelSpace], OpenMode.ForRead) :?> BlockTableRecord
        let mutable polylines = 0
        let mutable points: (float * float) list = []

        for oid in (ms |> Seq.cast<ObjectId>) do
            match tr.GetObject(oid, OpenMode.ForRead) with
            | :? Polyline as pl when pl.Closed ->
                polylines <- polylines + 1
                points <-
                    [ for i in 0 .. pl.NumberOfVertices - 1 ->
                        let p = pl.GetPoint2dAt i
                        (p.X, p.Y) ]
            | _ -> ()

        tr.Commit()

        let delta =
            if List.length points <> List.length expected then
                Double.PositiveInfinity
            else
                List.zip expected points
                |> List.map (fun ((x1, y1), (x2, y2)) -> Math.Abs(x1 - x2) + Math.Abs(y1 - y2))
                |> List.max

        (polylines, List.length points, delta)

    /// XH_STAR 的本体：画 → 另存 → 读回 → 写报告。任何一步出错都写进报告，不静默。
    let run () =
        let target = outputPath ()
        say "star" ("目标 " + target)

        let head =
            [ "out", target
              "center", sprintf "%.1f,%.1f" centerX centerY
              "radius", sprintf "%.1f" radius ]

        if File.Exists target then
            let reason = "目标已存在，不覆盖：" + target
            say "star" ("error " + reason)
            writeReport (head @ [ "status", "ERROR"; "error", reason ])
        else
            try
                let db = HostApplicationServices.WorkingDatabase
                let erased, points = drawInto db
                say "star" (sprintf "清空种子实体 %d 个，画闭合多段线 %d 顶点" erased (List.length points))

                db.SaveAs(target, DwgVersion.Current)

                let size = FileInfo(target).Length
                let header = headerOf target
                let polylines, count, delta = readBack target points
                let status = if polylines = 1 && count = 10 && delta < 1e-9 then "OK" else "MISMATCH"

                say "star" (sprintf "保存 %d 字节，DWG 头 %s" size header)
                say "star" (sprintf "读回 闭合多段线 %d 条 / 顶点 %d 个 / 最大偏差 %g" polylines count delta)
                say "star" ("status " + status)

                let rows =
                    [ "erased", string erased
                      "vertices", string (List.length points)
                      "closed", "True"
                      "perimeter", sprintf "%.6f" (StarGeometry.perimeter points)
                      "saved", string size
                      "header", header
                      "readback_polylines", string polylines
                      "readback_vertices", string count
                      "readback_max_delta", sprintf "%g" delta
                      "status", status ]

                writeReport (head @ rows)
            with ex ->
                let reason = ex.GetType().Name + ": " + ex.Message
                say "star" ("error " + reason)
                writeReport (head @ [ "status", "ERROR"; "error", reason ])
