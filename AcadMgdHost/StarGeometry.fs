namespace AcadMgdHost

open System

/// 五角星的纯几何：只算 (x, y)，不碰任何 AutoCAD 类型。
/// 这份源文件同时被 AcadMgdHost 插件与 xUnitHost 的 Fact 编译（源文件链接），
/// 所以「画出来的形状」和「断言的形状」永远是同一个公式。
module internal StarGeometry =

    /// 内凹顶点 / 外凸顶点的半径比：sin18° / sin54° = 1/φ² ≈ 0.381966。
    /// 用这个比值，五个内凹点正好落在相邻外点的连线上，画出来才是标准的 ★。
    let innerRatio = Math.Sin(Math.PI / 10.0) / Math.Sin(3.0 * Math.PI / 10.0)

    /// 中心 (cx, cy)、外接圆半径 radius；从正上方起每 36° 一点，外 / 内交替，共 10 点。
    let vertices (cx: float) (cy: float) (radius: float) : (float * float) list =
        [ for i in 0 .. 9 ->
            let angle = Math.PI / 2.0 + float i * Math.PI / 5.0
            let r = if i % 2 = 0 then radius else radius * innerRatio
            (cx + r * Math.Cos angle, cy + r * Math.Sin angle) ]

    /// 首尾相接的十边形周长：报告里贴一个真实值，不写「成功」两个字。
    let perimeter (points: (float * float) list) : float =
        match points with
        | [] -> 0.0
        | first :: _ ->
            (points @ [ first ])
            |> List.pairwise
            |> List.sumBy (fun ((x1, y1), (x2, y2)) ->
                Math.Sqrt((x2 - x1) ** 2.0 + (y2 - y1) ** 2.0))
