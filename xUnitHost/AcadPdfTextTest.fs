/// PDF 导入的图纸里，文字常常被打散成 LWPOLYLINE 轮廓。要把它们「原位变回文字」，
/// 得先知道每条多段线究竟是哪个字 —— 轮廓本身不带字符信息。
///
/// 所以这条路分两步，acad 只被访问两次：
///   1. `*顶点*`：只用 `handent` + `entget` 把选中 9 条的**全部顶点坐标**写成文本文件（只读，不动图纸）；
///      本地按轮廓渲成图片认字（Python + matplotlib，不碰 acad）。
///   2. `*原位转化*`：一张 .lsp 里建 TEXT、按原包围盒拟合高宽、删掉原多段线，`UNDO` 编组、**不保存**。
///
/// 三条真机踩出来的规矩（都是实测，不是猜的）：
///   · `entmod` 必须拿 `(entget 实体名)` 的列表：`entmake` 返回的列表**没有 -1 组**，
///     用它 entmod 不报错、也不生效（宽度因子白设，墨迹还是原宽）。
///   · 删实体别放进 `(command "_.UNDO" "_BE") … "_E"` 编组里：那样 `entdel` 一条也删不掉
///     （POLY 计数纹丝不动）；移到编组外用 `vla-delete` 才真删。
///   · 判「还在不在图里」不能用 `handent`（它对被删实体照样返回句柄），
///     也不能靠 `vlax-ename->vla-object` 报错（它对被删实体照样能转）—— 看这一层 POLY 的计数。
///   · 脚本一律纯 ASCII：AutoCAD 2020 的 AutoLISP 按本地 ANSI(GBK) 读 .lsp；
///     图层名从原实体的 `(assoc 8)` 里取，不写进脚本。
///
/// 跑法：
///   XUNITHOST_ACAD_PID=8868 xUnitHost.exe -method "*顶点*" -noColor -showLiveOutput
///   XUNITHOST_ACAD_PID=8868 XUNITHOST_ACAD_TEXT=1 xUnitHost.exe -method "*原位转化*" -noColor -showLiveOutput
namespace XUnitHost

open System
open System.IO
open System.Text
open System.Threading
open System.Threading.Tasks
open Xunit
open FSharp.xUnit
open AutoCADHost
open OsHost

/// 黑板上的那次选择（2026-11-12 09:46 采集，PID 8868，图纸 高安屯PID.dwg）：
/// 9 条 LWPOLYLINE，全在 `PDF2_几何图形` 图层上。句柄写死在这里 —— 用户换了选择必须同步改。
module internal PdfTextPick =

    let pid = 8868

    let handles =
        [ "125E"; "125D"; "125C"; "125B"; "1259"; "1258"; "1257"; "1256"; "1255" ]

    /// 认字结果：把 9 条的顶点按真实坐标渲成图，再和候选字轮廓逐字套包围盒叠加核对——
    /// 1255=2、1256=5、1257=0、1258+1259=E、125B+125C=X、125D=S、125E=6（用户已确认）。
    let word = "250EXS6"

    /// 转储与回执的落点（项目里，方便事后复核）
    let dumpFile =
        @"C:\Application Data\GitHub\xp44mm\xUnitHost\TestResults\acad_pdftext_dump.txt"

/// 顶点转储与回写：脚本正文、「写稳」判据、以及对自己转储的解析。
module internal PdfTextDump =

    /// 可以从结果文件独占打开、且内容非空 ⇒ LISP 已经 close，结果可信
    let settled (path: string) : string option =
        if not (File.Exists path) then
            None
        else
            try
                use stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None)
                use reader = new StreamReader(stream, Encoding.UTF8)
                let text = reader.ReadToEnd().Trim()
                if text = "" then None else Some text
            with _ ->
                None

    let awaitSettled (path: string) (timeoutMs: int) : Task<string> =
        Task.Run(fun () ->
            let clock = Diagnostics.Stopwatch.StartNew()
            let mutable found: string option = None

            while found.IsNone && clock.ElapsedMilliseconds < int64 timeoutMs do
                found <- settled path
                if found.IsNone then Thread.Sleep 100

            match found with
            | Some text -> text
            | None -> failwithf "LISP 结果文件 %dms 内没写完：%s" timeoutMs path)

    /// 纯 ASCII 的 .lsp 正文：每条实体一行，格式
    ///   `句柄|T=类型|L=图层|F=标志|P=x,y|B=凸度|P=…|`
    /// 顶点与凸度按 entget 的先后顺序落，解析时 P 起新点、B 补上一个点。
    let script (handles: string list) (resultPath: string) : string =
        let quoted = handles |> List.map(sprintf "\"%s\"") |> String.concat " "

        let body =
            [ """;; Pure ASCII on purpose: AutoCAD 2020 evaluates .lsp in the local ANSI code page."""
              """(defun xh-rt (x) (rtos x 2 8))"""
              """(defun xh-line (h / e d out)"""
              """  (setq e (handent h))"""
              """  (if e"""
              """    (progn"""
              """      (setq d (entget e))"""
              """      (setq out (strcat h "|"))"""
              """      (foreach pr d"""
              """        (cond"""
              """          ((= (car pr) 0) (setq out (strcat out "T=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 8) (setq out (strcat out "L=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 70) (setq out (strcat out "F=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 10) (setq out (strcat out "P=" (xh-rt (cadr pr)) "," (xh-rt (caddr pr)) "|")))"""
              """          ((= (car pr) 42) (setq out (strcat out "B=" (xh-rt (cdr pr)) "|")))"""
              """        )"""
              """      )"""
              """      out"""
              """    )"""
              """    (strcat h "|MISSING|")"""
              """  )"""
              """)"""
              sprintf """(setq xh-hs '(%s))""" quoted
              sprintf """(setq xh-f (open "%s" "w"))""" (resultPath.Replace('\\', '/'))
              """(foreach xh-h xh-hs (write-line (xh-line xh-h) xh-f))"""
              """(write-line (strcat "DWG=" (getvar "DWGNAME")) xh-f)"""
              """(write-line (strcat "ACADVER=" (getvar "ACADVER")) xh-f)"""
              """(close xh-f)"""
              """(princ)""" ]

        String.Join("\r\n", body) + "\r\n"

    /// 一条实体的转储：顶点、凸度、包围盒
    type Item =
        { Handle: string
          Kind: string
          Layer: string
          Flags: int
          Points: (float * float * float) list }   // x, y, bulge

        member this.Width =
            match this.Points with
            | [] -> 0.0
            | points -> (points |> List.map(fun (x, _, _) -> x) |> List.max) - (points |> List.map(fun (x, _, _) -> x) |> List.min)

        member this.Height =
            match this.Points with
            | [] -> 0.0
            | points -> (points |> List.map(fun (_, y, _) -> y) |> List.max) - (points |> List.map(fun (_, y, _) -> y) |> List.min)

    let private num (text: string) =
        match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, value -> value
        | _ -> 0.0

    /// `句柄|T=…|L=…|F=…|P=x,y|B=b|…` → Item；`DWG=` / `ACADVER=` 这类尾行交回 None
    let parse (line: string) : Item option =
        let tokens =
            line.Split('|', StringSplitOptions.RemoveEmptyEntries) |> Array.toList

        match tokens with
        | handle :: (_ :: _ as rest) when line.Contains "|T=" ->
            let mutable kind = ""
            let mutable layer = ""
            let mutable flags = 0
            let mutable points = []

            for token in rest do
                if token.StartsWith "T=" then kind <- token.Substring 2
                elif token.StartsWith "L=" then layer <- token.Substring 2
                elif token.StartsWith "F=" then flags <- int (num (token.Substring 2))
                elif token.StartsWith "P=" then
                    let xy = token.Substring(2).Split(',')

                    if xy.Length = 2 then
                        points <- (num xy.[0], num xy.[1], 0.0) :: points
                elif token.StartsWith "B=" then
                    match points with
                    | (x, y, _) :: tail -> points <- (x, y, num (token.Substring 2)) :: tail
                    | [] -> ()

            Some
                { Handle = handle
                  Kind = kind
                  Layer = layer
                  Flags = flags
                  Points = List.rev points }
        | _ -> None

    /// 读转储（AutoCAD 的 `write-line` 按本地 ANSI 落盘 → 一律走 TextFile 认编码）
    let load (path: string) : Task<Item list> =
        task {
            let! _encoding, rows = TextFile.decodeAllLinesAsync path
            return rows |> Array.toList |> List.choose parse
        }

    /// 全部顶点合起来的包围盒：x 最小 / y 最小 / 宽 / 高
    let extent (items: Item list) =
        let points = items |> List.collect(fun item -> item.Points)

        match points with
        | [] -> failwith "转储里没有顶点"
        | _ ->
            let xs = points |> List.map(fun (x, _, _) -> x)
            let ys = points |> List.map(fun (_, y, _) -> y)
            List.min xs, List.min ys, List.max xs - List.min xs, List.max ys - List.min ys

    /// 写文字并删原多段线的那张 .lsp。顺序：找图层（原线还在就取它，被删过就找同字同位的旧 TEXT）
    /// → 解锁图层 → 删旧 TEXT → 建新 TEXT 并按原包围盒拟合 → 删原多段线 → 复锁。
    ///
    /// 拟合是两步实测：先按标称高度建出来，用 `(textbox)` 量墨迹盒，
    /// 再按 `目标高 / 墨迹高` 改 40 组（高度）、按 `目标宽 / 缩放后墨迹宽` 改 41 组（宽度因子），
    /// 并把墨迹盒左下角的偏移从插入点里减掉 —— 这样新字的墨迹和原轮廓严丝合缝。
    ///
    /// `entmod` 必须拿 `(entget 实体名)` 的列表：**缺 -1 组的列表 entmod 不认**，宽度因子会悄悄不生效。
    /// `UNDO` 编组，用户一次 U 就整段撤销；**全程不保存图纸**。
    let writeScript (text: string) (handles: string list) (x0: float) (y0: float) (w: float) (h: float) (logPath: string) =
        let rt (value: float) =
            value.ToString("0.00000000", Globalization.CultureInfo.InvariantCulture)

        let quoted = handles |> List.map(sprintf "\"%s\"") |> String.concat " "

        let body =
            [ """;; Pure ASCII on purpose: AutoCAD 2020 evaluates .lsp in the local ANSI code page."""
              """(vl-load-com)"""
              """(defun xh-rt (x) (rtos x 2 8))"""
              sprintf """(setq xh-txt "%s")""" text
              sprintf """(setq xh-hs '(%s))""" quoted
              sprintf """(setq xh-x0 %s xh-y0 %s xh-hh %s xh-ww %s)""" (rt x0) (rt y0) (rt h) (rt w)
              sprintf """(setq xh-f (open "%s" "w"))""" (logPath.Replace('\\', '/'))
              """;; 1) 图层：原多段线还在就取它的；已经被删过就找同字同位的旧 TEXT"""
              """(setq xh-lay nil)"""
              """(foreach xh-h xh-hs (if (and (null xh-lay) (handent xh-h)) (setq xh-lay (cdr (assoc 8 (entget (handent xh-h)))))))"""
              """(setq xh-old nil)"""
              """(setq xh-ss (ssget "_X" (list (cons 0 "TEXT") (cons 1 xh-txt))))"""
              """(if xh-ss"""
              """  (progn"""
              """    (setq xh-i 0)"""
              """    (while (< xh-i (sslength xh-ss))"""
              """      (setq xh-dd (entget (ssname xh-ss xh-i)))"""
              """      (if (and (< (abs (- (car (cdr (assoc 10 xh-dd))) xh-x0)) 0.05)"""
              """               (< (abs (- (cadr (cdr (assoc 10 xh-dd))) xh-y0)) 0.05))"""
              """        (progn"""
              """          (if (null xh-old) (setq xh-old (ssname xh-ss xh-i)))"""
              """          (if (null xh-lay) (setq xh-lay (cdr (assoc 8 xh-dd))))"""
              """        )"""
              """      )"""
              """      (setq xh-i (1+ xh-i))"""
              """    )"""
              """  )"""
              """)"""
              """(if (null xh-lay)"""
              """  (write-line "FAIL no-layer" xh-f)"""
              """  (progn"""
              """    (command "_.UNDO" "_BE")"""
              """    (setq xh-lo (tblobjname "LAYER" xh-lay))"""
              """    (setq xh-ld (entget xh-lo))"""
              """    (setq xh-flags (cdr (assoc 70 xh-ld)))"""
              """    (setq xh-locked (= 4 (logand 4 xh-flags)))"""
              """    (write-line (strcat "LAYER=" xh-lay " flags=" (xh-rt xh-flags) " locked=" (if xh-locked "1" "0")) xh-f)"""
              """    (if xh-locked (entmod (subst (cons 70 (- xh-flags 4)) (assoc 70 xh-ld) xh-ld)))"""
              """    (write-line (strcat "STALE=" (if xh-old (cdr (assoc 5 (entget xh-old))) "none")) xh-f)"""
              """    (if xh-old (entdel xh-old))"""
              """    (setq xh-el (entmake (list '(0 . "TEXT")"""
              """                              (cons 8 xh-lay)"""
              """                              (cons 10 (list xh-x0 xh-y0 0.0))"""
              """                              (cons 40 xh-hh)"""
              """                              (cons 41 1.0)"""
              """                              (cons 1 xh-txt)"""
              """                              (cons 7 (getvar "TEXTSTYLE")))))"""
              """    (if xh-el"""
              """      (progn"""
              """        (setq xh-e2 (entlast))"""
              """        (setq xh-tb (textbox (entget xh-e2)))"""
              """        (if xh-tb"""
              """          (progn"""
              """            (setq xh-bw (- (caadr xh-tb) (caar xh-tb)))"""
              """            (setq xh-bh (- (cadr (cadr xh-tb)) (cadr (car xh-tb))))"""
              """            (setq xh-ix (car (car xh-tb)))"""
              """            (setq xh-iy (cadr (car xh-tb)))"""
              """            (write-line (strcat "BOX0 w=" (xh-rt xh-bw) " h=" (xh-rt xh-bh) " ix=" (xh-rt xh-ix) " iy=" (xh-rt xh-iy)) xh-f)"""
              """            (setq xh-s (if (> xh-bh 0.0) (/ xh-hh xh-bh) 1.0))"""
              """            (setq xh-nh (* xh-hh xh-s))"""
              """            (setq xh-nx (- xh-x0 (* xh-ix xh-s)))"""
              """            (setq xh-ny (- xh-y0 (* xh-iy xh-s)))"""
              """            (setq xh-nwf (if (> (* xh-bw xh-s) 0.0) (/ xh-ww (* xh-bw xh-s)) 1.0))"""
              """            (setq xh-dd (entget xh-e2))"""
              """            (setq xh-dd (subst (cons 10 (list xh-nx xh-ny 0.0)) (assoc 10 xh-dd) xh-dd))"""
              """            (setq xh-dd (subst (cons 40 xh-nh) (assoc 40 xh-dd) xh-dd))"""
              """            (setq xh-dd (subst (cons 41 xh-nwf) (assoc 41 xh-dd) xh-dd))"""
              """            (entmod xh-dd)"""
              """            (write-line (strcat "FIT height=" (xh-rt xh-nh) " wf=" (xh-rt xh-nwf) " ins=" (xh-rt xh-nx) "," (xh-rt xh-ny)) xh-f)"""
              """          )"""
              """          (write-line "FAIL textbox" xh-f)"""
              """        )"""
              """        (setq xh-dd (entget xh-e2))"""
              """        (setq xh-tb2 (textbox xh-dd))"""
              """        (write-line (strcat "BOX1 w=" (xh-rt (- (caadr xh-tb2) (caar xh-tb2))) " h=" (xh-rt (- (cadr (cadr xh-tb2)) (cadr (car xh-tb2)))) " ix=" (xh-rt (car (car xh-tb2))) " iy=" (xh-rt (cadr (car xh-tb2)))) xh-f)"""
              """        (write-line (strcat "NEW handle=" (cdr (assoc 5 xh-dd)) " layer=" (cdr (assoc 8 xh-dd)) " txt=" (cdr (assoc 1 xh-dd)) " h=" (xh-rt (cdr (assoc 40 xh-dd))) " wf=" (xh-rt (cdr (assoc 41 xh-dd)))) xh-f)"""
              """      )"""
              """      (write-line "FAIL entmake" xh-f)"""
              """    )"""
              """    (if xh-locked (entmod (subst (cons 70 xh-flags) (assoc 70 (entget xh-lo)) (entget xh-lo))))"""
              """    (write-line (strcat "RELOCK=" (if xh-locked "1" "0")) xh-f)"""
              """    (command "_.UNDO" "_E")"""
              """;; 删原多段线：必须在 UNDO 编组之外做 —— 编组里的 entdel 实测一条也删不掉"""
              """    (setq xh-gone 0)"""
              """    (foreach xh-h xh-hs"""
              """      (setq xh-x (vl-catch-all-apply (quote vlax-ename->vla-object) (list (handent xh-h))))"""
              """      (if (not (vl-catch-all-error-p xh-x))"""
              """        (if (not (vl-catch-all-error-p (vl-catch-all-apply (quote vla-delete) (list xh-x))))"""
              """          (setq xh-gone (1+ xh-gone))"""
              """        )"""
              """      )"""
              """    )"""
              """    (write-line (strcat "ERASED=" (xh-rt xh-gone)) xh-f)"""
              """  )"""
              """)"""
              """(close xh-f)"""
              """(princ)""" ]

        String.Join("\r\n", body) + "\r\n"

/// PDF 几何图形 → 文字的两次真机作业：先转储顶点，再原地写 TEXT。
/// 黑板上的第二次选择（2026-11-12 10:14 采集，PID 8868，图纸 高安屯PID.dwg）：
/// 5 条 LWPOLYLINE，同样在 `PDF2_几何图形` 图层上。这次**不认字**，直接交给 AutoCAD 自带的
/// `PDFSHXTEXT`（Recognize SHX Text）：它拿选区几何去比 SHX 字库，命中就把几何换成 MTEXT。
/// 所以它只吃「单线笔画」的字；闭合轮廓（TrueType 被转曲的那种）它比不出来。
module internal ShxPick =

    let pid = 8868

    let handles = [ "1215"; "1214"; "1213"; "1212"; "1211" ]

    /// 上一轮写下的那条 250EXS6 TEXT 的句柄：借它的图层名（中文写进 .lsp 会被按 GBK 读坏）
    let layerFrom = "3D1F"

    /// 这 5 个字的识别结果（本地渲图：S S 3 0 4 —— 不锈钢牌号 SS304）
    let word = "SS304"

    /// 顶点转储（PDFSHXTEXT 一旦认不出来，这些数据就是退路 —— 省一次 acad 访问）
    let dumpFile =
        @"C:\Application Data\GitHub\xp44mm\xUnitHost\TestResults\acad_shx_dump.txt"

    /// 本轮的判据落点：图层标志、前后计数、识别出的文字
    let logFile =
        @"C:\Application Data\GitHub\xp44mm\xUnitHost\TestResults\acad_shx_log.txt"

/// 一张 .lsp 干两件事：
///   1. 先把 5 条的顶点转储出去（**先写先关**：后面真卡住了，这段也一定落在盘上）；
///   2. `UCS _W` + 关捕捉 + `CMDECHO 0`，把选区喂给 `PDFSHXTEXT`，前后比对这一层的
///      POLY / TEXT 计数，再把这一层上所有 TEXT/MTEXT 的内容与位置原样报回来。
///
/// 为什么不 `(load)` 等 EndLisp：`PDFSHXTEXT` 结束时会报「N of M groups converted」，
/// 万一走的是模态提示框，求值就停在那儿，等 EndLisp 会一直挂着 —— 所以发出去就不等，自己盯文件。
module internal ShxRun =

    let private rt (value: float) =
        value.ToString("0.00000000", Globalization.CultureInfo.InvariantCulture)

    let script (handles: string list) (dumpPath: string) (logPath: string) : string =
        let quoted = handles |> List.map(sprintf "\"%s\"") |> String.concat " "

        let body =
            [ """;; Pure ASCII on purpose: AutoCAD 2020 evaluates .lsp in the local ANSI code page."""
              """(vl-load-com)"""
              """(defun xh-rt (x) (rtos x 2 8))"""
              """(defun xh-g (d k def / v) (setq v (cdr (assoc k d))) (if v v def))"""
              """(defun xh-line (h / e d out)"""
              """  (setq e (handent h))"""
              """  (if e"""
              """    (progn"""
              """      (setq d (entget e))"""
              """      (setq out (strcat h "|"))"""
              """      (foreach pr d"""
              """        (cond"""
              """          ((= (car pr) 0) (setq out (strcat out "T=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 8) (setq out (strcat out "L=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 70) (setq out (strcat out "F=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 10) (setq out (strcat out "P=" (xh-rt (cadr pr)) "," (xh-rt (caddr pr)) "|")))"""
              """          ((= (car pr) 42) (setq out (strcat out "B=" (xh-rt (cdr pr)) "|")))"""
              """        )"""
              """      )"""
              """      out"""
              """    )"""
              """    (strcat h "|MISSING|")"""
              """  )"""
              """)"""
              sprintf """(setq xh-hs '(%s))""" quoted
              """(setq xh-lay nil)"""
              """(foreach xh-h xh-hs (if (and (null xh-lay) (handent xh-h)) (setq xh-lay (cdr (assoc 8 (entget (handent xh-h)))))))"""
              """;; 句柄全失效时给个存在的图层，别让 tblobjname 报错把脚本整段打断"""
              """(if (null xh-lay) (setq xh-lay "0"))"""
              """;; ---- 1) 顶点转储：写一段就关，后面卡住也不影响它 ----"""
              sprintf """(setq xh-f (open "%s" "w"))""" (dumpPath.Replace('\\', '/'))
              """(write-line (strcat "LAYER=" xh-lay) xh-f)"""
              """(foreach xh-h xh-hs (write-line (xh-line xh-h) xh-f))"""
              """(close xh-f)"""
              """;; ---- 2) PDFSHXTEXT ----"""
              sprintf """(setq xh-f (open "%s" "w"))""" (logPath.Replace('\\', '/'))
              """(setq xh-oc (getvar "CMDECHO") xh-oo (getvar "OSMODE"))"""
              """(setvar "CMDECHO" 0)"""
              """(setvar "OSMODE" 0)"""
              """(setvar "EXPERT" 5)"""
              """(setq xh-ld (entget (tblobjname "LAYER" xh-lay)))"""
              """(write-line (strcat "LAYER=" xh-lay " flags=" (xh-rt (xh-g xh-ld 70 0))) xh-f)"""
              """(setq xh-p0 (sslength (ssget "_X" (list (cons 0 "LWPOLYLINE") (cons 8 xh-lay)))))"""
              """(setq xh-t0 (sslength (ssget "_X" (list (cons 0 "TEXT,MTEXT") (cons 8 xh-lay)))))"""
              """(write-line (strcat "CNT before poly=" (itoa xh-p0) " text=" (itoa xh-t0)) xh-f)"""
              """(write-line (strcat "CNT all before=" (itoa (if (setq xh-q (ssget "_X" '((0 . "TEXT,MTEXT")))) (sslength xh-q) 0))) xh-f)"""
              """;; 文字方向要跟 UCS 的 X 轴一致，否则它一个都比不出来；这张图的字是正的，回世界坐标系"""
              """(command "_.UCS" "_W")"""
              """(setq xh-ss (ssadd))"""
              """(foreach xh-h xh-hs (if (handent xh-h) (ssadd (handent xh-h) xh-ss)))"""
              """(write-line (strcat "SEL=" (itoa (sslength xh-ss))) xh-f)"""
              """(setq xh-r (vl-catch-all-apply (quote (lambda () (command "_.PDFSHXTEXT" xh-ss "")))))"""
              """(write-line (strcat "CMD=" (if (vl-catch-all-error-p xh-r) (strcat "ERR " (vl-catch-all-error-message xh-r)) "ok")) xh-f)"""
              """(setq xh-p1 (sslength (ssget "_X" (list (cons 0 "LWPOLYLINE") (cons 8 xh-lay)))))"""
              """(setq xh-t1 (sslength (ssget "_X" (list (cons 0 "TEXT,MTEXT") (cons 8 xh-lay)))))"""
              """(write-line (strcat "CNT after poly=" (itoa xh-p1) " text=" (itoa xh-t1)) xh-f)"""
              """(write-line (strcat "CNT all after=" (itoa (if (setq xh-q (ssget "_X" '((0 . "TEXT,MTEXT")))) (sslength xh-q) 0))) xh-f)"""
              """;; ---- 3) 这一层上所有文字，内容与位置原样报回来 ----"""
              """(setq xh-xs (ssget "_X" (list (cons 0 "TEXT,MTEXT") (cons 8 xh-lay))))"""
              """(if xh-xs"""
              """  (progn"""
              """    (setq xh-i 0)"""
              """    (while (< xh-i (sslength xh-xs))"""
              """      (setq xh-d (entget (ssname xh-xs xh-i)))"""
              """      (setq xh-p (xh-g xh-d 10 (list 0.0 0.0 0.0)))"""
              """      (write-line (strcat "TXT=" (xh-g xh-d 5 "?") "|" (xh-g xh-d 0 "?")"""
              """                          "|h=" (xh-rt (xh-g xh-d 40 0.0))"""
              """                          "|rot=" (xh-rt (xh-g xh-d 50 0.0))"""
              """                          "|ins=" (xh-rt (car xh-p)) "," (xh-rt (cadr xh-p))"""
              """                          "|s=" (xh-g xh-d 1 "")) xh-f)"""
              """      (setq xh-i (1+ xh-i))"""
              """    )"""
              """  )"""
              """)"""
              """(command "_.UCS" "_P")"""
              """(setvar "CMDECHO" xh-oc)"""
              """(setvar "OSMODE" xh-oo)"""
              """(write-line "DONE" xh-f)"""
              """(close xh-f)"""
              """(princ)""" ]

        String.Join("\r\n", body) + "\r\n"

    /// 兜底：`PDFSHXTEXT` 把几何吃掉了却没在原图层留下文字时用这张。
    /// 先扫这个区域（**不限图层**）看它是不是把 MTEXT 建到了当前层 —— 是就搬回原图层；
    /// 全图都没有才按原包围盒补一条 TEXT 并两步拟合。
    ///
    /// 图层名是中文，写进 .lsp 会被按 GBK 读坏 —— 借同层那个 250EXS6（TEXT）的 `(assoc 8)`。
    let rescueScript
        (text: string)
        (x0: float)
        (y0: float)
        (w: float)
        (h: float)
        (layerFrom: string)
        (logPath: string)
        : string =
        let body =
            [ """;; Pure ASCII on purpose: AutoCAD 2020 evaluates .lsp in the local ANSI code page."""
              """(vl-load-com)"""
              """(defun xh-rt (x) (rtos x 2 8))"""
              sprintf """(setq xh-txt "%s")""" text
              sprintf """(setq xh-x0 %s xh-y0 %s xh-hh %s xh-ww %s)""" (rt x0) (rt y0) (rt h) (rt w)
              sprintf """(setq xh-f (open "%s" "w"))""" (logPath.Replace('\\', '/'))
              """;; 图层名从这里借：同层那条文字还在"""
              sprintf """(setq xh-lay (if (handent "%s") (cdr (assoc 8 (entget (handent "%s")))) nil))""" layerFrom layerFrom
              """(if (null xh-lay)"""
              """  (progn"""
              """    (setq xh-s (ssget "_X" '((0 . "TEXT"))))"""
              """    (if xh-s (setq xh-lay (cdr (assoc 8 (entget (ssname xh-s 0))))))"""
              """  )"""
              """)"""
              """(if (null xh-lay)"""
              """  (write-line "FAIL no-layer" xh-f)"""
              """  (progn"""
              """    (write-line "LAYER=hit" xh-f)"""
              """;; ---- 1) 区域里（各图层）已经存在的文字 ----"""
              """    (setq xh-hit nil)"""
              """    (setq xh-s (ssget "_X" '((0 . "TEXT,MTEXT"))))"""
              """    (if xh-s"""
              """      (progn"""
              """        (setq xh-i 0)"""
              """        (while (< xh-i (sslength xh-s))"""
              """          (setq xh-d (entget (ssname xh-s xh-i)))"""
              """          (setq xh-p (cdr (assoc 10 xh-d)))"""
              """          (if (and (>= (car xh-p) 9.28) (<= (car xh-p) 9.78) (>= (cadr xh-p) 17.16) (<= (cadr xh-p) 17.35))"""
              """            (progn"""
              """              (setq xh-hit (ssname xh-s xh-i))"""
              """              (write-line (strcat "FOUND h=" (cdr (assoc 5 xh-d)) " onlayer=" (if (equal (cdr (assoc 8 xh-d)) xh-lay) "1" "0") " s=" (cdr (assoc 1 xh-d))) xh-f)"""
              """              (if (not (equal (cdr (assoc 8 xh-d)) xh-lay))"""
              """                (progn"""
              """                  (entmod (subst (cons 8 xh-lay) (assoc 8 xh-d) xh-d))"""
              """                  (write-line "MOVED=1" xh-f)"""
              """                )"""
              """              )"""
              """            )"""
              """          )"""
              """          (setq xh-i (1+ xh-i))"""
              """        )"""
              """      )"""
              """    )"""
              """;; ---- 2) 一条都没有：按原包围盒补写，两步拟合 ----"""
              """    (if (null xh-hit)"""
              """      (progn"""
              """        (write-line "NOHIT=1" xh-f)"""
              """        (setq xh-el (entmake (list '(0 . "TEXT")"""
              """                                   (cons 8 xh-lay)"""
              """                                   (cons 10 (list xh-x0 xh-y0 0.0))"""
              """                                   (cons 40 xh-hh)"""
              """                                   (cons 41 1.0)"""
              """                                   (cons 1 xh-txt)"""
              """                                   (cons 7 (getvar "TEXTSTYLE")))))"""
              """        (if xh-el"""
              """          (progn"""
              """            (setq xh-e2 (entlast))"""
              """            (setq xh-tb (textbox (entget xh-e2)))"""
              """            (if xh-tb"""
              """              (progn"""
              """                (setq xh-bw (- (caadr xh-tb) (caar xh-tb)))"""
              """                (setq xh-bh (- (cadr (cadr xh-tb)) (cadr (car xh-tb))))"""
              """                (setq xh-ix (car (car xh-tb)))"""
              """                (setq xh-iy (cadr (car xh-tb)))"""
              """                (write-line (strcat "BOX0 w=" (xh-rt xh-bw) " h=" (xh-rt xh-bh)) xh-f)"""
              """                (setq xh-sc (if (> xh-bh 0.0) (/ xh-hh xh-bh) 1.0))"""
              """                (setq xh-nh (* xh-hh xh-sc))"""
              """                (setq xh-nx (- xh-x0 (* xh-ix xh-sc)))"""
              """                (setq xh-ny (- xh-y0 (* xh-iy xh-sc)))"""
              """                (setq xh-nwf (if (> (* xh-bw xh-sc) 0.0) (/ xh-ww (* xh-bw xh-sc)) 1.0))"""
              """                (setq xh-dd (entget xh-e2))"""
              """                (setq xh-dd (subst (cons 10 (list xh-nx xh-ny 0.0)) (assoc 10 xh-dd) xh-dd))"""
              """                (setq xh-dd (subst (cons 40 xh-nh) (assoc 40 xh-dd) xh-dd))"""
              """                (setq xh-dd (subst (cons 41 xh-nwf) (assoc 41 xh-dd) xh-dd))"""
              """                (entmod xh-dd)"""
              """                (setq xh-dd (entget xh-e2))"""
              """                (setq xh-tb2 (textbox xh-dd))"""
              """                (write-line (strcat "BOX1 w=" (xh-rt (- (caadr xh-tb2) (caar xh-tb2))) " h=" (xh-rt (- (cadr (cadr xh-tb2)) (cadr (car xh-tb2))))) xh-f)"""
              """                (write-line (strcat "NEW h=" (cdr (assoc 5 xh-dd)) " txt=" (cdr (assoc 1 xh-dd)) " height=" (xh-rt (cdr (assoc 40 xh-dd))) " wf=" (xh-rt (cdr (assoc 41 xh-dd)))) xh-f)"""
              """              )"""
              """              (write-line "FAIL textbox" xh-f)"""
              """            )"""
              """          )"""
              """          (write-line "FAIL entmake" xh-f)"""
              """        )"""
              """      )"""
              """    )"""
              """    (setq xh-s (ssget "_X" (list (cons 0 "TEXT,MTEXT") (cons 8 xh-lay))))"""
              """    (write-line (strcat "CNT text=" (itoa (if xh-s (sslength xh-s) 0))) xh-f)"""
              """    (setq xh-s (ssget "_X" (list (cons 0 "LWPOLYLINE") (cons 8 xh-lay))))"""
              """    (write-line (strcat "CNT poly=" (itoa (if xh-s (sslength xh-s) 0))) xh-f)"""
              """    (write-line "DONE" xh-f)"""
              """  )"""
              """)"""
              """(close xh-f)"""
              """(princ)""" ]

        String.Join("\r\n", body) + "\r\n"


/// 黑板上的第三次选择（2026-11-12 10:46 采集，PID 8868，图纸 高安屯PID.dwg）：
/// 55 条，全在 `PDF_几何图形` 图层上（注意不是 PDF2），其中 0x4DC 是 ARC，其余 LWPOLYLINE。
/// 这次用户直接给了字：`去2#尿素热解炉`。中文 `PDFSHXTEXT` 认不了（不支持亚洲大字库），
/// 所以走「建 TEXT + 按 bbox 两步拟合 + 删原线」那条已经验证过的路。
module internal BigPick =

    let pid = 8868

    /// 0x508 递减到 0x4D2，共 55 条；句柄写成大写十六进制、不带 0x（`(assoc 5)` 就长这样）
    let handles =
        [ for h in 0x508 .. -1 .. 0x4D2 -> h.ToString "X" ]

    let word = "去2#尿素热解炉"

    /// 「去2#尿素热解炉」的 GBK 字节（14 字节 / 8 个字符）。
    /// 脚本必须纯 ASCII —— 中文只能靠 `(chr n)` 把这些字节拼回字符串。
    let wordBytes =
        [ 0xC8; 0xA5   // 去
          0x32; 0x23   // 2 #
          0xC4; 0xF2   // 尿
          0xCB; 0xD8   // 素
          0xC8; 0xC8   // 热
          0xBD; 0xE2   // 解
          0xC2; 0xAF ] // 炉

    /// 改之前先把 55 条的顶点转储出来（万一要重建，这是唯一的原料）
    let dumpFile =
        @"C:\Application Data\GitHub\xp44mm\xUnitHost\TestResults\acad_big_dump.txt"

    let logFile =
        @"C:\Application Data\GitHub\xp44mm\xUnitHost\TestResults\acad_big_log.txt"

/// 一张 .lsp 全干完（只访问 acad 一次）：
///   0. 转储 55 条顶点（改之前）；
///   1. `GetBoundingBox` 合并包围盒（**ARC 也认**，比逐点算可靠），图层取自第一条；
///   2. 借一个「内容含中文」的 TEXT/MTEXT 的文字样式 —— 免得写出来是问号；
///   3. 建 TEXT 并按 `(textbox)` 实测墨迹盒两步拟合（改 40 高、改 41 宽度因子、挪插入点）；
///   4. 墨迹高与目标差 ≤20% 才删原线（差太多说明样式把字高钉死了，宁可留原线给人看）；
///   5. 回执写文件：`STRLEN` / `bbox` / `BOX0` / `BOX1` / `txtlen` / `ERASED` / `CNT geom`。
module internal BigRun =

    let private rt (value: float) =
        value.ToString("0.00000000", Globalization.CultureInfo.InvariantCulture)

    let private chrSeq (bytes: int list) =
        bytes |> List.map(sprintf "(chr %d)") |> String.concat " "

    let script (handles: string list) (bytes: int list) (dumpPath: string) (logPath: string) : string =
        let quoted = handles |> List.map(sprintf "\"%s\"") |> String.concat " "
        let chars = chrSeq bytes

        let body =
            [ """;; Pure ASCII on purpose: AutoCAD 2020 evaluates .lsp in the local ANSI code page."""
              """;; 中文不进脚本：字串由下面的 (chr n) 把 GBK 字节拼回来"""
              """(vl-load-com)"""
              """(defun xh-rt (x) (rtos x 2 8))"""
              """;; GetBoundingBox 回来的可能是 VARIANT 也可能直接是 SAFEARRAY"""
              """(defun xh-pt (v)"""
              """  (if (= (type v) 'VARIANT) (setq v (vlax-variant-value v)))"""
              """  (if (= (type v) 'SAFEARRAY) (setq v (vlax-safearray->list v)))"""
              """  v"""
              """)"""
              """(defun xh-line (h / e d out)"""
              """  (setq e (handent h))"""
              """  (if e"""
              """    (progn"""
              """      (setq d (entget e))"""
              """      (setq out (strcat h "|"))"""
              """      (foreach pr d"""
              """        (cond"""
              """          ((= (car pr) 0) (setq out (strcat out "T=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 8) (setq out (strcat out "L=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 70) (setq out (strcat out "F=" (vl-princ-to-string (cdr pr)) "|")))"""
              """          ((= (car pr) 10) (setq out (strcat out "P=" (xh-rt (cadr pr)) "," (xh-rt (caddr pr)) "|")))"""
              """          ((= (car pr) 42) (setq out (strcat out "B=" (xh-rt (cdr pr)) "|")))"""
              """        )"""
              """      )"""
              """      out"""
              """    )"""
              """    (strcat h "|MISSING|")"""
              """  )"""
              """)"""
              sprintf """(setq xh-hs '(%s))""" quoted
              sprintf """(setq xh-txt (strcat %s))""" chars
              """;; ---- 0) 改之前先转储 ----"""
              sprintf """(setq xh-f (open "%s" "w"))""" (dumpPath.Replace('\\', '/'))
              """(foreach xh-h xh-hs (write-line (xh-line xh-h) xh-f))"""
              """(close xh-f)"""
              """;; ---- 1) 合并包围盒 + 取图层 ----"""
              sprintf """(setq xh-f (open "%s" "w"))""" (logPath.Replace('\\', '/'))
              """(write-line (strcat "STRLEN=" (itoa (strlen xh-txt))) xh-f)"""
              """(setq xh-minx nil xh-miny nil xh-maxx nil xh-maxy nil xh-lay nil xh-n 0)"""
              """(foreach xh-h xh-hs"""
              """  (setq xh-e (handent xh-h))"""
              """  (if xh-e"""
              """    (progn"""
              """      (setq xh-n (1+ xh-n))"""
              """      (if (null xh-lay) (setq xh-lay (cdr (assoc 8 (entget xh-e)))))"""
              """      (setq xh-mn (vlax-3d-point 0.0 0.0 0.0))"""
              """      (setq xh-mx (vlax-3d-point 0.0 0.0 0.0))"""
              """      (vla-getboundingbox (vlax-ename->vla-object xh-e) 'xh-mn 'xh-mx)"""
              """      (setq xh-a (xh-pt xh-mn))"""
              """      (setq xh-b (xh-pt xh-mx))"""
              """      (if (and (listp xh-a) (listp xh-b))"""
              """        (progn"""
              """          (if (or (null xh-minx) (< (car xh-a) xh-minx)) (setq xh-minx (car xh-a)))"""
              """          (if (or (null xh-miny) (< (cadr xh-a) xh-miny)) (setq xh-miny (cadr xh-a)))"""
              """          (if (or (null xh-maxx) (> (car xh-b) xh-maxx)) (setq xh-maxx (car xh-b)))"""
              """          (if (or (null xh-maxy) (> (cadr xh-b) xh-maxy)) (setq xh-maxy (cadr xh-b)))"""
              """        )"""
              """      )"""
              """    )"""
              """  )"""
              """)"""
              """(if (or (null xh-minx) (null xh-lay))"""
              """  (write-line "FAIL no-bbox" xh-f)"""
              """  (progn"""
              """    (setq xh-x0 xh-minx)"""
              """    (setq xh-y0 xh-miny)"""
              """    (setq xh-ww (- xh-maxx xh-minx))"""
              """    (setq xh-hh (- xh-maxy xh-miny))"""
              """    (write-line (strcat "N=" (itoa xh-n) " bbox=" (xh-rt xh-x0) "," (xh-rt xh-y0) " w=" (xh-rt xh-ww) " h=" (xh-rt xh-hh)) xh-f)"""
              """;; 图层可能锁着：能 entmake、但 entmod / 删都静默失效"""
              """    (setq xh-lo (tblobjname "LAYER" xh-lay))"""
              """    (setq xh-ld (entget xh-lo))"""
              """    (setq xh-flags (cdr (assoc 70 xh-ld)))"""
              """    (setq xh-locked (= 4 (logand 4 xh-flags)))"""
              """    (write-line (strcat "LAYER flags=" (xh-rt xh-flags) " locked=" (if xh-locked "1" "0")) xh-f)"""
              """    (if xh-locked (entmod (subst (cons 70 (- xh-flags 4)) (assoc 70 xh-ld) xh-ld)))"""
              """;; ---- 2) 文字样式：借一个内容含中文的 TEXT/MTEXT ----"""
              """    (setq xh-sty (getvar "TEXTSTYLE"))"""
              """    (setq xh-src "current")"""
              """    (setq xh-s (ssget "_X" '((0 . "TEXT,MTEXT"))))"""
              """    (if xh-s"""
              """      (progn"""
              """        (setq xh-i 0)"""
              """        (while (< xh-i (sslength xh-s))"""
              """          (if (= xh-src "current")"""
              """            (progn"""
              """              (setq xh-d (entget (ssname xh-s xh-i)))"""
              """              (setq xh-c (cdr (assoc 1 xh-d)))"""
              """              (if (and xh-c (> (strlen xh-c) 0) (> (ascii (substr xh-c 1 1)) 127))"""
              """                (progn"""
              """                  (if (assoc 7 xh-d) (setq xh-sty (cdr (assoc 7 xh-d))))"""
              """                  (setq xh-src (cdr (assoc 5 xh-d)))"""
              """                )"""
              """              )"""
              """            )"""
              """          )"""
              """          (setq xh-i (1+ xh-i))"""
              """        )"""
              """      )"""
              """    )"""
              """    (setq xh-sd (entget (tblobjname "STYLE" xh-sty)))"""
              """    (write-line (strcat "STY from=" xh-src " fixedH=" (xh-rt (if (assoc 40 xh-sd) (cdr (assoc 40 xh-sd)) 0.0))) xh-f)"""
              """;; ---- 3) 建 TEXT 并两步拟合 ----"""
              """    (setq xh-el (entmake (list '(0 . "TEXT")"""
              """                               (cons 8 xh-lay)"""
              """                               (cons 7 xh-sty)"""
              """                               (cons 10 (list xh-x0 xh-y0 0.0))"""
              """                               (cons 40 xh-hh)"""
              """                               (cons 41 1.0)"""
              """                               (cons 1 xh-txt))))"""
              """    (if xh-el"""
              """      (progn"""
              """        (setq xh-e2 (entlast))"""
              """        (setq xh-tb (textbox (entget xh-e2)))"""
              """        (if xh-tb"""
              """          (progn"""
              """            (setq xh-bw (- (caadr xh-tb) (caar xh-tb)))"""
              """            (setq xh-bh (- (cadr (cadr xh-tb)) (cadr (car xh-tb))))"""
              """            (setq xh-ix (car (car xh-tb)))"""
              """            (setq xh-iy (cadr (car xh-tb)))"""
              """            (write-line (strcat "BOX0 w=" (xh-rt xh-bw) " h=" (xh-rt xh-bh)) xh-f)"""
              """            (setq xh-sc (if (> xh-bh 0.0) (/ xh-hh xh-bh) 1.0))"""
              """            (setq xh-nh (* xh-hh xh-sc))"""
              """            (setq xh-nx (- xh-x0 (* xh-ix xh-sc)))"""
              """            (setq xh-ny (- xh-y0 (* xh-iy xh-sc)))"""
              """            (setq xh-nwf (if (> (* xh-bw xh-sc) 0.0) (/ xh-ww (* xh-bw xh-sc)) 1.0))"""
              """            (setq xh-dd (entget xh-e2))"""
              """            (setq xh-dd (subst (cons 10 (list xh-nx xh-ny 0.0)) (assoc 10 xh-dd) xh-dd))"""
              """            (setq xh-dd (subst (cons 40 xh-nh) (assoc 40 xh-dd) xh-dd))"""
              """            (setq xh-dd (subst (cons 41 xh-nwf) (assoc 41 xh-dd) xh-dd))"""
              """            (entmod xh-dd)"""
              """            (setq xh-dd (entget xh-e2))"""
              """            (setq xh-tb2 (textbox xh-dd))"""
              """            (setq xh-bh2 (- (cadr (cadr xh-tb2)) (cadr (car xh-tb2))))"""
              """            (setq xh-bw2 (- (caadr xh-tb2) (caar xh-tb2)))"""
              """            (write-line (strcat "BOX1 w=" (xh-rt xh-bw2) " h=" (xh-rt xh-bh2)) xh-f)"""
              """            (write-line (strcat "NEW h=" (cdr (assoc 5 xh-dd)) " txtlen=" (itoa (strlen (cdr (assoc 1 xh-dd)))) " height=" (xh-rt (cdr (assoc 40 xh-dd))) " wf=" (xh-rt (cdr (assoc 41 xh-dd)))) xh-f)"""
              """;; ---- 4) 墨迹高对得上才删原线；差太多说明样式把字高钉死了，留着给人看 ----"""
              """            (if (and (> xh-hh 0.0) (<= (abs (- xh-bh2 xh-hh)) (* 0.2 xh-hh)))"""
              """              (progn"""
              """                (setq xh-gone 0)"""
              """                (foreach xh-h xh-hs"""
              """                  (setq xh-x (vl-catch-all-apply (quote vlax-ename->vla-object) (list (handent xh-h))))"""
              """                  (if (not (vl-catch-all-error-p xh-x))"""
              """                    (if (not (vl-catch-all-error-p (vl-catch-all-apply (quote vla-delete) (list xh-x))))"""
              """                      (setq xh-gone (1+ xh-gone))"""
              """                    )"""
              """                  )"""
              """                )"""
              """                (write-line (strcat "ERASED=" (itoa xh-gone)) xh-f)"""
              """              )"""
              """              (write-line "SKIP-DELETE fit-out-of-tolerance" xh-f)"""
              """            )"""
              """          )"""
              """          (write-line "FAIL textbox" xh-f)"""
              """        )"""
              """      )"""
              """      (write-line "FAIL entmake" xh-f)"""
              """    )"""
              """    (if xh-locked (entmod (subst (cons 70 xh-flags) (assoc 70 (entget xh-lo)) (entget xh-lo))))"""
              """    (setq xh-s (ssget "_X" (list (cons 0 "LWPOLYLINE,ARC") (cons 8 xh-lay))))"""
              """    (write-line (strcat "CNT geom=" (itoa (if xh-s (sslength xh-s) 0))) xh-f)"""
              """    (write-line "DONE" xh-f)"""
              """  )"""
              """)"""
              """(close xh-f)"""
              """(princ)""" ]

        String.Join("\r\n", body) + "\r\n"


type AcadPdfTextTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``采集选中的 PDF 几何图形顶点``() =
        Report.title output "acad　PDF 几何图形 → 顶点转储（只读，不动图纸）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_dump.lsp")

        Directory.CreateDirectory(Path.GetDirectoryName PdfTextPick.dumpFile) |> ignore

        if File.Exists PdfTextPick.dumpFile then
            File.Delete PdfTextPick.dumpFile

        File.WriteAllText(lspFile, PdfTextDump.script PdfTextPick.handles PdfTextPick.dumpFile, Encoding.ASCII)

        Report.line output "脚本" (sprintf "%s（%d 条句柄）" lspFile PdfTextPick.handles.Length)
        Report.line output "落点" PdfTextPick.dumpFile

        let work =
            task {
                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                let identity =
                    STA.run(fun () -> sprintf "PID %d｜%s" pid acad.ActiveDocument.Name)

                Report.line output "现场" identity

                let! outcome =
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.runLispFileAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                Report.line output "LISP" (sprintf "收尾=%A" outcome)

                let! raw = PdfTextDump.awaitSettled PdfTextPick.dumpFile 15000
                let! items = PdfTextDump.load PdfTextPick.dumpFile
                let total = List.length items
                let lines = raw.Split([| '\r'; '\n' |], StringSplitOptions.RemoveEmptyEntries)

                Report.line output "实体" (sprintf "%d 条（转储 %d 行）" total lines.Length)

                for item in items do
                    Report.line
                        output
                        (sprintf "  句柄 %s" item.Handle)
                        (sprintf
                            "%s　图层 %s　标志 %d　%d 顶点　%.3f × %.3f"
                            item.Kind
                            item.Layer
                            item.Flags
                            item.Points.Length
                            item.Width
                            item.Height)

                for line in lines do
                    if line.StartsWith "DWG=" || line.StartsWith "ACADVER=" then
                        Report.line output "自报" line

                Should.equal PdfTextPick.handles.Length total

                Report.conclusion output "通过：9 条多段线的顶点全部落到结果文件，图纸里一个实体都没动"
            }

        work.GetAwaiter().GetResult()

    [<Fact>]
    member this.``选择原位转化为文字``() =
        Report.title output "acad　PDF 几何图形 → 原位 TEXT（一张 .lsp 写完就收，不保存）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"
        let armed = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_TEXT"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做")

        if armed <> "1" then
            Assert.Skip("这条会改用户的图纸（写 TEXT + 删原多段线）：确实要做就设 XUNITHOST_ACAD_TEXT=1")

        if not (File.Exists PdfTextPick.dumpFile) then
            Assert.Skip("先跑 *顶点* 把顶点转储出来，这条要拿它的包围盒定位")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_write.lsp")
        let logFile = PdfTextPick.dumpFile.Replace("acad_pdftext_dump", "acad_pdftext_write")

        if File.Exists logFile then
            File.Delete logFile

        let work =
            task {
                let! items = PdfTextDump.load PdfTextPick.dumpFile
                let x0, y0, w, h = PdfTextDump.extent items

                Report.line output "字串" (sprintf "「%s」（%d 条多段线合起来）" PdfTextPick.word PdfTextPick.handles.Length)
                Report.line output "包围盒" (sprintf "左下 %.4f, %.4f　宽 %.4f　高 %.4f" x0 y0 w h)

                File.WriteAllText(
                    lspFile,
                    PdfTextDump.writeScript PdfTextPick.word PdfTextPick.handles x0 y0 w h logFile,
                    Encoding.ASCII
                )

                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                Report.line output "现场" (STA.run(fun () -> sprintf "PID %d｜%s" pid acad.ActiveDocument.Name))

                let! outcome =
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.runLispFileAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                Report.line output "LISP" (sprintf "收尾=%A" outcome)

                let! _raw = PdfTextDump.awaitSettled logFile 15000
                let! _encoding, rows = TextFile.decodeAllLinesAsync logFile

                for row in rows do
                    Report.line output "  回执" row

                let has (prefix: string) = rows |> Array.exists(fun row -> row.StartsWith prefix)

                Should.equal true (has "LAYER=")
                Should.equal true (rows |> Array.exists(fun row -> row.Contains "NEW handle="))
                Should.equal true (rows |> Array.exists(fun row -> row.Contains(sprintf "txt=%s" PdfTextPick.word)))
                Should.equal true (rows |> Array.exists(fun row -> row.StartsWith "ERASED="))
                Should.equal true (has "RELOCK=")

                Report.conclusion output "通过：新 TEXT 按原包围盒拟合落位、9 条多段线已删、图层锁状态已还原；UNDO 可整段撤销，图纸未保存"
            }

        work.GetAwaiter().GetResult()

    /// 收尾：把剩下的多段线删净，只靠 POLY 计数与 TEXT 状态说话。
    ///
    /// 两条 LISP 事实（都实测过）：
    ///   · `handent` 对被删实体照样返回句柄 —— 它只能证明「句柄认得」，证明不了「还在图里」；
    ///   · `vlax-ename->vla-object` 遇到被删实体会**报错**，所以判在不在必须包 `vl-catch-all-apply`，
    ///     否则脚本会在写第一条状态时整段中断（日志只留半行，结果文件还被句柄锁住）。
    [<Fact>]
    member this.``删净剩下的多段线``() =
        Report.title output "acad　删净 9 条多段线：POLY 计数 + 逐条 try 转换"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"
        let armed = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_DIAG"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID")

        if armed <> "1" then
            Assert.Skip("会删实体：设 XUNITHOST_ACAD_DIAG=1 才做")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_erase.lsp")

        // 上一轮 LISP 报错时会把结果文件留在 acad 手里把着锁，Delete 会抛 —— 每轮换一个新名字
        let logFile =
            Path.Combine(
                Path.GetDirectoryName PdfTextPick.dumpFile,
                sprintf "acad_pdftext_erase_%s.txt" (DateTime.Now.ToString "HHmmss")
            )

        let quoted = PdfTextPick.handles |> List.map(sprintf "\"%s\"") |> String.concat " "

        let body =
            [ """;; 判「还在不在图里」：能转成 VLA 对象就在，转不动就是被删了"""
              """(vl-load-com)"""
              """(defun xh-rt (x) (rtos x 2 8))"""
              """(defun xh-st (h / e r)"""
              """  (setq e (handent h))"""
              """  (if (null e)"""
              """    "gone" """
              """    (progn"""
              """      (setq r (vl-catch-all-apply 'vlax-ename->vla-object (list e)))"""
              """      (if (vl-catch-all-error-p r) "erased" "alive")"""
              """    )"""
              """  )"""
              """)"""
              """(defun xh-cnt (lay / s)"""
              """  (setq s (ssget "_X" (list (cons 0 "LWPOLYLINE") (cons 8 lay))))"""
              """  (if s (sslength s) 0)"""
              """)"""
              sprintf """(setq xh-hs '(%s))""" quoted
              sprintf """(setq xh-f (open "%s" "w"))""" (logFile.Replace('\\', '/'))
              """(setq xh-lay (cdr (assoc 8 (entget (handent "1255")))))"""
              """(write-line (strcat "POLY0=" (xh-rt (xh-cnt xh-lay))) xh-f)"""
              """(foreach xh-h xh-hs"""
              """  (if (= (xh-st xh-h) "alive")"""
              """    (vl-catch-all-apply 'vla-delete (list (vlax-ename->vla-object (handent xh-h))))"""
              """  )"""
              """)"""
              """(write-line (strcat "POLY1=" (xh-rt (xh-cnt xh-lay))) xh-f)"""
              """(foreach xh-h xh-hs (write-line (strcat "H=" xh-h " " (xh-st xh-h)) xh-f))"""
              """(setq xh-s (ssget "_X" (list (cons 0 "TEXT") (cons 1 "250EXS6"))))"""
              """(write-line (strcat "TEXT-count=" (xh-rt (if xh-s (sslength xh-s) 0))) xh-f)"""
              """(if xh-s"""
              """  (progn"""
              """    (setq xh-d (entget (ssname xh-s 0)))"""
              """    (write-line (strcat "TEXT handle=" (cdr (assoc 5 xh-d)) " layer=" (cdr (assoc 8 xh-d)) " h=" (xh-rt (cdr (assoc 40 xh-d))) " wf=" (xh-rt (cdr (assoc 41 xh-d))) " ins=" (xh-rt (car (cdr (assoc 10 xh-d)))) "," (xh-rt (cadr (cdr (assoc 10 xh-d))))) xh-f)"""
              """  )"""
              """)"""
              """(close xh-f)"""
              """(princ)""" ]

        File.WriteAllText(lspFile, String.Join("\r\n", body) + "\r\n", Encoding.ASCII)

        let work =
            task {
                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                let! outcome =
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.runLispFileAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                Report.line output "LISP" (sprintf "收尾=%A" outcome)

                let! _raw = PdfTextDump.awaitSettled logFile 15000
                let! _encoding, rows = TextFile.decodeAllLinesAsync logFile

                for row in rows do
                    Report.line output "  回执" row

                Should.equal true (rows |> Array.exists(fun row -> row.StartsWith "POLY1="))
                Should.equal true (rows |> Array.exists(fun row -> row.StartsWith "TEXT-count=1"))

                Report.conclusion output "POLY0 → POLY1 的下降量就是真删掉的数量；逐条 erased/gone 即告干净"
            }

        work.GetAwaiter().GetResult()
    /// PDFSHXTEXT 这条路：不认字、不拟合，让 AutoCAD 自己拿几何去比 SHX 字库。
    /// 认出来就把几何换成 MTEXT；认不出来 POLY / TEXT 计数一动不动 —— 判据只看计数，不看外观。
    [<Fact>]
    member this.``PDFSHXTEXT 原位替字``() =
        Report.title output "acad　PDFSHXTEXT 直接替字（先转储退路，再跑识别；一张 .lsp 发出去不等）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"
        let armed = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_SHX"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做")

        if armed <> "1" then
            Assert.Skip("这条会改用户的图纸（换文字 + 删原线）：确实要做就设 XUNITHOST_ACAD_SHX=1")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_shx.lsp")

        // 上一轮若被卡住，结果文件还在 acad 手里把着锁 —— 删不掉就算了，LISP 会覆盖写
        for path in [ ShxPick.dumpFile; ShxPick.logFile ] do
            if File.Exists path then
                try
                    File.Delete path
                with _ ->
                    ()

        /// `CNT before poly=5 text=1` → (5, 1)；没这行就给 (-1, -1)
        let count (prefix: string) (rows: string[]) =
            rows
            |> Array.tryFind(fun row -> row.StartsWith prefix)
            |> Option.map(fun row ->
                let parts = row.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)

                let numOf (key: string) =
                    parts
                    |> Array.tryFind(fun part -> part.StartsWith key)
                    |> Option.map(fun part -> int (part.Substring key.Length))
                    |> Option.defaultValue -1

                numOf "poly=", numOf "text=")
            |> Option.defaultValue (-1, -1)

        /// `CNT all before=1` → 1；没有这行给 -1
        let total (prefix: string) (rows: string[]) =
            rows
            |> Array.tryFind(fun row -> row.StartsWith prefix)
            |> Option.map(fun row -> int (row.Substring prefix.Length))
            |> Option.defaultValue -1

        File.WriteAllText(
            lspFile,
            ShxRun.script ShxPick.handles ShxPick.dumpFile ShxPick.logFile,
            Encoding.ASCII
        )

        Report.line output "脚本" (sprintf "%s（%d 条句柄）" lspFile ShxPick.handles.Length)
        Report.line output "落点" (sprintf "%s　＋　%s" ShxPick.dumpFile ShxPick.logFile)

        let work =
            task {
                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                Report.line output "现场" (STA.run(fun () -> sprintf "PID %d｜%s" pid acad.ActiveDocument.Name))

                // 发出去就不等：PDFSHXTEXT 结束会报「N of M groups converted」，万一走模态框，
                // 求值就停在那儿，等 EndLisp 会一直挂着 —— 自己盯文件更稳。
                do!
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.loadLispAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                let! _dump = PdfTextDump.awaitSettled ShxPick.dumpFile 20000

                Report.line output "转储" "顶点已落盘（这一段做完才去跑 PDFSHXTEXT，所以它一定在）"

                let mutable rows: string[] = [||]

                try
                    let! _log = PdfTextDump.awaitSettled ShxPick.logFile 40000
                    let! _encoding, decoded = TextFile.decodeAllLinesAsync ShxPick.logFile
                    rows <- decoded
                with ex ->
                    Report.line output "日志" (sprintf "没等到闭合：%s（大概率弹了提示框，屏幕上点掉再看）" ex.Message)

                for row in rows do
                    Report.line output "  回执" row

                let p0, t0 = count "CNT before" rows
                let p1, t1 = count "CNT after" rows
                let a0 = total "CNT all before=" rows
                let a1 = total "CNT all after=" rows

                let tally = sprintf "多段线 %d → %d　本层文字 %d → %d　全图文字 %d → %d" p0 p1 t0 t1 a0 a1
                Report.line output "计数" tally

                // 判据只看全图：PDFSHXTEXT 把 MTEXT 建在**当前图层**，不是来源图层，
                // 按本层数字说话会把「识别成功」误判成「删了线没写字」。
                Should.equal 5 (p0 - p1)
                Should.equal true (a1 > a0)

                let texts =
                    rows
                    |> Array.filter(fun row -> row.StartsWith "TXT=")
                    |> Array.map(fun row -> row.Substring 4)

                Report.conclusion
                    output
                    (sprintf
                        "通过：5 条几何已被 PDFSHXTEXT 换掉，这一层留下的文字是「%s」；图纸未保存"
                        (String.Join("；", texts)))
            }

        work.GetAwaiter().GetResult()
    /// `PDFSHXTEXT` 把 5 条几何吃掉了，却只在原图层留下那一条旧的 250EXS6 —— 这张兜底收尾。
    /// 区域内（不限图层）有文字就搬回原图层；一条都没有才按原包围盒补写 `SS304` 并两步拟合。
    [<Fact>]
    member this.``PDFSHXTEXT 之后补写``() =
        Report.title output "acad　兜底：找回 PDFSHXTEXT 建到别处的文字，找不到就按 bbox 补写 SS304"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"
        let armed = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_SHX"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做")

        if armed <> "1" then
            Assert.Skip("这条会改用户的图纸：确实要做就设 XUNITHOST_ACAD_SHX=1")

        if not (File.Exists ShxPick.dumpFile) then
            Assert.Skip("先跑 *PDFSHXTEXT* 把 5 条的顶点转储出来，这条要拿它的包围盒定位")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_shx_rescue.lsp")
        let logFile = ShxPick.logFile.Replace("acad_shx_log", "acad_shx_rescue")

        if File.Exists logFile then
            try
                File.Delete logFile
            with _ ->
                ()

        let work =
            task {
                let! items = PdfTextDump.load ShxPick.dumpFile
                let x0, y0, w, h = PdfTextDump.extent items

                Report.line output "字串" (sprintf "「%s」（本地渲图认出：S S 3 0 4）" ShxPick.word)
                Report.line output "包围盒" (sprintf "左下 %.8f, %.8f　宽 %.8f　高 %.8f" x0 y0 w h)

                File.WriteAllText(
                    lspFile,
                    ShxRun.rescueScript ShxPick.word x0 y0 w h ShxPick.layerFrom logFile,
                    Encoding.ASCII
                )

                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                Report.line output "现场" (STA.run(fun () -> sprintf "PID %d｜%s" pid acad.ActiveDocument.Name))

                do!
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.loadLispAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                let mutable rows: string[] = [||]

                try
                    let! _log = PdfTextDump.awaitSettled logFile 40000
                    let! _encoding, decoded = TextFile.decodeAllLinesAsync logFile
                    rows <- decoded
                with ex ->
                    Report.line output "日志" (sprintf "没等到闭合：%s" ex.Message)

                for row in rows do
                    Report.line output "  回执" row

                let has (prefix: string) = rows |> Array.exists(fun row -> row.StartsWith prefix)

                Should.equal true (has "DONE")
                Should.equal true (has "FOUND" || rows |> Array.exists(fun row -> row.Contains "NEW h="))

                Report.conclusion
                    output
                    "通过：区域内的文字已归位到原图层（或按 bbox 补写了 SS304）；图纸未保存"
            }

        work.GetAwaiter().GetResult()
    /// 55 条几何 → 一条中文 TEXT。一次访问干完：LISP 自己量 bbox（含 ARC）、借中文样式、
    /// 两步拟合、墨迹高对得上才删原线。中文靠 `(chr n)` 拼 GBK 字节，脚本本身保持纯 ASCII。
    [<Fact>]
    member this.``55 条原位转化为中文``() =
        Report.title output "acad　55 条 PDF 几何 → 一条中文 TEXT（一次访问，bbox 在 LISP 里算）"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"
        let armed = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_CN"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做")

        if armed <> "1" then
            Assert.Skip("这条会改用户的图纸（写 TEXT + 删 55 条线）：确实要做就设 XUNITHOST_ACAD_CN=1")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_cn.lsp")

        for path in [ BigPick.dumpFile; BigPick.logFile ] do
            if File.Exists path then
                try
                    File.Delete path
                with _ ->
                    ()

        File.WriteAllText(
            lspFile,
            BigRun.script BigPick.handles BigPick.wordBytes BigPick.dumpFile BigPick.logFile,
            Encoding.ASCII
        )

        Report.line
            output
            "字串"
            (sprintf "「%s」　%d 个字符 / %d 个 GBK 字节" BigPick.word BigPick.word.Length BigPick.wordBytes.Length)

        Report.line
            output
            "条数"
            (sprintf "%d 条（%s … %s）" BigPick.handles.Length (List.head BigPick.handles) (List.last BigPick.handles))

        let work =
            task {
                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                Report.line output "现场" (STA.run(fun () -> sprintf "PID %d｜%s" pid acad.ActiveDocument.Name))

                do!
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.loadLispAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                // 顶点转储先落盘：后面万一卡住，重建原线的原料还在
                let! _dump = PdfTextDump.awaitSettled BigPick.dumpFile 20000

                Report.line output "转储" "55 条顶点已落盘（改之前的快照）"

                let mutable rows: string[] = [||]

                try
                    let! _log = PdfTextDump.awaitSettled BigPick.logFile 40000
                    let! _encoding, decoded = TextFile.decodeAllLinesAsync BigPick.logFile
                    rows <- decoded
                with ex ->
                    Report.line output "日志" (sprintf "没等到闭合：%s" ex.Message)

                for row in rows do
                    Report.line output "  回执" row

                let has (prefix: string) = rows |> Array.exists(fun row -> row.StartsWith prefix)
                let any (prefixes: string list) = prefixes |> List.exists has

                Should.equal true (has "DONE")
                Should.equal true (any [ "STRLEN=8"; "STRLEN=14" ])
                Should.equal true (has "ERASED=55")

                Report.conclusion output "通过：55 条几何已换成一条中文 TEXT（墨迹盒按原 bbox 拟合，图层未变）；图纸未保存"
            }

        work.GetAwaiter().GetResult()
    /// 只读校验：把新写那条（3D99）与图纸里原有的中文文字（3D03）的 `(assoc 1)` **逐字节**取回来。
    /// 这一版 AutoLISP 的字符串是 ANSI 字节串（`(ascii (substr s 1 1))` 能拿到 >127 就是证据），
    /// 所以「字节序列 == 目标字串的 GBK 编码」就能证明存进去的是 GBK、不是被逐字节拆坏的乱码。
    [<Fact>]
    member this.``校验中文按 GBK 存进去``() =
        Report.title output "acad　只读回读：3D99 的 group 1 逐字节 vs 目标字串的 GBK 编码"

        let pidText = Environment.GetEnvironmentVariable "XUNITHOST_ACAD_PID"

        if String.IsNullOrWhiteSpace pidText then
            Assert.Skip("没设 XUNITHOST_ACAD_PID：这条要在活着的 acad 上做")

        let pid = int pidText
        let lspFile = Path.Combine(Path.GetTempPath(), "acad_pdftext_cnbytes.lsp")
        let logFile = ShxPick.logFile.Replace("acad_shx_log", "acad_cn_bytes")

        if File.Exists logFile then
            try
                File.Delete logFile
            with _ ->
                ()

        let body =
            [ """;; 纯 ASCII；只读，不改图纸"""
              """(vl-load-com)"""
              """(defun xh-bytes (s / i out)"""
              """  (setq out "" i 1)"""
              """  (while (<= i (strlen s))"""
              """    (setq out (strcat out (itoa (ascii (substr s i 1))) " "))"""
              """    (setq i (1+ i))"""
              """  )"""
              """  out"""
              """)"""
              sprintf """(setq xh-f (open "%s" "w"))""" (logFile.Replace('\\', '/'))
              """(foreach xh-h '("3D99" "3D03")"""
              """  (setq xh-e (handent xh-h))"""
              """  (if xh-e"""
              """    (write-line (strcat xh-h "|" (xh-bytes (cdr (assoc 1 (entget xh-e))))) xh-f)"""
              """    (write-line (strcat xh-h "|MISSING") xh-f)"""
              """  )"""
              """)"""
              """(write-line "DONE" xh-f)"""
              """(close xh-f)"""
              """(princ)""" ]

        File.WriteAllText(lspFile, String.Join("\r\n", body) + "\r\n", Encoding.ASCII)

        let work =
            task {
                let! acad = AutoCADHost.AutoCADTask.GetApplicationAsync pid

                Report.line output "现场" (STA.run(fun () -> sprintf "PID %d｜%s" pid acad.ActiveDocument.Name))

                do!
                    STA.runAsync(fun () -> AutoCADHost.AutoCADTask.loadLispAsync acad.ActiveDocument lspFile)
                        .Unwrap()

                let! _raw = PdfTextDump.awaitSettled logFile 20000
                let! _encoding, rows = TextFile.decodeAllLinesAsync logFile

                for row in rows do
                    Report.line output "  回执" row

                let bytesOf (handle: string) =
                    rows
                    |> Array.tryFind(fun row -> row.StartsWith(handle + "|"))
                    |> Option.map(fun row ->
                        row.Substring(handle.Length + 1)
                        |> fun s -> s.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
                        |> Array.map int
                        |> Array.toList)
                    |> Option.defaultValue []

                let actual = bytesOf "3D99"
                let neighbour = bytesOf "3D03"

                Report.line output "目标" (BigPick.wordBytes |> List.map string |> String.concat " ")
                Report.line output "实际" (actual |> List.map string |> String.concat " ")
                Report.line output "旁证" (sprintf "3D03 共 %d 字节（图纸原有中文，作对照）" neighbour.Length)

                Should.equal BigPick.wordBytes actual

                Report.conclusion output "通过：新文字存的就是目标字串的 GBK 字节，逐字节一致"
            }

        work.GetAwaiter().GetResult()
