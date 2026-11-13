/// 进程账本：「谁起的谁关」全靠它。
[<RequireQualifiedAccess>]
module OsHost.Process

open System.Diagnostics
open System.Threading

/// 当前同名进程的进程号集合（image 不带 .exe，例："WINWORD" / "EXCEL" / "acad"）
let ids (image: string) : Set<int> =
    try
        Process.GetProcessesByName(image) |> Array.map(fun p -> p.Id) |> Set.ofArray
    with _ ->
        Set.empty

/// 进程何时启动（Unix 毫秒）
let startedMs (pid: int) : int64 option =
    if pid <= 0 then
        None
    else
        try
            Some(Clock.unixMs (Process.GetProcessById pid).StartTime)
        with _ ->
            None

/// 进程的可执行文件路径（核对黑板上的 EXE 路径用）；取不到返回 None。
let exePath (pid: int) : string option =
    if pid <= 0 then
        None
    else
        try
            use proc = Process.GetProcessById pid
            Some proc.MainModule.FileName
        with _ ->
            None

/// 等这些进程真的消失（Quit 是异步的：只投递 WM_QUIT，进程还要一会儿才走）。
/// 返回超时后还活着的那些。
let waitGone (image: string) (pids: Set<int>) (timeoutMs: int) : Set<int> =
    let clock = Stopwatch.StartNew()
    let mutable rest = Set.intersect pids (ids image)

    while not rest.IsEmpty && clock.ElapsedMilliseconds < int64 timeoutMs do
        Thread.Sleep 100
        rest <- Set.intersect pids (ids image)

    rest

/// 结束进程（进程外服务器 Quit 之后常常赖着不走，只能来硬的）
let kill (pids: Set<int>) : unit =
    for pid in pids do
        try
            (Process.GetProcessById pid).Kill()
        with _ ->
            ()

/// 收拾「这次自己起的」实例：起之前就在跑的（用户手上的）一概不碰。
/// 先请它自己退（quit），graceMs 内还不走就直接结束。
/// 返回被强制结束的进程名 —— 有返回值就说明哪一步卡住了（正常退出返回 None）。
let shutdown (image: string) (graceMs: int) (before: Set<int>) (quit: unit -> unit) : string option =
    let ours = Set.difference (ids image) before

    if ours.IsEmpty then
        None
    else
        (try
            quit ()
            with _ ->
                ())

        let rest = waitGone image ours graceMs
        kill rest

        if rest.IsEmpty then None else Some image

