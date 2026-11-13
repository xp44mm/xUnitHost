/// 窗口 → 进程：判断「这只应用对象到底是哪个进程的」就靠这两个调用。
[<RequireQualifiedAccess>]
module OsHost.Win32

open System
open System.Collections.Concurrent
open System.Diagnostics
open System.IO
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open System.Threading
open System.Threading.Tasks

/// GA_ROOT = 2：从文档/子窗口往上找顶层窗口
[<DllImport("user32.dll")>]
extern nativeint GetAncestor(nativeint hwnd, uint32 flags)

[<DllImport("user32.dll")>]
extern uint32 GetWindowThreadProcessId(nativeint hwnd, uint32& pid)

/// 从任意子窗口往上找顶层窗口（Word 的 Application 没有 Hwnd，只能从文档窗口往上找 OpusApp）
let root (hwnd: nativeint) : nativeint =
    if hwnd = 0n then
        0n
    else
        GetAncestor(hwnd, 2u)

/// 这个窗口属于哪个进程；判不出来（没有窗口 / 拿不到）就是 None
let pidOfWindow (hwnd: nativeint) : int option =
    if hwnd = 0n then
        None
    else
        let mutable pid = 0u
        GetWindowThreadProcessId(hwnd, &pid)
        |> ignore
        if pid = 0u then None else Some(int pid)

/// SW_RESTORE = 9：最小化的窗口还原出来（只 ShowWindow 不还原，切过去还是最小化）
[<Literal>]
let SwRestore = 9

[<DllImport("user32.dll")>]
extern bool ShowWindow(nativeint hwnd, int cmdShow)

[<DllImport("user32.dll")>]
extern bool SetForegroundWindow(nativeint hwnd)

/// 把这只进程的主窗口还原并切到前台，返回前台切换是否成功。
/// Windows 有前台锁（用户正在别的程序里输入时不许抢），失败是正常结果、不是错误；
/// 调用方按返回值决定要不要提示「请自己点任务栏」。
let bringToFront (pid: int) : bool =
    try
        let proc = System.Diagnostics.Process.GetProcessById pid
        let hwnd = proc.MainWindowHandle

        if hwnd = IntPtr.Zero then
            false
        else
            ShowWindow(nativeint hwnd, SwRestore)
            |> ignore
            SetForegroundWindow(nativeint hwnd)
    with _ ->
        false
