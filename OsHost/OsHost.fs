/// 操作系统这一层：线程（STA）、进程（账本 / 启动时间 / 等它退 / 结束它）、
/// 窗口（hwnd → 进程号）、时间（Unix 毫秒）、文件（规整路径 / 最后写入时间）、
/// COM 底层（运行对象表 / CoCreateInstance）、异常归因。
///
/// 列在这里的都是「三家 host 一模一样、写三遍纯属抄书」的活儿：
///   · Word / Excel / AutoCAD 的自动化都要求 STA，起线程的方式没区别；
///   · 「谁起的谁关」都靠「起前起后各扫一次同名进程，差集就是这次起的」；
///   · 「拿到的是不是那一只」都靠「窗口 → 进程号」；
///   · 地址里的「最后保存毫秒」都取自磁盘文件时间。
///
/// 本库不认 Office 类型：只认句柄、进程号、IUnknown 和磁盘上的文件。
namespace OsHost

open System
open System.Collections.Concurrent
open System.Diagnostics
open System.IO
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes
open System.Threading
open System.Threading.Tasks

/// 一个「现场」的身份 —— 地址里的指纹就是从它来的。
/// 三家 host 各自拍一张（各自知道自己的活动文档在哪儿），形状统一在这儿定。
type Identity =
    { /// 进程号；0 = 判不出来
      Pid: int
      /// 进程启动时间（Unix 毫秒）。pid 会被系统复用，加上它才认得出「是不是原来那只」
      StartedMs: int64
      /// 锚定的文档（活动文档）落盘路径；新文档还没存过就是 None
      DocPath: string option
      /// 最后一次实质性保存：文档文件在磁盘上的最后写入时间（Unix 毫秒）；0 = 还没保存过
      SavedMs: int64
      /// 有没有没保存的改动。打字就会变，只用来提示，**不参与核对**
      Dirty: bool }

/// STA 线程。两种用法，别混：
///   · `run`  —— 长活：固定一条后台线程排队干活（host 开着的那段时间都用它）；
///   · `here` —— 短活：当前线程已经是 STA 就地跑，不是就临时起一条，跑完就收。
[<RequireQualifiedAccess>]
module Sta =

    let private jobs = new BlockingCollection<unit -> unit>()

    let private started =
        lazy
            (let worker =
                Thread(
                    (fun () ->
                        for job in jobs.GetConsumingEnumerable() do
                            job ()),
                    IsBackground = true
                )

             worker.SetApartmentState ApartmentState.STA
             worker.Start())

    /// 长活：所有 COM 调用排在同一个 STA 线程上做。
    /// host 是长活的：RCW 一旦跨公寓使用就要封送，而封送要求创建它的 STA 线程在泵消息 ——
    /// 干脆固定一条线程，open / exec / close 全在它上面跑，谁调用都不会串公寓。
    let run (work: unit -> 'T) : 'T =
        started.Force() |> ignore
        let pending = TaskCompletionSource<obj>()

        jobs.Add(fun () ->
            try
                pending.SetResult(box (work ()))
            with ex ->
                pending.SetException ex)

        unbox<'T> (pending.Task.GetAwaiter().GetResult())

    /// 短活：当前线程不是 STA 时，把活儿丢给一条临时 STA 线程跑完再回来
    let here (work: unit -> 'T) : 'T =
        if Thread.CurrentThread.GetApartmentState() = ApartmentState.STA then
            work ()
        else
            let mutable result = Unchecked.defaultof<'T>
            let mutable failure: exn = null

            let body =
                ThreadStart(fun () ->
                    try
                        result <- work ()
                    with ex ->
                        failure <- ex)

            let worker = Thread(body)
            worker.SetApartmentState ApartmentState.STA
            worker.Start()
            worker.Join()

            if not (isNull failure) then
                raise failure

            result

/// 时间：地址里一律用 Unix 毫秒（跨进程、跨机器都没有歧义），报告里才换成人能读的。
[<RequireQualifiedAccess>]
module Clock =

    let private epochTicks = DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks

    let unixMs (t: DateTime) : int64 = (t.ToUniversalTime().Ticks - epochTicks) / 10000L

    /// 0 就是「还没保存过」，不摆一个 1970 出来
    let show (ms: int64) : string =
        if ms <= 0L then
            "还没保存过"
        else
            DateTimeOffset.FromUnixTimeMilliseconds(ms).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss.fff")

/// 文件：文档路径与「最后一次实质性保存」都从磁盘上问，不听内存里的标记。
[<RequireQualifiedAccess>]
module Paths =

    /// 文档路径规整成绝对路径；新文档（没存过盘）返回 None。
    /// Word 的新文档 FullName 是 ""，Excel 的是 "Book1"，AutoCAD 的是 "Drawing1.dwg"。
    let canonical (raw: string) : string option =
        if String.IsNullOrWhiteSpace raw then
            None
        else
            try
                if Path.IsPathRooted raw then Some(Path.GetFullPath raw) else None
            with _ ->
                None

    /// 「最后一次实质性保存」的判据：文档文件在磁盘上的最后写入时间。
    /// 不用 Document.Saved 之类的内存标记 —— 那个只说明「有没有未保存改动」，
    /// 说不出「上次保存是什么时候」；而磁盘时间正是保存真的落了盘的证据。
    let savedMs (path: string option) : int64 =
        match path with
        | Some p when not (String.IsNullOrWhiteSpace p) ->
            try
                let file = FileInfo p
                if file.Exists then Clock.unixMs file.LastWriteTimeUtc else 0L
            with _ ->
                0L
        | _ -> 0L

/// 进程账本：「谁起的谁关」全靠它。
[<RequireQualifiedAccess>]
module Proc =

    /// 当前同名进程的进程号集合（image 不带 .exe，例："WINWORD" / "EXCEL" / "acad"）
    let ids (image: string) : Set<int> =
        try
            Process.GetProcessesByName(image) |> Array.map(fun p -> p.Id) |> Set.ofArray
        with _ ->
            Set.empty

    /// 进程启动时间（Unix 毫秒）：pid 复用检测的锚。拿不到返回 None
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

/// 窗口 → 进程：判断「这只应用对象到底是哪个进程的」就靠这两个调用。
[<RequireQualifiedAccess>]
module Win32 =

    /// GA_ROOT = 2：从文档/子窗口往上找顶层窗口
    [<DllImport("user32.dll")>]
    extern nativeint GetAncestor(nativeint hwnd, uint32 flags)

    [<DllImport("user32.dll")>]
    extern uint32 GetWindowThreadProcessId(nativeint hwnd, uint32& pid)

    /// 从任意子窗口往上找顶层窗口（Word 的 Application 没有 Hwnd，只能从文档窗口往上找 OpusApp）
    let root (hwnd: nativeint) : nativeint =
        if hwnd = 0n then 0n else GetAncestor(hwnd, 2u)

    /// 这个窗口属于哪个进程；判不出来（没有窗口 / 拿不到）就是 None
    let pidOfWindow (hwnd: nativeint) : int option =
        if hwnd = 0n then
            None
        else
            let mutable pid = 0u
            GetWindowThreadProcessId(hwnd, &pid) |> ignore
            if pid = 0u then None else Some(int pid)

/// 异常归因：COM 异常外层多是无信息量的 "Exception from HRESULT: 0x..."，有用的在里层
[<RequireQualifiedAccess>]
module Errors =

    let explain (ex: exn) : string =
        let mutable e = ex

        while not (isNull e.InnerException) do
            e <- e.InnerException

        if String.IsNullOrWhiteSpace e.Message then
            sprintf "%s (0x%X)" (e.GetType().Name) e.HResult
        else
            e.Message

/// COM 底层（不认任何 Office 类型，只认 ProgID 和 IUnknown）。
/// 不走 Marshal.GetActiveObject / Type.GetTypeFromProgID：netstandard2.0 里都没有这两个 API，
/// 直接 P/Invoke ole32 / oleaut32，行为在 .NET Framework 与 .NET 5+ 上完全一致。
[<RequireQualifiedAccess>]
module Com =

    /// 本地服务器（EXE）：Word / Excel / AutoCAD 都是进程外服务器，永远走这个上下文
    [<Literal>]
    let private CLSCTX_LOCAL_SERVER = 4u

    [<DllImport("ole32.dll", CharSet = CharSet.Unicode)>]
    extern int CLSIDFromProgID(string progId, Guid& clsid)

    /// 从运行对象表（ROT）取已注册的实例；没有在跑时返回 MK_E_UNAVAILABLE (0x800401E3)
    [<DllImport("oleaut32.dll")>]
    extern int GetActiveObject(Guid& clsid, nativeint reserved, nativeint& punk)

    [<DllImport("ole32.dll")>]
    extern int CoCreateInstance(Guid& clsid, nativeint outer, uint32 context, Guid& iid, nativeint& punk)

    let private iidUnknown = Guid "00000000-0000-0000-C000-000000000046"

    /// 按注册名精确读 ROT 用的两条 API。插件侧用 CreateItemMoniker("!", 名字) + Register 登记，
    /// 消费侧造同一个 moniker 用 GetObject 取回来 —— 不枚举、不用 GetActiveObject 猜实例。
    [<DllImport("ole32.dll")>]
    extern int GetRunningObjectTable(int reserved, nativeint& pprot)

    [<DllImport("ole32.dll", CharSet = CharSet.Unicode)>]
    extern int CreateItemMoniker(string delim, string name, nativeint& ppmk)

    /// 把 P/Invoke 交回来的 IUnknown 裸指针变成 RCW，并把原生那一份引用放掉。
    let private comObject (punk: nativeint) : obj option =
        if punk = IntPtr.Zero then
            None
        else
            try
                Some(Marshal.GetObjectForIUnknown punk)
            finally
                // GetObjectForIUnknown 自己加过一次引用，这里放开 P/Invoke 交回来的那份
                Marshal.Release punk |> ignore

    /// 打开 ROT，拿到 IRunningObjectTable（RCW 由 GC 收，原生引用已放掉）
    let private runningTable () : IRunningObjectTable option =
        let mutable punk = IntPtr.Zero

        if GetRunningObjectTable(0, &punk) <> 0 then
            None
        else
            match comObject punk with
            | Some raw -> Some(raw :?> IRunningObjectTable)
            | None -> None

    /// 造一个 item moniker。ROT 里的名字形如 !WordVSTO.Instance.27232，
    /// "!" 是分隔符，后面整串是对象名。
    let private itemMoniker (rotName: string) : IMoniker option =
        if String.IsNullOrWhiteSpace rotName then
            None
        else
            let delim, item =
                if rotName.StartsWith "!" then "!", rotName.Substring 1 else "!", rotName

            let mutable punk = IntPtr.Zero

            if CreateItemMoniker(delim, item, &punk) <> 0 then
                None
            else
                match comObject punk with
                | Some raw -> Some(raw :?> IMoniker)
                | None -> None

    let private withClsid (progId: string) (work: Guid -> 'T option) : 'T option =
        let mutable clsid = Guid.Empty

        if CLSIDFromProgID(progId, &clsid) <> 0 then
            None
        else
            work clsid

    /// 从 ROT 里取现成实例：一个 ProgID 一个 ProgID 地试，都不行就 None（绝不新建，也不该抛）
    let running (progIds: string list) : obj option =
        progIds
        |> List.tryPick(fun progId ->
            withClsid progId (fun clsid ->
                let mutable id = clsid
                let mutable punk = IntPtr.Zero
                let hr = GetActiveObject(&id, IntPtr.Zero, &punk)

                if hr <> 0 || punk = IntPtr.Zero then
                    None
                else
                    try
                        Some(Marshal.GetObjectForIUnknown punk)
                    finally
                        // GetObjectForIUnknown 自己加过一次引用，这里放开 ROT 那份
                        Marshal.Release punk |> ignore))

    /// 起一个新实例（每次都是新的进程，不会附着到用户已开的那个）。
    /// 一个 ProgID 一个 ProgID 地试，全不行就抛 COMException。
    let create (progIds: string list) : obj =
        let one (progId: string) : obj option =
            withClsid progId (fun clsid ->
                let mutable id = clsid
                let mutable iid = iidUnknown
                let mutable punk = IntPtr.Zero
                let hr = CoCreateInstance(&id, IntPtr.Zero, CLSCTX_LOCAL_SERVER, &iid, &punk)

                if hr <> 0 then
                    Marshal.ThrowExceptionForHR hr

                try
                    Some(Marshal.GetObjectForIUnknown punk)
                finally
                    Marshal.Release punk |> ignore)

        match progIds |> List.tryPick one with
        | Some value -> value
        | None ->
            raise (
                COMException(
                    "本机没有注册 "
                    + String.Join(" / ", progIds)
                    + "：对应程序未安装，或从未成功启动过。"
                )
            )

    /// 按 ROT 注册名精确取现成实例（例：!WordVSTO.Instance.27232）。**绝不新建**：
    /// 名字没登记、登记的实例已经退出，都只返回 None。多实例机器上这是唯一不会认错人的挂法。
    let runningAt (rotName: string) : obj option =
        if String.IsNullOrWhiteSpace rotName then
            None
        else
            try
                match runningTable (), itemMoniker rotName with
                | Some table, Some moniker ->
                    let mutable value: obj = null

                    if table.GetObject(moniker, &value) <> 0 || isNull value then
                        None
                    else
                        Some value
                | _ -> None
            with _ ->
                None

    /// 自测 / 提供黑板的一侧：把自己创建的托管对象登记到 ROT（强注册），返回 cookie。
    /// 坑：拿别人的 COM 代理来登记会 E_INVALIDARG，这里只登记自己 new 出来的对象。
    let register (rotName: string) (value: obj) : Result<int, string> =
        try
            match runningTable (), itemMoniker rotName with
            | Some table, Some moniker -> Ok(table.Register(1, value, moniker))
            | _ -> Error "打开 ROT / 造 moniker 失败"
        with ex ->
            Error(Errors.explain ex)

    /// 撤掉 register 登记的条目（cookie 就是 register 返回的那个值）。
    let revoke (cookie: int) : unit =
        try
            match runningTable () with
            | Some table -> table.Revoke cookie
            | None -> ()
        with _ ->
            ()
