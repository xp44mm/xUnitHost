/// COM 底层（不认任何 Office 类型，只认 ProgID 和 IUnknown）。
/// 不走 Marshal.GetActiveObject / Type.GetTypeFromProgID：netstandard2.0 里都没有这两个 API，
/// 直接 P/Invoke ole32 / oleaut32，行为在 .NET Framework 与 .NET 5+ 上完全一致。
[<RequireQualifiedAccess>]
module OsHost.Com 

open System
open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes

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

/// 从 ROT 里取现成实例：一个 ProgID 一个 ProgID 地试，都不行就 None（绝不新建，也不该抛异常）
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

[<DllImport("ole32.dll")>]
extern int CreateBindCtx(int reserved, nativeint& ppbc)

/// 造一个绑定上下文（读 moniker 的显示名要它）
let private bindContext () : IBindCtx option =
    let mutable punk = IntPtr.Zero

    if CreateBindCtx(0, &punk) <> 0 then
        None
    else
        match comObject punk with
        | Some raw ->
            (try
                Some(raw :?> IBindCtx)
                with _ ->
                None)
        | None -> None

/// 枚举 ROT 当前登记的全部对象，逐条给 (显示名, 对象)。
///
/// 只在「调用方只有 PID、没有精确 ROT 注册名」时才用，而且**必须自己核身份**
/// （哪一家的 host 就用哪一家的窗口 → PID 比对），绝不能拿第一条就当目标。
/// 黑板模式有精确名字，仍走 runningAt，不枚举（见 references/gotchas.md）。
/// 单条取名字 / 绑定失败就跳过那一条；绝不新建实例。
let runningEntries () : (string * obj) list =
    try
        match runningTable () with
        | None -> []
        | Some table ->
            let ctx = bindContext ()
            let mutable enumerator: IEnumMoniker = Unchecked.defaultof<IEnumMoniker>
            table.EnumRunning(&enumerator)

            if obj.ReferenceEquals(enumerator, null) then
                []
            else
                let found = ResizeArray<string * obj>()
                let batch: IMoniker array = Array.zeroCreate 1
                let mutable go = true

                while go && found.Count < 512 do
                    batch.[0] <- Unchecked.defaultof<IMoniker>

                    if enumerator.Next(1, batch, IntPtr.Zero) <> 0 || obj.ReferenceEquals(batch.[0], null) then
                        go <- false
                    else
                        let moniker = batch.[0]

                        let name =
                            try
                                match ctx with
                                | Some context ->
                                    let mutable display: string = null
                                    moniker.GetDisplayName(context, null, &display)
                                    if String.IsNullOrWhiteSpace display then "-" else display
                                | None -> "-"
                            with _ ->
                                "-"

                        try
                            let mutable value: obj = null

                            if table.GetObject(moniker, &value) = 0 && not (isNull value) then
                                found.Add((name, value))
                        with _ ->
                            ()

                List.ofSeq found
    with _ ->
        []

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



