/// 测试项目前部的公用件：一段测试里要反复用的杂活，都在这儿写一份。
///
///   · `Report` —— 把发生的事写成报告。**写进 `output` 的那份就是交回给人的返回信息**。
///   · `Busy`   —— COM 说「我正忙」「服务器起不来」时等一下再来（AutoCAD 冷启动尤其爱来这套）。
///   · `Doc`    —— 动笔之前先认人：文件名 + 最后实质保存时间，只往没落盘的新文档里写。
///   · `Acad`   —— AutoCAD 通道的闸门：没有实例时默认不替人新起一个。
///
/// 再往下的操作系统功课（线程、进程账本、窗口 → 进程号、时间）在 `OsHost` 里 —— 那几样
/// 三个 host 库也要用，所以不能只放在测试项目里。
namespace XUnitHost

open System
open System.Diagnostics
open System.IO
open System.Threading
open Xunit
open OsHost

/// 报告：标题一行 → 事实若干行 → 最后一行结论。
module internal Report =

    /// 当前同名进程的进程号集合（转给操作系统层，测试里少写一层名字）
    let procIds (name: string) : Set<int> = Proc.ids name

    /// Word 正文末尾带段落标记，读回来统一去掉
    let marks = [| '\r'; '\a'; '\n'; ' ' |]

    /// 「第二个进程」就是**这个测试程序自己**：再起一份、只跑其中那一条 Fact。
    /// 所以要的是自己的可执行文件 —— 不用去目录里翻别人的 exe，更不用为「另一个进程」另编一个程序。
    let selfExe () : string =
        match Environment.ProcessPath with
        | null
        | "" -> Path.Combine(AppContext.BaseDirectory, "xUnitHost.exe")
        | path -> path

    /// Result 拆开：失败就把原因当异常抛出去（测试里 Error 就是失败）
    let ok (result: Result<'T, string>) : 'T =
        match result with
        | Ok value -> value
        | Error message -> failwith message

    /// 本地时间 → Unix 毫秒（和地址里的算法一致，用来核对「最后保存」那一刻）
    let unixMs (t: DateTime) : int64 = Clock.unixMs t

    /// Unix 毫秒 → 人能读的时刻（报告里不摆裸时间戳）
    let clock (ms: int64) : string = Clock.show ms

    /// 这几个 Fact 要求「自己从零起一个实例」：机器上已经有同类进程在跑就没法做
    /// （也绝不能去碰你手上的那份图纸/工作簿）。这种情况直接跳过，别报失败。
    let requireClean (image: string) : unit =
        let running = procIds image

        if not running.IsEmpty then
            let ids = running |> Set.toList |> List.map string |> String.concat ", "

            Assert.Skip(
                sprintf
                    "本机已经有 %s 在跑（pid %s），这个 Fact 要从零起一个实例；先关掉它再跑"
                    image
                    ids
            )

    /// 中文按两格宽算，「标签」那一列在等宽字体里才对得齐
    let private displayWidth (s: string) =
        s |> Seq.sumBy(fun c -> if int c >= 0x2E80 then 2 else 1)

    let private pad (width: int) (s: string) = String(' ', max 2 (width - displayWidth s))

    /// 报告标题
    let title (output: ITestOutputHelper) (text: string) =
        output.WriteLine ""
        output.WriteLine("■ {0}", text)
        output.WriteLine(String('-', 68))

    /// 报告里的一行：`  标签   内容`
    let line (output: ITestOutputHelper) (label: string) (text: string) =
        output.WriteLine("  {0}{1}{2}", label, pad 8 label, text)

    /// 二级行：`  标签   子标签   内容`
    let line2 (output: ITestOutputHelper) (label: string) (sub: string) (text: string) =
        output.WriteLine("  {0}{1}{2}{3}{4}", label, pad 8 label, sub, pad 10 sub, text)

    /// 一句话结论（放最后一行）
    let conclusion (output: ITestOutputHelper) (text: string) = line output "结论" text

/// 忙重试：目标程序正忙 / 还没就绪时不要立刻放弃。
/// RPC_E_CALL_REJECTED、RPC_E_SERVERCALL_RETRYLATER，以及 AutoCAD 刚起来时爱报的 RPC_E_SERVERFAULT。
module internal Busy =

    let private busyHResults = [ 0x80010001u; 0x8001010Au; 0x80010105u ]

    /// 连接级别的事故（服务器进程崩了 / 连接断了 / 还没就绪）：换个实例重来一次还有救
    let private deathHResults = [ 0x800706BEu; 0x80010108u; 0x80010105u ]

    let private anyOf (codes: uint32 list) (ex: exn) : bool =
        let mutable e = ex
        let mutable hit = false

        while not (isNull e) && not hit do
            if List.contains (uint32 e.HResult) codes then
                hit <- true

            e <- e.InnerException

        hit

    /// 忙就重试，超时把最后一次的真实异常抛出去
    let retry (timeoutMs: int) (action: unit -> 'T) : 'T =
        let clock = Stopwatch.StartNew()
        let mutable attempt = 0
        let mutable failure: exn = null
        let mutable result = Unchecked.defaultof<'T>
        let mutable finished = false

        while not finished do
            try
                result <- action ()
                finished <- true
            with ex ->
                if anyOf busyHResults ex && clock.ElapsedMilliseconds < int64 timeoutMs then
                    attempt <- attempt + 1
                    Thread.Sleep(min 1000 (200 * attempt))
                else
                    failure <- ex
                    finished <- true

        if not (isNull failure) then
            raise failure

        result

    /// 这条异常是不是「连接断了 / 服务器没了」—— 那种可以整个重来一次
    let isDeath (ex: exn) : bool = anyOf deathHResults ex

/// 认人：动手之前先确认「这就是我要改的那一份」。
/// 判据跟地址指纹一个道理 —— **文件名 + 最后实质保存时间**。
module internal Doc =

    /// 一份文档的身份
    type Identity =
        { Name: string
          /// 文档落在磁盘上的目录；空 = 还没落盘（就是这次新建的临时文档）
          Path: string
          /// 磁盘上那份的最后写入时刻；没落盘为 None
          SavedAt: DateTime option }

    /// 取 COM 属性，取不到就给默认值（新建文档有些属性会抛）
    let private safeGet (get: unit -> 'T) (fallback: 'T) : 'T =
        try
            let value = get ()

            if isNull (box value) then fallback else value
        with _ ->
            fallback

    /// 取身份。`path` 为空 = 还没落到磁盘上。
    let identify (fullName: string) (path: string) : Identity =
        let savedAt =
            if String.IsNullOrEmpty path || not (File.Exists fullName) then
                None
            else
                Some(File.GetLastWriteTime fullName)

        { Name = (if String.IsNullOrEmpty fullName then "（还没有名字）" else fullName)
          Path = path
          SavedAt = savedAt }

    /// 报告里那句话
    let describe (id: Identity) =
        match id.SavedAt with
        | None -> sprintf "「%s」还没落盘（最后保存 —）" id.Name
        | Some t -> sprintf "「%s」最后保存 %s" id.Name (t.ToString "yyyy-MM-dd HH:mm:ss")

    /// 从 COM 对象上取身份（Word 的 doc / Excel 的 book / AutoCAD 的 doc 都是这么问）
    let ofCom (fullName: unit -> string) (path: unit -> string) : Identity =
        identify (safeGet fullName "") (safeGet path "")

    /// **只允许往「没落盘的新文档」里写**。落到磁盘上的一律停手 ——
    /// 那是用户手上的文件，写进去就是把别人的东西弄脏了。
    let ensureFresh (what: string) (id: Identity) =
        match id.SavedAt with
        | Some _ ->
            failwithf "停手：%s %s 是磁盘上的文件，不是这次新建的临时文档，一个字都不写" what (describe id)
        | None -> ()

/// AutoCAD 通道的闸门：没有实例在跑时**默认不新起**。
/// AutoCAD 没有隐藏启动，一弹就是一个真窗口，而绝大多数任务根本不碰 CAD。
/// 真要测这条通道，跑之前设 `XUNITHOST_LAUNCH_ACAD=1`。
module internal Acad =

    let mayLaunch () =
        match Environment.GetEnvironmentVariable "XUNITHOST_LAUNCH_ACAD" with
        | null -> false
        | value -> value = "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase)
