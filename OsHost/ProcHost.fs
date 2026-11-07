/// 两条命令的**通用实现**：起一只常驻进程、按地址把它关掉。
///
/// 「起 / 关」跟进程是什么没关系：起前起后扫差集认出这次新起的，签一张带指纹的地址；
/// 回来时按地址核种类 / 归属 / pid / 启动时间，过了闸门才动手。任意进程都能用，
/// Office 只是其中一种 —— Excel 特有的「补工作簿 / ROT 认人 / Quit / 放 RCW」由调用方
/// 以策略回调注入，本层不认任何 Office 类型。
namespace OsHost

open System
open System.Diagnostics
open System.Threading

/// 起一只常驻进程的结果
type Launched =
    { /// 强类型地址（要一行文本交给脚本，用 Address.format）
      Address: ProcAddress
      Pid: int
      StartedMs: int64
      /// 起之前就在跑的同名进程：非空时按地址回来要格外小心
      Others: Set<int> }

/// 按地址关掉一只常驻进程的结果
type Closed =
    { Address: string
      Pid: int
      StartedMs: int64
      /// 策略认回来的进程号；0 = 没走认人（普通进程直接按 pid 结束）
      RecognizedPid: int
      /// 优雅退出跑成功了吗（没给策略 = false）
      Quit: bool
      /// 进程真的消失了吗
      Gone: bool
      /// 有没有动用 Proc.kill
      Forced: bool }

[<RequireQualifiedAccess>]
module ProcHost =

    /// 普通进程没有文档：现场只有 pid 和启动时间，地址里自然就是 0 / -
    let identityOf (pid: int) : Identity =
        { Pid = pid
          StartedMs = Proc.startedMs pid |> Option.defaultValue 0L
          DocPath = None
          SavedMs = 0L
          Dirty = false }

    /// 新进程可能晚几十毫秒才出现在进程表里，扫几轮再下结论
    let private waitNew (image: string) (before: Set<int>) : Set<int> =
        let mutable ours = Set.empty<int>
        let mutable tries = 10

        while ours.IsEmpty && tries > 0 do
            ours <- Set.difference (Proc.ids image) before

            if ours.IsEmpty then
                Thread.Sleep 200
                tries <- tries - 1

        ours

    // ---------------- 第一条命令：起一只常驻进程 ----------------

    /// 通用 start：boot 真正起进程（COM / Process.Start 都行），identify 拍现场。
    /// 只认「这次新起的」差集；没起新进程就绝不认领用户手上那只。
    let startWith
        (image: string)
        (kind: string)
        (boot: unit -> unit)
        (identify: int -> Identity)
        : Result<Launched, string> =
        let before = Proc.ids image

        let booted =
            try
                boot ()
                Ok()
            with ex ->
                Error(Errors.explain ex)

        match booted with
        | Error message -> Error message
        | Ok() ->
            let ours = waitNew image before

            if ours.IsEmpty then
                Error(
                    sprintf
                        "没有新起 %s 进程（可能复用了本机已有的实例）；不动你手上的"
                        image
                )
            else
                // 差集里不止一只时判不准是哪只，宁可让地址里 pid = 0（close 会拒接）
                let ownPid = if ours.Count = 1 then Set.minElement ours else 0

                let id =
                    try
                        identify ownPid
                    with _ ->
                        identityOf ownPid

                Ok
                    { Address = Address.ofIdentity kind true id
                      Pid = id.Pid
                      StartedMs = id.StartedMs
                      Others = before }

    /// 普通进程版：现场自动从进程表拍，不用调用方给 identify
    let start (image: string) (kind: string) (boot: unit -> unit) : Result<Launched, string> =
        startWith image kind boot identityOf

    // ---------------- 第二条命令：按地址关掉它 ----------------

    let private alive (pid: int) : bool =
        try
            (Process.GetProcessById pid).Id = pid
        with _ ->
            false

    /// 等这个进程号真的消失（按 pid 问，不依赖进程名 —— 任意进程都适用）
    let private waitPid (pid: int) (timeoutMs: int) : bool =
        let clock = Stopwatch.StartNew()
        let mutable here = true

        while here && clock.ElapsedMilliseconds < int64 timeoutMs do
            here <- alive pid

            if here then
                Thread.Sleep 100

        not here

    /// 通用 close 的闸门：六段地址 → 归属必须 w → pid > 0 → 进程还在 → 启动时间对得上。
    /// 过了闸门才让调用方的 graceful 接手（认人 + 优雅退出），最后统一等 / 兜底结束。
    ///
    /// graceful 返回 `Ok (是否优雅退出成功, 认到的 pid)`；普通进程直接给 (false, 0)，
    /// 于是省掉等待、按地址里的 pid 结束。返回 Error 就是拒接，一根手指都不动。
    let closeWith
        (text: string)
        (graceful: ProcAddress -> Result<bool * int, string>)
        : Result<Closed, string> =
        match Address.parse text with
        | Error message -> Error message
        | Ok a when not a.Owned ->
            Error "地址里的归属是 u（你原来就开着的实例）：close 一根手指都不碰"
        | Ok a when a.Pid <= 0 ->
            Error "地址里没有进程号：证明不了是哪只进程，close 不动手"
        | Ok a ->
            if not (alive a.Pid) then
                Error(sprintf "地址已失效：%s pid %d 已经不在了" a.Kind a.Pid)
            else
                match Proc.startedMs a.Pid with
                | Some ms when a.StartedMs > 0L && ms <> a.StartedMs ->
                    Error(
                        sprintf
                            "地址已失效：pid %d 的启动时间对不上，这个号现在被别的程序占着"
                            a.Pid
                    )
                | _ ->
                    match graceful a with
                    | Error message -> Error message
                    | Ok(quitOk, recognized) ->
                        // 优雅退出过才值得等；普通进程 / Quit 抛异常就直接来硬的
                        let waited = if quitOk then waitPid a.Pid 5000 else false
                        let forced = not waited

                        if forced then
                            Proc.kill (Set.singleton a.Pid)

                        let gone = waitPid a.Pid 3000

                        Ok
                            { Address = text
                              Pid = a.Pid
                              StartedMs = a.StartedMs
                              RecognizedPid = recognized
                              Quit = quitOk
                              Gone = gone
                              Forced = forced }

    /// 普通进程版：不认人、不优雅退出，闸门核完直接按 pid 结束
    let close (text: string) : Result<Closed, string> =
        closeWith text (fun _ -> Ok(false, 0))
