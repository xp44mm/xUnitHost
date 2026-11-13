/// 方式二：合作完成任务。
///
/// 和「一条命令开-做-关」相对：这里 host 开一次、连着用几步，中间每步都把
/// **地址（Id）+ 结果**交回给大模型，大模型据此写下一步；最后大模型说「完事了」，
/// 再按 **第一步之前的状态** 把 host 收掉：
///   · 第一步之前没有实例 → 是我们起的 → 关掉（图纸/工作簿丢改动、Quit、清残进程）
///   · 第一步之前就在跑 → 是你手上的 → 一个字都不动，原样留着
///
/// 地址不只是「进程号」—— 光有进程号接回来并不保险：进程可能退了、pid 可能被系统
/// 复用给别的程序、中间那份文档可能被换掉或者被别人另存过一次。所以地址里带一整套
/// **指纹**，回来时逐项核对：
///
///     word:7572:1762340880123:1762340987654:3f5a91c2:w
///          └pid┘└─进程启动毫秒─┘└─最后保存毫秒─┘└路径指纹┘└归属┘
///
///   · pid + 进程启动时间 —— 认「还是原来那只进程」，pid 被复用立刻识破
///   · 最后保存毫秒      —— 文档文件的最后写入时间，即**最后一次实质性保存**的时刻；
///                          中间被别人存过一次，地址就对不上了（这正是要的保险）
///   · 路径指纹          —— 文档路径的短哈希，认「还是同一份文档」，换文档立刻识破
///   · 归属 w/u          —— 本次起的（收尾可关）/ 你原来就开着的（收尾一个字不动）
///
/// 对不上默认**拒绝**（`Host.resolve`），并告诉你这只 host 现在的地址；
/// 只想确认「还是那个进程」、愿意接受漂移的，用 `Host.resolveWith Drift.Lenient`。
///
/// 还有一条同样重要：**进 host 的两条路不能混**。
///   · `openSession` = 首次进入（对外签约）：本机没有就真的起一个，记「谁起的」，收尾按原状态关；
///   · `resolve`     = 按地址回来（harness 内部交接）：只接现成的那只，**绝不新建** ——
///     地址说的人不在就报错。顺手变一个出来，等于把别人的会话顶掉，还认错了人。
///
/// 这个文件放在测试项目的**前部**：它是 harness 的协调层，只服务「测试怎么跟 host 打配合」。
/// 三家 host 各自的脾气在各自的库里，操作系统那点事（线程/进程/窗口/时间）在 `OsHost` 里。
/// 「第二个进程」不是另编的程序，就是测试程序自己再起一份（`SecondProcessTest.fs`），
/// 拿的也是这一份规则 —— 地址怎么编、指纹怎么核，只有一处。
namespace XUnitHost

open System
open System.Collections.Concurrent
open System.IO
open System.Runtime.InteropServices
open System.Security.Cryptography
open System.Text
open OsHost

/// 三类 host
type Kind =
    | Word
    | Excel

/// 拿到手的应用对象（三种之一，模式匹配取用）
type App =
    | WordApp of Microsoft.Office.Interop.Word.Application
    | ExcelApp of Microsoft.Office.Interop.Excel.Application

/// 现场快照：这只 host 此刻是什么样（地址里的指纹就是从它来的）。
/// 三家 host 各自拍自己的（`WordHost.Host.identity` 之类），形状统一在 `OsHost.Identity`。
type Stamp = Identity

/// 地址里携带的那部分指纹（跨进程时，这就是核对「是不是那一只」的全部依据）
type Fingerprint =
    { Pid: int
      StartedMs: int64
      SavedMs: int64
      /// 文档路径的短哈希（8 位十六进制）；"-" = 还没存过盘
      PathHash: string }

/// 核对不上时的态度
[<RequireQualifiedAccess>]
type Drift =
    /// 默认：保存时间/文档对不上就拒绝（要的就是这份保险）
    | Strict
    /// 只要求「还是原来那只进程」，保存时间/文档的漂移记在 Session.Drift 里放行
    | Lenient

/// 一个「开着的 host」：地址 + 应用对象 + 是不是这次自己起的 + 现场指纹
type Session =
    { /// 地址：下一步凭它回来找这个 host，例：`word:7572:1762340880123:1762340987654:3f5a91c2:w`
      Address: string
      Kind: Kind
      App: App
      /// 本次之前就有实例在跑（你手上的）→ 收尾时不许退它
      Owned: bool
      /// 自己起的那些进程号（收尾只碰这些）
      Pids: Set<int>
      /// 拍这张快照时的现场（执行一步后会重新拍，见 exec 返回的新 Session）
      Stamp: Stamp
      /// 接回来时发现的漂移（空 = 严丝合缝）
      Drift: string list }

/// 地址簿：`openSession` 发的地址都记在这儿，同进程里下一步凭地址就能马上找回来。
/// 跨进程由 `Host.resolve` 负责（把地址里的指纹映射回 COM 里那只实例）。
module private Registry =

    let private table = ConcurrentDictionary<string, Session>()

    let add (session: Session) : unit = table.[session.Address] <- session

    let remove (address: string) : unit =
        match table.TryRemove(address) with
        | true, _ -> ()
        | _ -> ()

    let find (address: string) : Session option =
        match table.TryGetValue(address) with
        | true, session -> Some session
        | _ -> None

    /// 这只应用对象是不是「我们手上已经有一只」的那个
    let tracked (app: App) : bool =
        let same (a: App) (b: App) =
            match a, b with
            | WordApp x, WordApp y -> Object.ReferenceEquals(x, y)
            | ExcelApp x, ExcelApp y -> Object.ReferenceEquals(x, y)
            | _ -> false

        table.Values |> Seq.exists(fun session -> same session.App app)

[<RequireQualifiedAccess>]
module Host =

    // ---------------- 小工具 ----------------

    let private imageName (kind: Kind) : string =
        match kind with
        | Kind.Word -> "WINWORD"
        | Kind.Excel -> "EXCEL"

    let private kindName (kind: Kind) : string =
        match kind with
        | Kind.Word -> "word"
        | Kind.Excel -> "excel"

    // ---------------- 指纹：拍快照 / 编地址 / 核对 ----------------

    /// 路径的短指纹：地址里放不下整条路径，放 8 位十六进制；没路径就是 "-"
    let private pathHash (path: string option) : string =
        match path with
        | Some p when not (String.IsNullOrWhiteSpace p) ->
            use sha = SHA256.Create()
            let digest = sha.ComputeHash(Encoding.UTF8.GetBytes(p.ToLowerInvariant()))
            digest.[0..3] |> Array.map(fun b -> b.ToString "x2") |> String.concat ""
        | _ -> "-"

    /// 现场拍一张快照 —— 指纹全从这儿来。各家自己知道「活动文档在哪儿、谁在写它」，
    /// 所以先问 host 要身份，再补一件只有会话层关心的事：进程号判不出来时怎么办。
    ///
    /// 进程号有三种来路，依次退：
    ///   1) 从应用主窗口反查（最准；但 Word 一份文档都没有时压根没有文档窗口）；
    ///   2) 调用方本来就知道的那个 pid（自己起的进程、上次快照里的 pid）；
    ///   3) 机器上这种进程只有一个 —— 那就是它。
    /// 三条都不成立就记 0（「判不出」），地址里如实写 0，核对时跳过 pid 这一项。
    let private snapshot (kind: Kind) (fallbackPid: int) (app: App) : Stamp =
        let seen =
            match app with
            | WordApp a -> WordHost.Host.identity a
            | ExcelApp a -> ExcelHost.Host.identity a

        let pid =
            if seen.Pid > 0 then
                seen.Pid
            elif fallbackPid > 0 then
                fallbackPid
            else
                match Process.ids (imageName kind) |> Set.toList with
                | [ only ] -> only
                | _ -> 0

        if pid = seen.Pid then
            seen
        else
            { seen with
                Pid = pid
                StartedMs = Process.startedMs pid |> Option.defaultValue 0L }

    /// 编地址：种类:pid:启动毫秒:最后保存毫秒:路径指纹:归属
    let private addressOf (kind: Kind) (owned: bool) (stamp: Stamp) : string =
        sprintf
            "%s:%d:%d:%d:%s:%s"
            (kindName kind)
            stamp.Pid
            stamp.StartedMs
            stamp.SavedMs
            (pathHash stamp.DocPath)
            (if owned then "w" else "u")

    /// 核对「凭地址找回来的这只」还是不是当初那只。返回漂移说明（空 = 完全对上）。
    let private verify
        (given: string)
        (kind: Kind)
        (owned: bool)
        (expected: Fingerprint)
        (cur: Stamp)
        (drift: Drift)
        : Result<string list, string> =
        let current = addressOf kind owned cur

        if expected.Pid > 0 && cur.Pid > 0 && expected.Pid <> cur.Pid then
            Error(
                sprintf
                    "接不上 %s：COM 交出来的是 pid=%d 那只，不是你要的 pid=%d（同时开了多个实例，ROT 只暴露一只）"
                    given
                    cur.Pid
                    expected.Pid
            )
        elif expected.StartedMs > 0L && cur.StartedMs > 0L && expected.StartedMs <> cur.StartedMs then
            Error(
                sprintf
                    "地址已失效：%s（进程 %d 的启动时间对不上 —— 原来那只早退了，这个 pid 现在被别的程序占着）"
                    given
                    expected.Pid
            )
        else
            let notes =
                [ if expected.SavedMs <> cur.SavedMs then
                      yield
                          sprintf
                              "最后一次保存变了：地址里是 %s，现在是 %s"
                              (Clock.show expected.SavedMs)
                              (Clock.show cur.SavedMs)

                  if expected.PathHash <> pathHash cur.DocPath then
                      yield
                          sprintf
                              "锚定的文档换了：地址里是 %s，现在是 %s"
                              (if expected.PathHash = "-" then "还没存过的文档" else expected.PathHash)
                              (match cur.DocPath with
                               | Some p -> Path.GetFileName p
                               | None -> "还没存过的文档") ]

            if notes.IsEmpty then
                Ok []
            else
                match drift with
                | Drift.Strict ->
                    Error(
                        sprintf
                            "地址对不上：%s —— %s。这只 host 现在的地址是 %s（拿新地址接着做）"
                            given
                            (String.Join("；", List.toArray notes))
                            current
                    )
                | Drift.Lenient -> Ok notes

    // ---------------- 取应用 ----------------
    // 进 host 有两条完全不同的路，别混着用：
    //
    //   · 首次进入（openSession）= 对外签约：从你手上接管，本机没有就真的起一个进程，
    //     要记「谁起的」（收尾按原状态关），AutoCAD 还要等它冷启动就绪 —— 用 getOrCreateOnSta（就在 Sta.run 里，别再排一次 STA）。
    //   · 按地址回来（resolve）  = 同事之间打个招呼：只接 ROT 里现成的那只，**绝不新建**。
    //     地址说的人不在就报错 —— 顺手变一个出来，等于把别人的会话顶掉，还认错了人。

    /// 首次进入：本机有就复用，没有就起一个
    let private openApp (kind: Kind) : App =
        match kind with
        | Kind.Word -> App.WordApp(WordHost.Host.GetOrCreateApplication())
        | Kind.Excel -> App.ExcelApp(ExcelHost.ExcelTask.getOrCreateOnSta())

    /// 按地址回来：只接现成的那只，没有就是 None（绝不新建）
    let private tryAttachApp (kind: Kind) : App option =
        match kind with
        | Kind.Word -> WordHost.Host.TryGetRunningApplication() |> Option.map App.WordApp
        | Kind.Excel -> ExcelHost.Host.TryGetRunningApplication() |> Option.map App.ExcelApp

    let word (session: Session) : Microsoft.Office.Interop.Word.Application =
        match session.App with
        | WordApp app -> app
        | _ -> invalidArg "session" "这个会话不是 Word"

    let excel (session: Session) : Microsoft.Office.Interop.Excel.Application =
        match session.App with
        | ExcelApp app -> app
        | _ -> invalidArg "session" "这个会话不是 Excel"


    // ---------------- 会话：建 / 重拍 / 开 / 用 / 收 ----------------

    /// 用一张快照建会话并记进地址簿
    let private attach
        (kind: Kind)
        (app: App)
        (owned: bool)
        (pids: Set<int>)
        (stamp: Stamp)
        (notes: string list)
        : Session =
        let session =
            { Address = addressOf kind owned stamp
              Kind = kind
              App = app
              Owned = owned
              Pids = pids
              Stamp = stamp
              Drift = notes }

        Registry.add session
        session

    /// 执行完一步后重新拍快照：文档被保存/换掉，地址就跟着变，
    /// 老地址自动作废（下一次拿老地址回来会被核对挡下来）
    let private resealWith (notes: string list) (session: Session) : Session =
        let stamp = snapshot session.Kind session.Stamp.Pid session.App
        let address = addressOf session.Kind session.Owned stamp

        let next =
            { session with
                Address = address
                Stamp = stamp
                Drift = notes }

        if address <> session.Address then
            // 老地址作废：还留着它，就等于留了一个「过了期还能用」的后门
            Registry.remove session.Address

        Registry.add next
        next

    let private reseal (session: Session) : Session = resealWith [] session

    /// 开一个 host：本机已有就复用（不算我们的），没有就起一个（算我们的）。
    /// 返回的 Address 就是「地址」，下一步拿它回来。
    let openSession (kind: Kind) : Session =
        STA.run (fun () ->
            let image = imageName kind
            let before = Process.ids image
            let app = openApp kind

            // 起之前没有、起之后有了 —— 这些进程是这次自己起的
            let pids = Set.difference (Process.ids image) before
            let owned = not pids.IsEmpty

            // 自己起的时候进程号是现成的；复用你那只就靠主窗口反查
            attach kind app owned pids (snapshot kind (if owned then Set.minElement pids else 0) app) [])

    /// 在 host 上执行一步，**不关 host**。
    /// 除了结果，还会把**重拍过指纹的新 Session**交回去：这一步要是保存了文档，
    /// 新地址就是唯一还有效的那个。
    let exec (session: Session) (action: App -> 'T) : Result<'T * Session, string> =
        STA.run (fun () ->
            try
                let value = action session.App
                Ok(value, reseal session)
            with ex ->
                Error(Errors.explain ex))

    /// 在 host 上执行一步（按种类给强类型 app，省得每步都模式匹配）
    let execWord
        (session: Session)
        (action: Microsoft.Office.Interop.Word.Application -> 'T)
        : Result<'T * Session, string> =
        exec session (fun app ->
            match app with
            | WordApp a -> action a
            | _ -> invalidArg "session" "这个会话不是 Word")

    let execExcel
        (session: Session)
        (action: Microsoft.Office.Interop.Excel.Application -> 'T)
        : Result<'T * Session, string> =
        exec session (fun app ->
            match app with
            | ExcelApp a -> action a
            | _ -> invalidArg "session" "这个会话不是 Excel")


    /// 凭地址找回来（只查本进程的地址簿：快、无副作用。要跨进程用 resolve）
    let findByAddress (address: string) : Session option = Registry.find address

    // ---------------- 跨进程：地址里的指纹 → COM 里那只实例 ----------------

    type private Spec =
        { Kind: Kind
          /// 地址里的进程号；0 = 地址里没有（existing）或当初判不出来
          Pid: int
          /// 地址里带的指纹；旧格式（word:7572 / excel:existing）没有
          Stamp: Fingerprint option
          Owned: bool }

    /// `word:7572:1762...:1762...:3f5a91c2:w` → 拆成可核对的东西。
    /// 也认老地址 `word:7572`（只核进程）和 `excel:existing`。
    let private parseAddress (address: string) : Result<Spec, string> =
        let bad () =
            Error(sprintf "地址格式不对：%s（应该是 word:7572:启动毫秒:保存毫秒:路径指纹:w 这种）" address)

        if String.IsNullOrWhiteSpace address then
            bad ()
        else
            let parts = address.Split ':'

            let kind =
                match parts.[0].Trim().ToLowerInvariant() with
                | "word" -> Some Kind.Word
                | "excel" -> Some Kind.Excel
                | _ -> None

            match kind with
            | None -> bad ()
            | Some _ when parts.Length < 2 -> bad ()
            | Some kind ->
                let tail = parts.[1].Trim()

                if tail.Equals("existing", StringComparison.OrdinalIgnoreCase) then
                    Ok
                        { Kind = kind
                          Pid = 0
                          Stamp = None
                          Owned = false }
                else
                    match Int32.TryParse tail with
                    | false, _ -> bad ()
                    | true, pid when pid < 0 -> bad ()
                    | true, pid when parts.Length = 2 ->
                        // 老格式：只有种类 + 进程号，没有保存时间的依据
                        Ok
                            { Kind = kind
                              Pid = pid
                              Stamp = None
                              Owned = true }
                    | true, pid ->
                        let number (index: int) =
                            if parts.Length > index then
                                match Int64.TryParse(parts.[index].Trim()) with
                                | true, value -> value
                                | _ -> 0L
                            else
                                0L

                        let owner =
                            if parts.Length > 5 then parts.[5].Trim().ToLowerInvariant() else "u"

                        Ok
                            { Kind = kind
                              Pid = pid
                              Stamp =
                                Some
                                    { Pid = pid
                                      StartedMs = number 2
                                      SavedMs = number 3
                                      PathHash =
                                        if parts.Length > 4 && parts.[4].Trim() <> "" then
                                            parts.[4].Trim()
                                        else
                                            "-" }
                              // 归属位缺失时按「你手上的」算：宁可少关，不可误关
                              Owned = (owner = "w") }

    /// 跨进程解析出来的会话：只有「自己起的」才把那个进程号记成待收尾的进程，
    /// 用户手上那只（u）收尾时一根手指头都不许碰
    let private ownedPids (spec: Spec) : Set<int> =
        if spec.Owned && spec.Pid > 0 then Set.singleton spec.Pid else Set.empty

    /// 放掉一只应用对象（解析失败 / 不是我们要的那只时别攥着）
    let private releaseApp (app: App) : unit =
        let target: obj =
            match app with
            | WordApp a -> box a
            | ExcelApp a -> box a

        try
            if not (isNull target) then
                Marshal.FinalReleaseComObject target |> ignore
        with _ ->
            ()

    /// 凭地址找回 host —— **跨进程也能用**，而且**核指纹**。
    ///
    /// 1) 同进程：地址簿里直接有，但仍然现场重拍一次快照核对 ——
    ///    文档在这期间被另存过（哪怕是本进程上一步干的），地址簿里的保存时间就已经过期了；
    /// 2) 别的进程：地址里的进程号先确认还活着，再让 COM 把那只 Application 交出来（ROT 里登记的），
    ///    然后逐项核对 pid / 进程启动时间 / 最后保存时间 / 文档路径指纹。
    ///
    /// 对不上就老实说，绝不乱动（也不会顺手把别人的文档当成你的）。
    let resolveWith (drift: Drift) (address: string) : Result<Session, string> =
        match parseAddress address with
        | Error message -> Error message
        | Ok spec ->
            match Registry.find address with
            | Some cached ->
                STA.run (fun () ->
                    match spec.Stamp with
                    | None ->
                        // 老格式地址：当初就没存指纹，只能原样信地址簿
                        Ok cached
                    | Some expected ->
                        let cur = snapshot cached.Kind cached.Stamp.Pid cached.App

                        match verify address cached.Kind cached.Owned expected cur drift with
                        | Error message -> Error message
                        | Ok notes -> Ok(resealWith notes { cached with Stamp = cur }))
            | None ->
                let image = imageName spec.Kind
                let alive = Process.ids image

                if spec.Pid > 0 && not (alive.Contains spec.Pid) then
                    let now =
                        if alive.IsEmpty then
                            "现在一个都没有"
                        else
                            "现在有 " + String.Join(", ", alive)

                    Error(
                        sprintf "地址已失效：%s（%s 进程 %d 已经不在了，%s）" address image spec.Pid now
                    )
                else
                    STA.run (fun () ->
                        match tryAttachApp spec.Kind with
                        | None ->
                            Error(
                                sprintf
                                    "接不上 %s：现在没有可接的 %s 实例。按地址回来只认现成的那只，绝不新建 —— 该重新 openSession"
                                    address
                                    image
                            )
                        | Some app ->
                            let cur = snapshot spec.Kind spec.Pid app

                            // 归属：地址说「本次起的」还不够 —— 没有 pid 就证明不了是自己起的，
                            // 一律按「你手上的」处理（收尾一个字不动）
                            let owned = spec.Owned && spec.Pid > 0

                            match spec.Stamp with
                            | None ->
                                let notes =
                                    [ if spec.Pid = 0 then
                                          "地址里没有进程号（existing），没核过是不是同一只"
                                      else
                                          "老格式地址只带进程号，没核最后保存时间" ]

                                Ok(attach spec.Kind app owned (ownedPids spec) cur notes)
                            | Some expected ->
                                match verify address spec.Kind owned expected cur drift with
                                | Error message ->
                                    // RCW 是按 IUnknown 身份缓存的：这只 app 很可能就是地址簿里那只，
                                    // 硬释放会把会话里的引用一并拆掉（再摸就是 InvalidComObjectException）。
                                    // 只放真正没见过的那只。
                                    if not (Registry.tracked app) then
                                        releaseApp app

                                    Error message
                                | Ok notes ->
                                    Ok(attach spec.Kind app owned (ownedPids spec) cur notes))

    /// 凭地址找回 host（默认严格核对指纹）
    let resolve (address: string) : Result<Session, string> = resolveWith Drift.Strict address

    /// 凭地址回到那台 host 上执行一步 —— 大模型那句「下一步拿地址回去接着做」。
    /// 和 exec 一样：做完不关 host，并把重拍过指纹的新 Session 交回去。
    let execByAddress (address: string) (action: App -> 'T) : Result<'T * Session, string> =
        match resolve address with
        | Ok session -> exec session action
        | Error message -> Error message

    /// 收尾：按**第一步之前的状态**把 host 关回去。
    /// 自己起的 → 关掉文档/工作簿（改动丢弃）+ Quit + 清掉残进程；
    /// 之前就有的 → 原样留着，一个字都不动。
    let close (session: Session) : unit =
        Registry.remove session.Address

        if session.Owned then
            STA.run (fun () ->
                (match session.App with
                 | WordApp app -> WordHost.Host.QuitWithoutSaving app
                 | ExcelApp app -> ExcelHost.Host.QuitWithoutSaving app
                 )

                // 松开 RCW：客户端还捏着引用时，Excel 这类服务器 Quit 了也不肯退
                let target: obj =
                    match session.App with
                    | WordApp app -> box app
                    | ExcelApp app -> box app

                (try
                    Marshal.FinalReleaseComObject target |> ignore
                 with _ ->
                     ())

                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()

                let image = imageName session.Kind
                // Excel 在「客户端还活着」时经常不肯退：等一小会儿，还不走就结束进程
                let stubborn = Process.waitGone image session.Pids 5000
                Process.kill stubborn
                // 结束进程也是异步的：等它真的消失，别让人刚 close 完又看见它
                Process.waitGone image session.Pids 3000 |> ignore)

    /// 凭地址收尾（跨进程也行）：先按严格指纹解析，再按原状态关。
    /// 收尾只关心「是不是同一只进程」，所以保存时间漂移了也照关（丢弃改动 + 退出），
    /// 但 pid 被复用 / 换了一只进程的情况仍然拒绝。
    let closeByAddress (address: string) : Result<unit, string> =
        let resolved =
            match resolve address with
            | Ok session -> Ok session
            | Error _ -> resolveWith Drift.Lenient address

        match resolved with
        | Error message -> Error message
        | Ok session ->
            close session
            Ok()

    /// 一行摘要（给大模型 / 日志看）：地址、种类、归属、几份文档、锚定哪份、最后保存什么时候
    let describe (session: Session) : string =
        let documents =
            try
                match session.App with
                | WordApp app -> app.Documents.Count
                | ExcelApp app -> app.Workbooks.Count
            with _ ->
                -1

        let doc =
            match session.Stamp.DocPath with
            | Some p -> Path.GetFileName p
            | None -> "还没存过盘的新文档"

        sprintf
            "%s（%s，%s，当前 %d 份，锚定 %s，最后保存 %s%s）"
            session.Address
            (kindName session.Kind)
            (if session.Owned then "本次起的" else "你原来就开着的")
            documents
            doc
            (Clock.show session.Stamp.SavedMs)
            (if session.Stamp.Dirty then "，有未保存改动" else "")

    // ---------------- 方式一：一条命令 ----------------
    // 直接转发给各 host 库的 exec / execQuery：开一次、做一段、关掉，不留 host。

    /// 一条命令：起 host → 执行 → 关掉。成功 Ok ()，失败 Ok 里拿不到东西就给 Error 原因。
    let execOnce (kind: Kind) (action: App -> unit) : Result<unit, string> =
        match kind with
        | Kind.Word -> WordHost.Host.exec (fun app -> action (WordApp app))
        | Kind.Excel -> ExcelHost.Host.exec (fun app -> action (ExcelApp app))

    /// 一条命令 + 取回一个值。
    let execOnceQuery (kind: Kind) (query: App -> 'T) : Result<'T, string> =
        match kind with
        | Kind.Word -> WordHost.Host.execQuery (fun app -> query (WordApp app))
        | Kind.Excel -> ExcelHost.Host.execQuery (fun app -> query (ExcelApp app))
