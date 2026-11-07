/// 「常驻进程」的地址格式 —— 规则只留这一处（操作系统层）。
///
/// 起一只进程、隔一会儿凭地址认回同一只进程，跟它是 Excel 还是记事本没关系：
///   · 种类          excel / word / cad，或者任意进程名
///   · pid + 启动毫秒   认「还是原来那只进程」，pid 被系统复用立刻识破
///   · 保存毫秒 + 路径指纹   认「还是原来那份文档」；普通进程没有文档，写 0 和 "-"
///   · 归属 w/u     w = 本工具起的（close 可关）；u = 用户原来开着的（一根手指都不许动）
///
/// 六段形状与 xUnitHost/Session.fs 完全一致：会话层签的 word:/excel:/cad: 地址这里解析得了，
/// 这里签的 excel: 地址 Host.resolve 也认得。**地址在代码里是值（ProcAddress）**，
/// 只有跨进程 / 打印时才由 format 变成一行文本。
namespace OsHost

open System
open System.Security.Cryptography
open System.Text

/// 地址格式的强类型表示（「怎么访问一只常驻进程」的完整依据）
type ProcAddress =
    { /// 进程种类（小写）：excel / word / cad / 任意进程名
      Kind: string
      /// 进程号；0 = 判不出来（close 不认 0）
      Pid: int
      /// 进程启动时间（Unix 毫秒）：pid 复用检测的锚
      StartedMs: int64
      /// 锚定文档最后一次落盘的毫秒；0 = 还没存过盘 / 普通进程没有文档
      SavedMs: int64
      /// 文档路径的 8 位短哈希；"-" = 没落盘 / 普通进程
      PathHash: string
      /// w = 本工具起的（close 可关）；u = 用户原来开着的（一根手指都不许动）
      Owned: bool }

/// 编址 / 解析 / 核对。任意进程共用这一套，不再每写一个命令行就抄一遍。
[<RequireQualifiedAccess>]
module Address =

    /// 与会话层同一套：绝对路径小写 → SHA256 前 4 字节 → 8 位 hex
    let pathHash (path: string option) : string =
        match path with
        | Some p when not (String.IsNullOrWhiteSpace p) ->
            use sha = SHA256.Create()
            let digest = sha.ComputeHash(Encoding.UTF8.GetBytes(p.ToLowerInvariant()))
            digest.[0..3] |> Array.map(fun b -> b.ToString "x2") |> String.concat ""
        | _ -> "-"

    /// 序列化成一行文本：<种类>:<pid>:<启动ms>:<保存ms>:<路径指纹>:<归属>
    let format (a: ProcAddress) : string =
        sprintf
            "%s:%d:%d:%d:%s:%s"
            a.Kind
            a.Pid
            a.StartedMs
            a.SavedMs
            a.PathHash
            (if a.Owned then "w" else "u")

    /// 用现场（三家 host 各自拍的 Identity，或普通进程的 identityOf）签一张地址
    let ofIdentity (kind: string) (owned: bool) (id: Identity) : ProcAddress =
        { Kind = kind.ToLowerInvariant()
          Pid = id.Pid
          StartedMs = id.StartedMs
          SavedMs = id.SavedMs
          PathHash = pathHash id.DocPath
          Owned = owned }

    let private bad (text: string) (reason: string) =
        Error(sprintf "地址不对：%s（%s）" text reason)

    /// 只认六段新地址。种类不限 —— 任意进程都签得出、解析得回。
    /// 老格式（excel:7572 / excel:existing）不带指纹，close 是破坏性操作，一律不收。
    let parse (text: string) : Result<ProcAddress, string> =
        if String.IsNullOrWhiteSpace text then
            bad text "空的"
        else
            let parts = text.Trim().Split ':'

            if parts.Length <> 6 then
                bad text "应该是 <种类>:<pid>:<启动ms>:<保存ms>:<路径指纹>:<归属> 六段"
            else
                let kind = parts.[0].Trim().ToLowerInvariant()

                if String.IsNullOrWhiteSpace kind then
                    bad text "种类段是空的"
                else
                    match
                        Int32.TryParse(parts.[1].Trim()),
                        Int64.TryParse(parts.[2].Trim()),
                        Int64.TryParse(parts.[3].Trim())
                    with
                    | (true, pid), (true, started), (true, saved) when pid > 0 ->
                        let owner = parts.[5].Trim().ToLowerInvariant()

                        if owner <> "w" && owner <> "u" then
                            bad text "归属只能是 w 或 u"
                        else
                            Ok
                                { Kind = kind
                                  Pid = pid
                                  StartedMs = started
                                  SavedMs = saved
                                  PathHash = parts.[4].Trim()
                                  Owned = (owner = "w") }
                    | _ -> bad text "pid / 启动毫秒 / 保存毫秒 得是数字，且 pid > 0"

    /// 现场核对：pid、启动时间、最后保存、锚定文档。返回漂移说明（空 = 严丝合缝）。
    /// 口径与 Session.verify 一致：判不出来的那一项跳过，其余对不上就报出来。
    let verify (expected: ProcAddress) (cur: Identity) : string list =
        [ if expected.Pid > 0 && cur.Pid > 0 && expected.Pid <> cur.Pid then
              yield sprintf "进程号对不上：地址里 %d，现场 %d" expected.Pid cur.Pid

          if expected.StartedMs > 0L && cur.StartedMs > 0L && expected.StartedMs <> cur.StartedMs then
              yield sprintf "进程启动时间对不上（pid %d 可能已被复用）" expected.Pid

          if expected.SavedMs <> cur.SavedMs then
              yield
                  sprintf
                      "最后一次保存变了：地址里 %s，现在 %s"
                      (Clock.show expected.SavedMs)
                      (Clock.show cur.SavedMs)

          let now = pathHash cur.DocPath

          if expected.PathHash <> now then
              yield sprintf "锚定的文档换了：地址里 %s，现在 %s" expected.PathHash now ]
