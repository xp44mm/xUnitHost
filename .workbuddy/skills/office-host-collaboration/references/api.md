# API 速查

## 项目结构

| 项目 | 目标框架 | 职责 |
|---|---|---|
| `OsHost` | netstandard2.0 | 操作系统层：STA 线程、进程账本、窗口 → 进程号、时间/文件时间、COM 底层、异常归因；**通用进程地址 `Address.fs`（`ProcAddress` + `Address` 规则）与起 / 关两条通用命令 `ProcHost.fs`**（任意进程，2026-11-06 从 ExcelCli 上收）；四家 host 共用 |
| `WordHost` / `ExcelHost` / `AutoCADHost` | netstandard2.0 | 各管各的：取实例（只取现有 / 取或建）、认自己是谁、收尾、一次性 exec |
| `PowerPointHost` | netstandard2.0 | 2026-11-06 18:12 用户明确要求新建的第 6 个项目。`Host.fs`：取/建实例、`CloseAllPresentations` + `QuitWithoutSaving`、`exec` / `execQuery`（方式一；**没有会话层 identity**）；`PptIO.fs`：Open / SaveAs / 幻灯片文字指纹 / 判宏 / 临时文件 / 只读读回；`PptUpgrade.fs`：`PptToPptx`(24) / `PptToPptm`(25) / `PptUpgrade.pptUpgrade`（`convert` 为对齐别名）。互操作程序集 `Interop.PowerPoint16.dll` 由本机 `MSPPT.OLB` 经 TlbImp 生成，放在 `refs/interop16/` |
| `AcadMgdHost` | netstandard2.0 / x64 | 2026-11-06 18:48 用户明确要求新建的第 7 个项目。**给 accoreconsole.exe / Core Console 用 NETLOAD 加载的 F# 插件**，与 `AutoCADHost` 的 COM 路线完全无关。只引 `acdbmgd.dll` + `accoremgd.dll`（`Private=false`），**绝不引 `acmgd.dll` 等 UI DLL**；`Commands.fs`（`Probe` / `Entry : IExtensionApplication` / `Commands`：`XH_HELLO`、`XH_DB`）+ `AssemblyInfo.fs`（`ExtensionApplication` / `CommandClass` 程序集属性）+ `test/run.cmd`（`/i` + 绝对路径 `/s` 脚本）。详见下面「AcadMgdHost」一节 |
| `xUnitHost` | net10.0 Exe | 测试项目，也是测试入口。**前部**公用件：`Common.fs`（报告/认人/忙重试）、`Session.fs`（会话层）、`ConvertCore.fs`（Item/Outcome/Paths）、`Convert.fs`（分派/批量/核对）、`ProcRx.fs`（Rx 版两条命令的事件流与组合编排）；四条转换住在各自 host 库末尾，两侧对称（`WordHost/DocIO.fs` + `DocUpgrade.fs` 含 DocToDocx / DocToDocm / DocUpgrade；`ExcelHost/BookIO.fs` + `XlsUpgrade.fs` 含 XlsToXlsx / XlsToXlsm / XlsUpgrade）；后面是各 Fact，包括「第二个进程」自己（`SecondProcessTest.fs`）、同步派发（`CommandTest.fs`）、Rx 派发（`RxCommandTest.fs`） |
| `xUnitHost/CommandTest.fs` | net10.0 Exe 内的一条 Fact | 两条命令的同步派发验证（`-m "*派发*"`）：`module Dispatch` 注入 Excel 策略（COM 启动 + `UserControl=true`；ROT 认人 + 四项指纹 + Quit + 放 RCW），也被 Rx 版 Fact 复用；Fact ① 派发 `ProcHost.startWith` 起常驻 Excel 并打印 `Launched`，② 拿①的地址派发 `ProcHost.closeWith` 并打印 `Closed`，末尾 `Proc.waitGone` 复核现场 |
| `xUnitHost/ProcRx.fs` + `RxCommandTest.fs` | 同上，`System.Reactive` 6.0.2 | Rx 版（`-m "*Rx 两条命令*"`）：`ProcessEvent`（Started / Closed / Exited / Refused）+ `ProcRx.start` / `close` / `flow`；`Observable.Create` 包同步命令，`Process.Exited` 事件流 + `Interval` 轮询 `Merge().Take(1)`，`SelectMany` + `Timer(300ms)` 把 Started 接上 close，`Timeout(30s)` 防挂死；退出检测在 close 一侧且**命令执行前**订阅，保证 `Started → Closed → Exited → OnCompleted` 顺序。旧 `CliTest.fs`（依赖已删的 `ExcelCli.exe`）已于 2026-11-06 18:07 经用户确认删除 |

现在是**七个项目**：操作系统层 → 四家 COM host + 一家 accoreconsole 插件 → 会话层（测试项目前部）→ `xUnitHost`。
独立命令行 `ExcelCli` 曾获批准新建，2026-11-06 16:09 用户又把它整个删掉；通用能力留在 `OsHost`。
同日 18:12 用户明确要求新建 `PowerPointHost`；18:48 又要求新建 `AcadMgdHost`，
`xUnitHost.slnx` 现为七个项目。
`PowerPointHost` 只挂在 `xUnitHost`（测试项目引它 + `Interop.PowerPoint16.dll`），不参与 `Convert.all`
批量分派：`.ppt` 走 `Host.execQuery` 里直接调 `pptUpgrade`。
往现有项目**加文件随时可以**（新 Fact、新公用模块），但**不许私自新建项目**：要开新的先问用户。
公用的一律往上收（`OsHost` 或 `xUnitHost/Common.fs`）—— `HostSession`、`DocConvert`、`CrossProcessWorker` 这三个小项目都是这么并掉的。

「另一个进程」不另编程序：跨进程那条 Fact 用 `Report.selfExe()` 把自己再起一份、
`-method "*第二个进程*"` 只跑那一条，地址和话走环境变量，回话写文件。

x64 平台（AutoCAD 互操作只有 64 位）。互操作程序集在 `refs/interop16/` 和
`C:\Program Files\Autodesk\AutoCAD 2020\`，各项目用 `Reference + HintPath` 直接引，不走 NuGet。
`System.Reactive` 6.0.2 **只加在 `xUnitHost`**（Rx 版两条命令用）；`OsHost` 与三家 host 库保持零新依赖。

## AcadMgdHost（accoreconsole 插件）

与 `AutoCADHost` 是**两条不相干的技术路线**：没有 COM / ROT / HWND / 会话地址，
程序集直接跑在 `accoreconsole.exe` 进程里，由 `NETLOAD` 加载。

| 东西 | 位置 / 用法 |
|---|---|
| 引用 | `C:\Program Files\Autodesk\AutoCAD 2020\acdbmgd.dll` + `accoremgd.dll`，`Private=false`（Copy Local = False，AutoCAD 自带，复制会出现两份）；**绝不引 `acmgd.dll` / `AcCui` / `AdUIMgd` 等 UI DLL**。不引 `OsHost`，不带 COM 依赖 |
| `module internal Probe` | `Commands.fs`：`line` / `banner`（程序集、TFM、宿主进程、CLR）/ `database`（`HostApplicationServices.WorkingDatabase` + Transaction 读模型空间实体数）。`line` 同时 `Console.WriteLine` 并按 `ACADMGDHOST_LOG` 追加日志文件 |
| `Entry : IExtensionApplication` | NETLOAD 即调 `Initialize`（无条件报一行 `loaded`）；`ACADMGDHOST_AUTORUN=1` 时直接跑 banner + database，不依赖命令注册 |
| `Commands` | `[<CommandMethod("XH_HELLO")>]` / `[<CommandMethod("XH_DB")>]` 两个自定义命令 |
| `AssemblyInfo.fs` | 编译顺序最后：`[<assembly: ExtensionApplication(typeof<Entry>)>]` + `[<assembly: CommandClass(typeof<Commands>)>]` |
| `test/run.cmd` | 可重跑记录：种子 DWG 复制到 `%TEMP%`（避安装目录只读警告）→ 生成绝对路径 .scr → `/i` + `/s` 调 accoreconsole → 末尾 `type` 插件日志 |

`.scr` 固定形状：`(setvar "secureload" 0)` / `(setvar "filedia" 0)` →
`(command "_.netload" "<绝对路径，正斜杠>")` → `XH_HELLO` / `XH_DB` → `_.quit _y`。
`/s` 只认绝对路径；`/i` 必给；`_.quit _y` 在控制台 = 放弃修改（完整版 acad.exe 里相反）。

**2026-11-06 本机实测（AutoCAD 2020 R23.1 / Q.47.0.0，CLR `4.0.30319.42000`）**：
netstandard2.0 的 F# 程序集被成功加载（报 `.NETStandard,Version=v2.0`）；
**`CommandMethod` 确实注册**（`XH_HELLO` / `XH_DB` 各打出一组 banner/database），
batchplot-plugin 记的「控制台里命令不注册」不是通用规则（它引了 `acmgd.dll` / `AcCui` 等 UI DLL）；
AcDbMgd 数据库 API 在无头进程可用（种子图 3 个实体）。
**新坑**：`Console.Out` 被 accoreconsole 重定向到 `%TEMP%\accc*.tmp`，外层 stdout 看不到
（且该文件删不掉，errno=13），所以插件另按 `ACADMGDHOST_LOG` 写一份可读日志。
`run.cmd` 是纯 ASCII / CRLF / **无 BOM**（cmd.exe 会吃 BOM），其余源文件 UTF-8 BOM + CRLF。

## XUnitHost.Host（会话层，源文件 `xUnitHost/Session.fs`）

```fsharp
// 首次进入（对外签约）：本机有就复用，没有就起一个
val openSession : Kind -> Session

// 执行一步，不关 host；返回 (结果, 重拍过指纹的新 Session)
val exec         : Session -> (App -> 'T) -> Result<'T * Session, string>
val execWord     : Session -> (Word.Application -> 'T) -> Result<'T * Session, string>
val execExcel    : Session -> (Excel.Application -> 'T) -> Result<'T * Session, string>
val execAcad     : Session -> (AcadApplication -> 'T) -> Result<'T * Session, string>

// 凭地址回来（跨进程也行），再执行一步
val execByAddress: string -> (App -> 'T) -> Result<'T * Session, string>

// 只要会话
val resolve      : string -> Result<Session, string>            // 严格核对指纹
val resolveWith  : Drift -> string -> Result<Session, string>   // Drift.Lenient 只认 pid+启动时间
val findByAddress: string -> Session option                     // 只查本进程地址簿，不碰 COM

// 收尾
val close         : Session -> unit
val closeByAddress: string -> Result<unit, string>              // 严格失败自动退宽松

// 摘要（含锚定文档、最后保存时间、有无未保存改动）
val describe : Session -> string

// 方式一：一条命令
val execOnce      : Kind -> (App -> unit) -> Result<unit, string>
val execOnceQuery : Kind -> (App -> 'T) -> Result<'T, string>
```

## 类型

```fsharp
type Kind = Word | Excel | AutoCAD
type App  = WordApp of Word.Application | ExcelApp of Excel.Application | AcadApp of AcadApplication

type Session =
    { Address: string        // 地址（下一步凭它回来）
      Kind: Kind
      App: App
      Owned: bool            // 本次起的 → 收尾可关；false → 一根手指都别动
      Pids: Set<int>         // 收尾只碰这些进程
      Stamp: Stamp           // 拍这张快照时的现场
      Drift: string list }   // 接回来时发现的漂移（空 = 严丝合缝）

// Stamp 就是 OsHost.Identity，三家 host 各自拍一张（WordHost.Host.identity 之类）
type Stamp = OsHost.Identity = { Pid: int; StartedMs: int64; DocPath: string option; SavedMs: int64; Dirty: bool }

type Fingerprint = { Pid: int; StartedMs: int64; SavedMs: int64; PathHash: string }

[<RequireQualifiedAccess>]
type Drift = Strict | Lenient
```

## 地址与指纹

```
<种类>:<pid>:<进程启动毫秒>:<最后保存毫秒>:<路径指纹>:<归属>
word:9232:1793863068197:1793863070426:790cb57f:w
```

| 字段 | 取值 | 核什么 |
|---|---|---|
| 种类 | `word` / `excel` / `cad`（也认 `autocad`） | 走哪套 host |
| pid | 进程号；`0` = 判不出（核对时跳过） | 是不是原来那只进程 |
| 进程启动毫秒 | `Process.StartTime` 转 Unix ms | **pid 复用** |
| 最后保存毫秒 | 锚定文档的 `FileInfo.LastWriteTimeUtc`；`0` = 还没存过盘 | **最后一次实质性保存** |
| 路径指纹 | 绝对路径小写后 SHA256 前 4 字节（8 位 hex）；`-` = 无路径 | 文档换了没有 |
| 归属 | `w` = 本次起的（可关）；`u` = 你原来开着的（不许动） | 收尾怎么处理 |

老地址仍能解析，但会在 `Drift` 里注明没核过：
`word:7572`（只有种类 + pid）、`excel:existing`（连 pid 都没有）。

**核对不过就拒**，错误信息末尾带「这只 host 现在的地址是 …」。宽松模式只要求 pid + 启动时间一致，
漂移项写进 `Session.Drift`。`Owned` 还需要 pid > 0 —— 没有 pid 就证明不了是自己起的。

## 操作系统层 `OsHost`

```fsharp
OsHost.Sta.run   : (unit -> 'T) -> 'T   // 长活：固定一条后台 STA 线程排队干（会话层用）
OsHost.Sta.here  : (unit -> 'T) -> 'T   // 短活：已是 STA 就地跑，否则临时起一条（各 host 的 exec）

OsHost.Proc.ids       : string -> Set<int>            // "WINWORD" / "EXCEL" / "acad"
OsHost.Proc.startedMs : int -> int64 option           // pid 复用检测的锚
OsHost.Proc.waitGone  : string -> Set<int> -> int -> Set<int>
OsHost.Proc.kill      : Set<int> -> unit
OsHost.Proc.shutdown  : string -> int -> Set<int> -> (unit -> unit) -> string option

OsHost.Win32.root       : nativeint -> nativeint      // GA_ROOT：子窗口 → 顶层窗口
OsHost.Win32.pidOfWindow: nativeint -> int option

OsHost.Clock.unixMs : DateTime -> int64
OsHost.Clock.show   : int64 -> string                 // 0 → 「还没保存过」

OsHost.Paths.canonical : string -> string option      // 文档路径 → 绝对路径；新文档 None
OsHost.Paths.savedMs   : string option -> int64       // 磁盘上的 LastWriteTimeUtc（Unix ms）

OsHost.Com.running : string list -> obj option        // ROT 里取现成实例，绝不新建
OsHost.Com.create  : string list -> obj               // CoCreateInstance 起一个

OsHost.Errors.explain : exn -> string                 // 剥到最内层的真实原因
```

### 通用进程地址 + 起 / 关两条命令（`OsHost/Address.fs`、`OsHost/ProcHost.fs`）

```fsharp
// 地址格式的强类型：六段，与 Session.fs 的 word:/excel:/cad: 地址同形，任意进程通用
type OsHost.ProcAddress =
    { Kind: string          // 进程种类：excel / word / cad / 任意进程名
      Pid: int
      StartedMs: int64
      SavedMs: int64        // 普通进程 = 0
      PathHash: string      // 普通进程 = "-"
      Owned: bool }         // w = 本工具起的（close 可关）；u = 用户原来开着的

OsHost.Address.pathHash   : string option -> string        // 绝对路径小写 → SHA256 前 4 字节 → 8 位 hex
OsHost.Address.format     : ProcAddress -> string          // <种类>:<pid>:<启动ms>:<保存ms>:<指纹>:<w|u>
OsHost.Address.ofIdentity : string -> bool -> Identity -> ProcAddress
OsHost.Address.parse      : string -> Result<ProcAddress, string>    // 只认六段；种类不限
OsHost.Address.verify     : ProcAddress -> Identity -> string list   // 空 = 严丝合缝

type OsHost.Launched = { Address: ProcAddress; Pid: int; StartedMs: int64; Others: Set<int> }
type OsHost.Closed   = { Address: string; Pid: int; StartedMs: int64; RecognizedPid: int
                         Quit: bool; Gone: bool; Forced: bool }

OsHost.ProcHost.identityOf : int -> Identity               // 普通进程的现场：SavedMs=0 / DocPath=None
OsHost.ProcHost.start      : string -> string -> (unit -> unit) -> Result<Launched,string>
OsHost.ProcHost.startWith  : string -> string -> (unit -> unit) -> (int -> Identity) -> Result<Launched,string>
                                                    // 自定义现场（Excel 用 ExcelHost.Host.identity）
OsHost.ProcHost.close      : string -> Result<Closed,string>                  // 普通进程：闸门核完按 pid 结束
OsHost.ProcHost.closeWith  : string -> (ProcAddress -> Result<bool * int, string>) -> Result<Closed,string>
                                                    // graceful = 认人 + 优雅退出，Ok(quitOk, 认到的 pid)
```

`startWith` 用「起前起后扫差集」只认这次新起的进程；`closeWith` 的闸门依次是：
六段地址 → 归属必须 `w` → pid > 0 → 进程还在 → 启动时间对得上（防 pid 复用）→ 调用方 graceful
→ 等 5 秒 → 赖着才按地址里的 pid `Proc.kill`。Excel 的 ROT 认人 / 四项指纹 / Quit / 放 RCW
都在 `xUnitHost/CommandTest.fs` 的 `Dispatch.close` graceful 里，OsHost 不认任何 Office 类型。

## Rx 版两条命令（源文件 `xUnitHost/ProcRx.fs`，只加在测试项目）

```fsharp
type ProcessEvent =
    | Started of OsHost.Launched
    | Closed  of OsHost.Closed
    | Exited  of pid:int * exitCode:int option * via:string   // via = "Process.Exited" / "轮询兜底"
    | Refused of reason:string

ProcRx.start : (unit -> Result<Launched,string>) -> IObservable<ProcessEvent>
ProcRx.close : ProcAddress -> (string -> Result<Closed,string>) -> int -> IObservable<ProcessEvent>
ProcRx.flow  : TimeSpan -> TimeSpan -> IObservable<ProcessEvent>
               -> (ProcAddress -> IObservable<ProcessEvent>) -> IObservable<ProcessEvent>
```

- `start` 送完 `Started` 就 Complete；`close` **在派发命令前**订阅退出检测
  （`Process.Exited` 事件 `Observable.Create` + `Interval` 轮询 `Merge().Take(1)`），
  `Closed` 先发、退出事件后补；`flow` 用 `SelectMany` + `Timer(settle)` 把 Started 接上 close，
  外层 `Timeout(30s)` 是「两个事件之间的最长间隔」，卡住就 OnError。
- 给事件留优先窗口靠 poll 间隔（Fact 里 1000ms）：200ms 时轮询会抢先，报告 `via` 如实显示来源。
  退出检测放 start 一侧是错的 —— 那样 Exited 会抢在 Closed 前面到。
- `Disposable.Create` 退订时只摘 handler + `Process.Dispose()`，**绝不 kill 进程**；
  真正结束进程只由 `ProcHost.closeWith` 按地址里的 pid 做。
- 示范代码里的 `Process.Start "excel.exe" /e`（会复用用户实例）和 `CloseMainWindow()` 遍历
  所有 excel 进程（会关掉用户手上的表）在本项目都不采用。

## 转换（六条实现在 host 库；批量入口在 `xUnitHost/Convert.fs`）

六条转换各自独立，住在对应 host 库编译列表的**末尾**；格式码、目标后缀、核对口径、文案
都只写在自己模块里。三家形状对称：各一份文件放三个入口（两条转换 + 一个统一入口）。
`Word` / `Excel` 接进 `Convert.all` 做批量；`.ppt` 目前只作独立函数用，没接进 `Paths.scan` / `Convert.all`
（避免真目录扫描和删原件的行为被悄悄扩大）。
host 库只回传 `(目标, 状态, 详情)`，不认识测试项目的 `Outcome`——
依赖方向不倒挂（host 不能引用测试项目）。

```fsharp
// xUnitHost/ConvertCore.fs：纯数据 + 路径，无 Office 调用
type Item / type Outcome
Paths.scan                     : string -> Item list        // 递归找 .doc/.xls（跳过点目录、~$ 锁文件）
Paths.targetOf                 : string -> string           // 还没判宏，一律先按无宏占位：.docx / .xlsx
Paths.existingTargetOf         : string -> string           // 删原件时用：.docm / .xlsm 在就是它，否则 .docx / .xlsx

// WordHost/DocIO.fs（WordHost 内，两条转换共用）
WordHost.DocIO.openDoc / closeDoc / saveDocAs / normalize / bodyText / hasVba / stageFor / cleanupStage
WordHost.DocIO.readText        : Application -> string -> string
WordHost.DocIO.detectMacro     : Application -> string -> Result<bool, string>  // 读不出 VBA 信息就 Error，原件保留

// WordHost/DocUpgrade.fs（末尾：三个入口合在一份文件）
WordHost.DocToDocx.format / targetOf / convert            // 12，无宏；带宏拒单
WordHost.DocToDocm.format / targetOf / convert            // 13，正文 + 复核 VBA 宏还在
WordHost.DocUpgrade.Route  = Docx | Docm                  // 纯规则，测试可直接验
WordHost.DocUpgrade.route   : bool -> Route               // 有宏 → Docm，没宏 → Docx
WordHost.DocUpgrade.convert : Application -> string -> string * string * string
                                                        // 统一入口：判宏 → 选路 → 转；判不出就失败、原件保留

// ExcelHost/BookIO.fs（ExcelHost 内，两条转换共用）
ExcelHost.BookIO.openBook / closeBook / bookPrint / hasVba / stageFor / cleanupStage
ExcelHost.BookIO.readPrint     : Application -> string -> string
ExcelHost.BookIO.detectMacro   : Application -> string -> Result<bool, string>

// ExcelHost/XlsUpgrade.fs（末尾：三个入口合在一份文件）
ExcelHost.XlsToXlsx.format / targetOf / convert           // 51，无宏；带宏拒单
ExcelHost.XlsToXlsm.format / targetOf / convert           // 52，格子值 + 复核 VBA 宏还在
ExcelHost.XlsUpgrade.Route   = Xlsx | Xlsm                // 纯规则，测试可直接验
ExcelHost.XlsUpgrade.route   : bool -> Route              // 有宏 → Xlsm，没宏 → Xlsx
ExcelHost.XlsUpgrade.convert : Application -> string -> string * string * string

// PowerPointHost/PptIO.fs（PowerPointHost 内，两条转换共用；参数是强类型按值，不是 obj&）
PowerPointHost.PptIO.openPpt / closePpt / savePptAs / normalize / slidePrint / hasVba / stageFor / cleanupStage
PowerPointHost.PptIO.readPrint   : Application -> string -> string    // 幻灯片文字指纹
PowerPointHost.PptIO.detectMacro : Application -> string -> Result<bool, string>
PowerPointHost.Host.exec / execQuery : (Application -> 'T) -> Result<'T, string>   // 进程名 POWERPNT

// PowerPointHost/PptUpgrade.fs（末尾：三个入口合在一份文件）
PowerPointHost.PptToPptx.format / targetOf / convert    // ppSaveAsOpenXMLPresentation = 24，无宏；带宏拒单
PowerPointHost.PptToPptm.format / targetOf / convert    // ppSaveAsOpenXMLPresentationMacroEnabled = 25，复核 VBA
PowerPointHost.PptUpgrade.Route = Pptx | Pptm           // 纯规则，测试可直接验
PowerPointHost.PptUpgrade.route : bool -> Route         // 有宏 → Pptm，没宏 → Pptx
PowerPointHost.PptUpgrade.pptUpgrade : Application -> string -> string * string * string
                                                      // 判宏 → 选路 → 转；判不出就失败、原件保留
PowerPointHost.PptUpgrade.convert = pptUpgrade          // 与另两个 Upgrade.convert 对齐的别名

// xUnitHost/Convert.fs：组装 Outcome、批量、事后核对（只分派 .doc / .xls）
Convert.all                    : Item list -> Outcome list   // .doc 走一个 Word 进程，.xls 走一个 Excel 进程
Convert.words / Convert.books  : Item list -> Outcome list   // 分别直接调 DocUpgrade / XlsUpgrade.convert
Convert.readWordText           : string -> Result<string, string>
Convert.readBookPrint          : string -> Result<string, string>   // 每张表的表名 + 格子值
Convert.canDelete              : Outcome -> bool             // 「转了」/「已有·一致」才允许删原件
```

真目录的两条 Fact（`ConvertTest.fs` 末段）靠环境变量叫起来，不设就跳过：

```bash
XUNITHOST_CONVERT_DIR=<目录> [XUNITHOST_CONVERT_BATCH=10]   xUnitHost.exe -showLiveOutput -method "*真目录*清点*"   # 只盘点
  xUnitHost.exe -showLiveOutput -method "*真目录*转*"     # 转 + 核对 + 删通过的原件
```

## 三个 host 库

```fsharp
// 对外签约：有就复用，没有就新建（Excel 会补空白工作簿，AutoCAD 会起分钟级的进程）
WordHost.Host.GetOrCreateApplication   : unit -> Application
ExcelHost.Host.GetOrCreateApplication  : unit -> Application
AutoCADHost.Host.GetOrCreateApplication: unit -> AcadApplication

// 同事打招呼：只接 ROT 里现成的，没有就是 None，绝不新建
WordHost.Host.TryGetRunningApplication   : unit -> Application option
ExcelHost.Host.TryGetRunningApplication  : unit -> Application option
AutoCADHost.Host.TryGetRunningApplication: unit -> AcadApplication option

// 认自己是谁：进程号 + 锚定文档 + 最后保存 + 有无未保存改动（地址指纹的来源）
WordHost.Host.identity     : Application -> OsHost.Identity
ExcelHost.Host.identity    : Application -> OsHost.Identity
AutoCADHost.Host.identity  : AcadApplication -> OsHost.Identity
WordHost.Host.pidOf        : Application -> int option          // 从文档窗口往上找 OpusApp
ExcelHost.Host.pidOf       : Application -> int option          // app.Hwnd
AutoCADHost.Host.pidOf     : AcadApplication -> int option      // app.HWND

// 收尾：丢改动 + Quit，不弹保存框
WordHost.Host.QuitWithoutSaving     : Application -> unit
ExcelHost.Host.QuitWithoutSaving    : Application -> unit
ExcelHost.Host.CloseAllWorkbooksWithoutSaving : Application -> int
AutoCADHost.Host.QuitWithoutSaving  : AcadApplication -> unit
AutoCADHost.Host.CloseActiveDocumentWithoutSaving : AcadApplication -> bool
AutoCADHost.Host.CloseAllDocumentsWithoutSaving   : AcadApplication -> int

// 方式一：起一个 → 干活 → 关掉（AutoCAD 要求本机没有实例在跑）
WordHost.Host.exec      : (Application -> unit) -> Result<unit, string>
WordHost.Host.execQuery : (Application -> 'T) -> Result<'T, string>
// Excel / AutoCAD 同名
```

## 取文档对象的常用写法

```fsharp
// Word：新文档 / 当前文档
let doc = app.Documents.Add()
let doc = app.ActiveDocument
doc.Content.Text <- "…"
doc.Content.Text.Trim [| '\r'; '\a'; '\n'; ' ' |]   // 末尾总有 \r\a

// Word 保存：参数是 ref object，必须 mutable 变量取地址（FS0421）
let mutable fileName: obj = box path
doc.SaveAs(&fileName) |> ignore

// Excel
let book  = app.Workbooks.Add()
let sheet = book.Worksheets.[1] :?> Excel.Worksheet
sheet.Range("A1").Value2 <- "…"

// AutoCAD
let doc = app.ActiveDocument
```

## 编译 / 运行

```bash
dotnet build xUnitHost.slnx -v q
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*关键字*" -noColor -showLiveOutput
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -noColor -showLiveOutput
```

`-list tests` 看所有 Fact；`-id <id>` 按 ID 跑单个用例；`-class "XUnitHost.CollaborationTest"` 按类跑
（Fact 改成成员方法后，`-method` 要带类名匹配）。

## 黑板模式：四家 GetApplication(info)（2026-11-07）

场景：WordVSTO / ExcelVSTO 插件把当前实例抄到黑板上（`[…] 采集时间 …` / PID / ROT / HWND / EXE / 文档 / 路径），
消费方只按 info 精确挂回那一只实例读写；**黑板的创建与销毁不归它管**，不并行，交接时数据要完整、不过时、没被别的任务改过。

- `OsHost/AppInfo.fs`：`AppKind = Word | Excel | PowerPoint | AutoCAD`；**`AppInfo` 是调用方要填的结构化参数表**：
  `Kind/Pid/Rot/Hwnd/ExePath/CapturedMs/DocName/DocPath`，全是结构化字段，没有一行展示文本。
  直接构造记录，或用类型上的**静态构造成员** `AppInfo.create (kind, pid, rot, hwnd, exePath, capturedMs)`
  起最小必填集再 `with` 补 `DocName` / `DocPath`（2026-11-07 11:37 取代原模块函数 `ofParts`）。
  **成员函数**（2026-11-07 11:44）：`AppKind` 上 `kind.image()` = 进程名、`kind.kindName()` = 给人看的名字；
  `AppInfo` 上 `info.describe()` = 一行摘要。同名 `module AppInfo` 里只剩
  `verify` / `ensureKind` / `attach`（+ 私有 `firstError`），它们内部改调这些成员。
  **库不解析文本**：`ofText` 与展示用的 `Title` 已从 `OsHost` 移除 —— 插件改标题、键名、措辞只该改调用方。
  采集时间由填报方给 Unix 毫秒，库不替你猜。
  `verify` 核「进程还在 / 启动时间不晚于采集 / HWND 仍属该 pid / EXE 一致 / ROT 名后缀 pid 一致」；
  `ensureKind`；`attach` = 种类 → 现场校验 → 按名取对象 → 交 work。
  **文档级交接检查推迟**（2026-11-07 13:51 用户定）：「换了文档 / 文档关了 / 采集后存过盘」这一层两路都不做，
  `checkDoc` 已从 `AppInfo.fs` 与 `GeneralAppInfo.fs` 撤掉；`DocName` / `DocPath` 字段保留，数据照收、暂不参与判断。
  旧实现可从 git 历史（`6be1725` 及之前）找回。
- `OsHost.Com` 新增：`runningAt : string -> obj option`（`CreateItemMoniker("!", 名)` + `IRunningObjectTable.GetObject`，
  **不枚举 ROT、不退回 GetActiveObject、绝不新建**）、`register` / `revoke`（自测登记本地对象用）、
  `OsHost.Proc.exePath : int -> string option`。
- 四家入口形状对称：`WordHost.Host.GetApplication : AppInfo -> Result<Application,string>`；
  `ExcelHost.Host.GetApplication` / `PowerPointHost.Host.GetApplication` /
  `AutoCADHost.Host.GetApplication : AppInfo -> Result<AcadApplication,string>`。
  种类不符 / 信息过时 / ROT 里没这个名字 / pid 对不上，一律 `Error`；pid 对上就直接交回 Application
  （13:51 起不再核当前文档，见上面的推迟说明）。
  绝不新建、绝不 Quit、不改 `DisplayAlerts`。`PowerPointHost.Host.pidOf` 用 `app.HWND`。
- **公寓**：`attach` 走 `Sta.run` 固定长活 STA 线程；后续对 Application 的读写也包 `Sta.run`。
  一步做完就直接 `AppInfo.attach info kind progIds work`，别让 RCW 跑出临时 STA 线程。
- **参考解析（不是库 API）**：`xUnitHost/BlackboardText.fs`，`module internal BlackboardText`：
  `caption`（只贴标题行给人看）/ `ofText : string -> Result<AppInfo,string>` —— 把插件复制的展示文本
  填成同一张参数表。格式强耦合的代码只留在这份里；用不上文本源的程序直接按字段填 `AppInfo`。
- 验证：`xUnitHost/BlackboardTest.fs` 四条 Fact（`-method "*黑板*"`）；第一条同时验
  「手工填报的参数表 == 参考解析出来的那条」；真机那条设
  `XUNITHOST_BLACKBOARD_FILE=<剪贴板文本存的文件>`，不设就 Skip。原「文档级校验」那条 Fact 已随推迟决定撤掉。

### 另一条路：GeneralAppInfo（2026-11-07 12:01，用户要求「原文件不动，另起炉灶」）

12:01 落地时 `OsHost/AppInfo.fs`、四家 `GetApplication(info)` 确未动（提交 `6be1725`）；
13:51 用户要求「文档级交接检查推迟」，两路库与四家入口才一起改，除此之外两条路仍并存。
新文件 `OsHost/GeneralAppInfo.fs`：

- `GeneralIdentity`：**参数占位**，取代 `AppKind` 上那两个 match —— `Image`（进程名）/ `Name`（显示名）/
  `Tag`（调用方自己的种类标记）；`static member create (image, name, tag)`。库只比字符串、不解释 Tag，
  四家之外的第五家不用改库。
- `GeneralAppInfo`：纯参数表 `Identity/Pid/Rot/Hwnd/ExePath/CapturedMs/DocName/DocPath`，
  `static member create (identity, pid, rot, hwnd, exePath, capturedMs)` / `info.describe()`。
- `module GeneralAppInfo`：`verify`（与 `AppInfo.verify` 同口径，另多一条「Image 为空」）、
  `ensureTag`（比 Tag，不比枚举）、`defaultAcquire`（ROT 名优先，空则按 ProgID 列表）、
  （文档级 `checkDoc` 已按 13:51 的决定撤掉，字段保留不判断）
  **`attachWith info tag acquire work`** —— `acquire : GeneralAppInfo -> obj option` 与
  `work : obj -> Result<'T,string>` 是**获取函数占位**，由调用方注入；`attach` 是套默认 `defaultAcquire` 的便捷版。
  全程 `Sta.run` 固定 STA 线程；取不到就 `Error`，绝不新建 / Quit / 退回 GetActiveObject。
- 生产层**不知道 `AppKind` 存在**；`AppInfo → GeneralAppInfo` 的桥接示例 `module internal GeneralBridge`
  留在测试项目 `xUnitHost/GeneralAppInfoTest.fs`（不是库 API）。
- 验证：三条 Fact，`-method "*通用参数表*"`。其中「获取函数占位」那条用 `Com.register` 在当前进程登记
  `BlackboardProbe`，`register / acquire / revoke` 都排在**同一条** `Sta.run` 线程上
  （避免跨公寓、也避免 `Sta.run` 嵌套），正向挂回 `Ok "XUnitHost.BlackboardProbe"` —— 不需要 Office 的完整附着链。
  第三条现在只对照 `verify` 与文档字段照抄，不再验「换了文档 / 文档关了 / 存过盘」四种情形。
