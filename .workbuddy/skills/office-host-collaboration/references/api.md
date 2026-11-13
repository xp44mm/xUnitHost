# API 速查

## 项目结构

| 项目 | 目标框架 | 职责 |
|---|---|---|
| `OsHost` | netstandard2.0 | 操作系统层：STA 线程、进程账本、窗口 → 进程号、时间/文件时间、COM 底层、异常归因、**现场快照 `Identity`（会话层地址的指纹就是从它来的）**；四家 host 共用 |
| `WordHost` / `ExcelHost` / `AutoCADHost` | netstandard2.0 | 各管各的：取实例（只取现有 / 取或建）、认自己是谁、收尾、一次性 exec |
| `PowerPointHost` | netstandard2.0 | `Host.fs`：取/建实例、`CloseAllPresentations` + `QuitWithoutSaving`、`exec` / `execQuery`（方式一；**没有会话层 identity**）；`PptIO.fs`：Open / SaveAs / 幻灯片文字指纹 / 判宏 / 临时文件 / 只读读回；`PptUpgrade.fs`：`PptToPptx`(24) / `PptToPptm`(25) / `PptUpgrade.pptUpgrade`（`convert` 为对齐别名）。互操作程序集 `Interop.PowerPoint16.dll` 由本机 `MSPPT.OLB` 经 TlbImp 生成，放在 `refs/interop16/` |
| `AcadMgdHost` | netstandard2.0 / x64 | **给 accoreconsole.exe / Core Console 用 NETLOAD 加载的 F# 插件**，与 `AutoCADHost` 的 COM 路线完全无关。只引 `acdbmgd.dll` + `accoremgd.dll`（`Private=false`），**绝不引 `acmgd.dll` 等 UI DLL**；`Probe.fs` + `Star.fs` / `StarGeometry.fs`（五角星）+ `Commands.fs`（`Entry : IExtensionApplication` / `Commands`：`XH_HELLO`、`XH_DB`、`XH_STAR`）+ `AssemblyInfo.fs`（`ExtensionApplication` / `CommandClass` 程序集属性）+ `test/run.cmd`（`/i` + 绝对路径 `/s` 脚本）。详见下面「AcadMgdHost」一节 |
| `xUnitHost` | net10.0 Exe | 测试项目，也是测试入口。**前部**公用件：`Common.fs`（报告/认人/忙重试）、`Session.fs`（会话层）、`ConvertCore.fs`（Item/Outcome/Paths）、`Convert.fs`（分派/批量/核对）；四条转换住在各自 host 库末尾，两侧对称（`WordHost/DocIO.fs` + `DocUpgrade.fs` 含 DocToDocx / DocToDocm / DocUpgrade；`ExcelHost/BookIO.fs` + `XlsUpgrade.fs` 含 XlsToXlsx / XlsToXlsm / XlsUpgrade）；后面是各 Fact，包括「第二个进程」自己（`SecondProcessTest.fs`） |

现在是**七个项目**（`xUnitHost.slnx` 就是这七行）：操作系统层 → 四家 COM host + 一家 accoreconsole 插件 → 会话层（测试项目前部）→ `xUnitHost`。
`PowerPointHost` 只挂在 `xUnitHost`（测试项目引它 + `Interop.PowerPoint16.dll`），不参与 `Convert.all`
批量分派：`.ppt` 走 `Host.execQuery` 里直接调 `pptUpgrade`。
往现有项目**加文件随时可以**（新 Fact、新公用模块），但**不许私自新建项目**：要开新的先问用户。
公用的一律往上收（`OsHost` 或 `xUnitHost/Common.fs`），别再另开小项目。

「另一个进程」不另编程序：跨进程那条 Fact 用 `Report.selfExe()` 把自己再起一份、
`-method "*第二个进程*"` 只跑那一条，地址和话走环境变量，回话写文件。

x64 平台（AutoCAD 互操作只有 64 位）。互操作程序集在 `refs/interop16/` 和
`C:\Program Files\Autodesk\AutoCAD 2020\`，各项目用 `Reference + HintPath` 直接引，不走 NuGet。

## AcadMgdHost（accoreconsole 插件）

与 `AutoCADHost` 是**两条不相干的技术路线**：没有 COM / ROT / HWND / 会话地址，
程序集直接跑在 `accoreconsole.exe` 进程里，由 `NETLOAD` 加载。

| 东西 | 位置 / 用法 |
|---|---|
| 引用 | `C:\Program Files\Autodesk\AutoCAD 2020\acdbmgd.dll` + `accoremgd.dll`，`Private=false`（Copy Local = False，AutoCAD 自带，复制会出现两份）；**绝不引 `acmgd.dll` / `AcCui` / `AdUIMgd` 等 UI DLL**。不引 `OsHost`，不带 COM 依赖 |
| `module internal Probe` | `Probe.fs`：`line` / `banner`（程序集、TFM、宿主进程、CLR）/ `database`（`HostApplicationServices.WorkingDatabase` + Transaction 读模型空间实体数）。`line` 同时 `Console.WriteLine` 并按 `ACADMGDHOST_LOG` 追加日志文件 |
| `module internal Star` + `StarGeometry` | `Star.fs`：在 `WorkingDatabase` 里画五角星 → `SaveAs` 到桌面（`ACADMGDHOST_STAR_OUT` 可覆盖）→ 从磁盘重开逐点读回核对；`StarGeometry.fs` 只有纯几何 `vertices` / `innerRatio`（sin18°/sin54°≈0.381966），**同一份源文件链接进 `xUnitHost` 的 Fact**，画的和断言的是同一个公式 |
| `Entry : IExtensionApplication` | NETLOAD 即调 `Initialize`（无条件报一行 `loaded`）；`ACADMGDHOST_AUTORUN=1` 时直接跑 banner + database，不依赖命令注册 |
| `Commands` | `[<CommandMethod("XH_HELLO")>]` / `XH_DB` / `XH_STAR` 三个自定义命令 |
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
// 没有 execAcad：acad 不进会话层（Kind 只有 Word | Excel）。
// acad 按 PID 挂回 + 发命令一律走 AutoCADHost.AutoCADTask，见下面「AutoCADHost.AutoCADTask」一节。

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
type Kind = Word | Excel          // acad 不在会话层：没有 Kind.AutoCAD
type App  = WordApp of Word.Application | ExcelApp of Excel.Application

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
OsHost.Com.runningEntries : unit -> (string * obj) list   // 枚举 ROT（必须自己核 pid）

// ROT 名的通用切 / 拼（OsHost/Rot.fs）：只认「<前缀><pid>」这个形状，纯字符串、不碰 COM
// 前缀是各家插件定的，固化在各自项目（WordHost.Rot / PowerPointHost.Rot），通用层不认识任何前缀
OsHost.Rot.parse : string -> (string * int) option    // !WordVSTO.Instance.27232 → Some("!WordVSTO.Instance.", 27232)
OsHost.Rot.build : string -> int -> string option     // "!WordVSTO.Instance." 27232 → Some "!WordVSTO.Instance.27232"

OsHost.Errors.explain : exn -> string                 // 剥到最内层的真实原因
```

### 读文本文件（`OsHost/TextFile.fs`）—— 编码和行末都替你想好了

```fsharp
OsHost.TextFile.decodeAllLinesAsync : string -> Task<string * string[]>   // 编码名 * 行数组：对外唯一入口
OsHost.TextFile.decodeAllTextAsync  : string -> Task<string * string>     // 编码名 * 全文
OsHost.TextFile.decodeAllLines      : byte[] -> string * string[]         // 同步底座（已经是字节时用）
OsHost.lines / OsHost.dec           : string -> Task<string * string[]>   // [<AutoOpen>] 短名，open OsHost 直接用
OsHost.TextFile.registerTextShort / tryTextShort / textShortNames          // 短名表（默认已登记 lines / dec）
```

**读任何来路不明的文本一律用它**，别写 `File.ReadAllLines`、别自己 `Encoding.GetEncoding` 试、别手切 `\r\n`：

- 判定顺序：① BOM（UTF-8 / UTF-16LE|BE / UTF-32LE|BE，编码名带 `-bom`，BOM 字节会剥掉，不会粘在第一行行首）→
  ② 全字节 < 0x80 → `ascii` → ③ 无 BOM 时 **UTF-8 与 GBK(936) 各自严格试解码**：
  只有一个过就报那个；都过且文本不同 → `utf-8|gbk`（**标明歧义，不默默选边**，正文按 UTF-8 给）；
  都不过 → 取替换字符少的那个，编码名带 `?` 表示存疑。
- 行：按 `\r\n` / `\n` / `\r` 切（裸 `\r` 也认），**末尾换行不多带一个空行**。
- 与项目约定对得上：写文件是 UTF-8 + BOM + CRLF，读回来第一行不带 BOM 残留、最后不多一个 `""`。

```fsharp
// FSI / 脚本最省事的写法：直接引编译产物，别抄源码
#r @"C:\Application Data\GitHub\xp44mm\xUnitHost\OsHost\bin\Debug\net10.0\OsHost.dll"
open OsHost
let enc, rows = lines @"C:\某个\未知编码.txt" |> Async.AwaitTask |> Async.RunSynchronously
printfn "%s / %d 行 / 首行 %s" enc rows.Length rows.[0]
```

自检：`-method "*文本读取*"`（`xUnitHost/TextFileTest.fs`，离线，不碰任何 Office 进程）。

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
WordHost.DocToDocx.format / targetOf / convert            // 12，→ .docx（宏不保留，2026-11-11 用户定）
WordHost.DocToDocm.format / targetOf / convert            // 13，→ .docm；只在 XUNITHOST_DOC_KEEP_MACRO=1 时走
WordHost.DocUpgrade.Route  = Docx | Docm                  // 纯规则，测试可直接验
WordHost.DocUpgrade.route   : bool -> Route               // 纯规则：有宏 → Docm，没宏 → Docx
WordHost.DocUpgrade.convert : Application -> string -> string * string * string
                                                        // 统一入口：**默认一律 .docx**；KEEP_MACRO=1 才按 route 走 Docm

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

## 三家 host 库（Word / Excel / PowerPoint）

> **`AutoCADHost` 没有 `Host.fs`** —— 它只有 `EventPair.fs` / `AcadBeginEnd.fs` / `Application.fs` /
> `AutoCADTask.fs`，驱动入口全在 `module AutoCADHost.AutoCADTask`（见下一节）。下面这段 `Host.*` 只管 Office 三家。

```fsharp
// 对外签约：有就复用，没有就新建（Excel 会补空白工作簿）
WordHost.Host.GetOrCreateApplication        : unit -> Application
ExcelHost.ExcelTask.getOrCreateOnSta        : unit -> Application   // 必须在固定 STA 线程上（Sta.run 里）

// 同事打招呼：只接 ROT 里现成的，没有就是 None，绝不新建
WordHost.Host.TryGetRunningApplication   : unit -> Application option
ExcelHost.Host.TryGetRunningApplication  : unit -> Application option

// 认自己是谁：进程号 + 锚定文档 + 最后保存 + 有无未保存改动（地址指纹的来源）
WordHost.Host.identity     : Application -> OsHost.Identity
ExcelHost.Host.identity    : Application -> OsHost.Identity
WordHost.Host.pidOf        : Application -> int option          // 从文档窗口往上找 OpusApp
ExcelHost.Host.pidOf       : Application -> int option          // app.Hwnd

// ROT 注册名（前缀固化在各家）：!WordVSTO.Instance.<pid> / !PowerPointVSTO.Instance.<pid>
WordHost.Rot.Prefix        : string                       // "!WordVSTO.Instance."
WordHost.Rot.pidOf         : string -> int option         // 前缀不对就 None（别人的登记项不认）
WordHost.Rot.nameOf        : int -> string option         // pid → "!WordVSTO.Instance.<pid>"
PowerPointHost.Rot.Prefix / pidOf / nameOf                // 同形，"!PowerPointVSTO.Instance."

// 收尾：丢改动 + Quit，不弹保存框
WordHost.Host.QuitWithoutSaving     : Application -> unit
ExcelHost.Host.QuitWithoutSaving    : Application -> unit
ExcelHost.Host.CloseAllWorkbooksWithoutSaving : Application -> int

// 方式一：起一个 → 干活 → 关掉
WordHost.Host.exec      : (Application -> unit) -> Result<unit, string>
WordHost.Host.execQuery : (Application -> 'T) -> Result<'T, string>
// Excel 同名
```

## 成品 Task 入口（`ExcelTask.fs` / `WordTask.fs`）

访问代码优先用这一层：`*Async` 名字、返回 `Task<_>`，`let!` 直接串。三家的形状一致：
**前置闸门同步判**（不合格立刻交回已完成且 faulted 的 `Task`，连 STA 队列都不排）→ 合格才排进
`OsHost.Sta.runAsync` 那条固定 STA 线程 → `Result` 的 `Error` 由 `unwrap` 抛异常置成 faulted，
await 时抛 `InvalidOperationException`，文案与同步版一致。

```fsharp
// 挂回现成的那一只（按 PID 硬核对；绝不新建、绝不 Quit）
ExcelHost.ExcelTask.GetApplicationAsync      : int -> Task<Application>
AutoCADHost.AutoCADTask.GetApplicationAsync  : int -> Task<IAcadApplication>
WordHost.WordTask.GetApplicationAsync        : int -> Task<Application>
PowerPointHost.Host.GetApplication           : int -> Result<Application, string>   // 同步

// 起或取（对外签约：真会起进程；Excel 会补一张空白工作簿，XLL 一起加载）
ExcelHost.ExcelTask.getOrCreateOnSta            : unit -> Application        // 已经在 STA 上时用这条（同步）
ExcelHost.ExcelTask.GetOrCreateApplicationAsync : unit -> Task<Application>  // 不在 STA 时用这条
```

RCW 仍只在固定 STA 线程上用：`await` 之后每一步包 `Sta.run` / `Sta.runAsync`。
验证：`-m "*Excel 挂回*"`（离线三条 PID 闸门 + 真机）、`-m "*不合格*"`（Word / PPT 闸门）。

## AutoCADHost.AutoCADTask（acad 的 COM 驱动入口）

源文件 `AutoCADHost/AutoCADTask.fs`（`module AutoCADHost.AutoCADTask`，编译顺序排在 `Application.fs` 之后）。
**做 acad 的活就用这里的函数**，别自己拼 `SendCommand`、别自己 `add_BeginCommand`。
所有输入类型都是接口：`IAcadApplication` / `IAcadDocument`（不是 coclass `AcadApplication`）。

```fsharp
// 挂回：按 PID 取活着的那只。PID ≤ 0 / 不是 acad → 已完成且 faulted 的 Task，绝不新建
val GetApplicationAsync : int -> Task<IAcadApplication>

// 命令：BeginCommand → EndCommand。结果 = 命令名。commandAsync 用于「启动动作不是发命令行」
val sendCommandAsync : EventPair.Options -> IAcadDocument -> string -> Task<string>
val commandAsync     : EventPair.Options -> IAcadApplication -> (unit -> unit) -> Task<string>

// 打开图纸：BeginOpen → EndOpen，结果 = 图纸全名
val openDrawingAsync : EventPair.Options -> IAcadApplication -> (unit -> unit) -> Task<string>

// 模态对话框：BeginModal → EndModal（对话框关掉 AutoCAD 才真的空下来）
val modalAsync : EventPair.Options -> IAcadApplication -> Task<unit>

// LISP：BeginLisp → EndLisp（Esc 走 LispCancelled）
type LispOutcome = Completed | Cancelled
val runLispAsync          : IAcadDocument -> string -> Task<LispOutcome>                      // 默认策略
val sendLispAndWaitAsync  : EventPair.Options -> IAcadDocument -> string -> Task<LispOutcome>
val runLispFileAsync      : IAcadDocument -> string -> Task<LispOutcome>                      // (load "...")，等

// LISP：只交出去，不等（fire-and-forget）
val sendLispAsync : IAcadDocument -> string -> Task<unit>
val loadLispAsync : IAcadDocument -> string -> Task<unit>

// 事件策略（EventPair.fs）
type EventPair.Options =
    { RequireBegin: bool           // 必须先见 Begin 才认 End，防串台
      Timeout: TimeSpan option     // None = 不设时限（可能永久挂住，不推荐）
      Cancellation: CancellationToken
      Unsubscribe: IDisposable -> unit }   // 退订必须在挂 RCW 的那条 STA 上
val EventPair.defaults : Options   // RequireBegin=true / 30 s / 不取消 / 就地 Dispose
```

判据（选错函数就会 30 秒超时或提前返回）：

| 送进去的东西 | 用它 | 等的钩子 |
|---|---|---|
| 内置命令 `_.LINE` / `_.CIRCLE` / `_.ZOOM` | `sendCommandAsync` | EndCommand（`SendCommand` 返回不算完） |
| 一整段 LISP `(progn …)` / `(setq …)` / `(load "…")` | `runLispAsync` / `runLispFileAsync` | EndLisp（内部 `(command …)` 的 EndCommand 先到，别误判） |
| LISP 但后面没有串行动作 | `sendLispAsync` / `loadLispAsync` | 不等，返回只代表「已交出去」 |
| 打开一张图 | `openDrawingAsync` | EndOpen（比撞忙重试可靠） |
| 有模态框挡着 | `modalAsync` | EndModal |

线程与公寓：这些入口都必须在挂着 RCW 的那条 STA 线程上调用（`start` 也在调用线程同步跑）。
写法固定是 `Sta.runAsync(fun () -> AutoCADTask.xxx …).Unwrap()`；`GetApplicationAsync` 取回的 RCW
后续也要在 `Sta.run` 里用。`AutoCADHost.AutoCADTask` 与 `System.Threading.Tasks.Task` 同名，
测试文件里起别名 `module AcadTaskCheck = AutoCADHost.AutoCADTask`。

两个已知坑（真机 R23.1 验过）：`(getvar "PID")` 不存在（返回 nil），要证明「跑在哪只 acad 里」
用 `(getvar "DWGNAME")` / `(getvar "ACADVER")`；写文件用 `(vl-princ-to-string …)`，
`itoa` 碰到 nil / 实数直接报错，只留一个 0 字节还被锁住的文件。`EndLisp` 只代表**求值结束**、
不代表成功 —— 要数据就让 LISP 自己写结果文件，F# 侧等「能独占打开且内容非空」。

验证：`-method "*EndLisp*"`、`-method "*发一行 LISP*"`、`-method "*挂回 PID*"`、
`-method "*acad 发内置命令*"`；真机全部要 `XUNITHOST_ACAD_PID=<pid>`，没设就 Skip。

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

## 黑板模式：按信息挂回现成实例

场景：WordVSTO / ExcelVSTO / CuislCad 插件把当前实例抄到黑板上（`[…] 采集时间 …` / KIND / PID / ROT / HWND / EXE / 文档 / 路径），
消费方只按信息精确挂回那一只实例读写；**黑板的创建与销毁不归它管**，不并行，交接时数据要完整、不过时、没被别的任务改过。

**四家入口同形，参数只有一个进程号**：

```fsharp
ExcelHost.ExcelTask.GetApplicationAsync      : int -> Task<Application>
AutoCADHost.AutoCADTask.GetApplicationAsync  : int -> Task<IAcadApplication>
WordHost.WordTask.GetApplicationAsync        : int -> Task<Application>
PowerPointHost.Host.GetApplication           : int -> Result<Application, string>   // 同步
```

同一个语义：前置闸门**同步**判（PID > 0 且确实是对应进程；不合格立刻交回已完成且 faulted 的 Task / `Error`，
不排 STA 队列）→ 合格才排进 `OsHost.Sta.runAsync` 那条固定 STA 线程 → 从 ROT 取现成那只并核 pid：
Excel / AutoCAD / PowerPoint 用 `Hwnd → PID`；Word 用 `ActiveWindow` 上溯 `OpusApp`，零文档时退到 ROT 名后缀。
`unwrap` 抛的异常置成 faulted，await 时抛 `InvalidOperationException`，文案与同步版的 `Error` 一致。

- **`GeneralAppInfo` 已于 2026-11-13 删除**（连同 `verify` / `attachWith` / `defaultAcquire`、`GeneralDocInfo`、
  `GeneralAppInfoTest.fs`）：能唯一确定一只实例的只有 `Pid`；进程名可重算（各库 `SelfImage`），
  `Rot` / `ExePath` 只作佐证，而 `Rot` 在真实调用里一直是空串。
- `OsHost.Com`：`running`（ROT 里取现成实例）、`runningEntries`（枚举，**必须自己核 pid**）、
  `runningAt`（按 moniker 名精确取，**不枚举、不退回 GetActiveObject、绝不新建**）、`register` / `revoke`（自测登记用）。
- **映像名常量**：`WordHost.Host.SelfImage = "WINWORD"`、`ExcelHost.Host.SelfImage = "EXCEL"`、
  `PowerPointHost.Host.SelfImage = "POWERPNT"`；acad 的 `"acad"` 在 `AutoCADHost/Application.fs` 里是**私有**常量
  （`checkPid` 用它核实「确实是这个进程」）。
- **参考解析（不是库 API）**：`xUnitHost/BlackboardText.fs`，`module internal BlackboardText`：
  `caption` / `ofText : string -> Result<BoardApp,string>` / `docOfText : string -> BoardDoc`。
  `BoardApp = { Image; Pid; Rot }`、`BoardDoc = { Name; Path }` 都定义在同一份文件里（`namespace XUnitHost`），
  只服务测试项目；格式强耦合的代码只留在这份里。
- **公寓**：挂回走固定 STA；后续对 Application 的读写也包 `Sta.run`。
- 验证：`-method "*黑板*"`（ROT 探针）、`-method "*不合格*"`（`AttachGateTest`：Word / PPT 闸门就地拒接）、
  `-method "*Excel 挂回*"` 四条（含过时黑板信息按 PID 拒接）。
