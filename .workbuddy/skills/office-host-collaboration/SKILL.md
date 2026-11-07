---
name: office-host-collaboration
description: 当任务落在 xUnitHost 工作区、需要在本机驱动 Word / Excel / AutoCAD 时使用——写自动化步骤、加改测试，或真的要在某份文档、工作簿、图纸上干活。这个技巧规定两种工作模式（独立一条命令 / 合作多步传地址）的选用判据，并强制「先写代码 → 编译 → 跑测试 → 再对真机执行 → 汇报」的闭环，把代码留在项目里当记录。
agent_created: true
---

# 在 xUnitHost 里驱动 Office / AutoCAD

## 这个技巧解决什么

整个解决方案有**七个项目**：操作系统层 `OsHost`（`OsHost.fs` 基础能力 + `Address.fs` 通用进程地址
`ProcAddress` + `ProcHost.fs` 起 / 关两条通用命令）、四家 COM host 库（`WordHost` / `ExcelHost` /
`PowerPointHost` / `AutoCADHost`）、accoreconsole 插件库 `AcadMgdHost`（**只引 acdbmgd + accoremgd，
不走 COM**）、和测试程序 `xUnitHost`（xunit v3）。独立命令行 `ExcelCli` 曾获批准新建，
但 2026-11-06 16:09 用户已把它整个删掉；同日 18:12 用户明确要求新建 `PowerPointHost`（第 6 个项目），
18:48 又要求新建 `AcadMgdHost`（第 7 个项目）。
测试项目的**前部**放公用件：`Common.fs`（报告 / 认人 / 忙重试）、
`Session.fs`（会话层）、`ConvertCore.fs`（Item / Outcome / Paths）、
`Convert.fs`（只分派）；后面全是 Fact。

升级住在**各自 host 库的末尾**，形状完全对称：
`WordHost/DocUpgrade.fs`（.doc 升级；`DocToDocx` / `DocToDocm` / `DocUpgrade` 三个入口合在一份文件里）、
`ExcelHost/XlsUpgrade.fs`（.xls 升级；`XlsToXlsx` / `XlsToXlsm` / `XlsUpgrade`）、
`PowerPointHost/PptUpgrade.fs`（.ppt 升级；`PptToPptx` / `PptToPptm` / `PptUpgrade.pptUpgrade`）；
各自共用的 COM 细节在 `WordHost/DocIO.fs` / `ExcelHost/BookIO.fs` / `PowerPointHost/PptIO.fs`。
前两家接进 `Convert.all` 做批量；`.ppt` 目前是独立函数 `pptUpgrade`，**没接进批量分派**
（避免真目录扫描行为被悄悄扩大），验证在 `xUnitHost/PptUpgradeTest.fs`。

「另一个进程」在测试里仍是测试程序自己再起一份、换一条 Fact 跑
（`SecondProcessTest.fs`）；批量转格式不编命令行工具。
**两条命令（起常驻进程 / 按地址关掉它）都在 `OsHost`**：`Address.fs` 管地址格式的强类型，
`ProcHost.fs` 管 `start` / `startWith` / `close` / `closeWith`（任意进程通用）。
同步版验证在 `xUnitHost/CommandTest.fs` 一条 Fact 里：派发第一条起一只常驻 Excel、把带回来的
`Launched` 打印出来；拿地址派发第二条关掉它、把 `Closed` 打印出来。Excel 特有的 COM 启动 /
ROT 认人 / Quit 以策略回调留在 `module Dispatch`（**两个 Fact 共用这一份**），不写临时脚本。
**Rx 版**在 `xUnitHost/ProcRx.fs` + `RxCommandTest.fs`：把两条同步命令包成
`IObservable<ProcessEvent>`，`Observable.Create` / `SelectMany` / `Timer` / `Merge` / `Timeout`
串成 `Started → Closed → Exited → OnCompleted`，`Process.Exited` 事件转流、`Interval` 轮询兜底、
`Disposable` 收句柄；测试项目引 `System.Reactive` 6.0.2。
`xUnitHost/CliTest.fs` 三条旧 Fact 原依赖已删的 `ExcelCli.exe`，2026-11-06 18:07 用户确认「cli 没用了，可以删除」后已删除，
`xUnitHost.fsproj` 同步去掉 Compile 行；全量不再有 FAIL。

任何「在真的 Word 里改一份文档」这类任务，**不许**临时手敲 COM 调用或写个一次性脚本草草了事——
按下面的模式选型 + 四步闭环走。

## 铁律：先写代码，编译、测试，再执行

1. **写代码**：把要做的事写成 F# 的一个 Fact（或改写已有 Fact）。写在 `xUnitHost/*Test.fs` 里，
   照下面的「测试怎么写」：类 + `ITestOutputHelper`，报告写进 `output`。
2. **编译**：`dotnet build xUnitHost.slnx -v q` —— 必须 0 警告 0 错误。
3. **测试**：`-method "*关键字*"` 单跑这次相关的 Fact，带 `-showLiveOutput`。
   只有确实需要「三类程序都过一遍」时才全量 `-a` —— 全量会把 Word / Excel / AutoCAD 都拉起来。
4. **执行**：测试通过之后，这一步才算真的对真机做过。**副作用就是代码本身**——留在项目里，
   下次谁都能重跑。别删、别写成临时脚本。

**别顺手拉起用不到的程序**。做 Word/Excel 的活就不该动 AutoCAD 或 PowerPoint：测试里只要
`GetOrCreateApplication` 就会在没有实例时真的起一个（AutoCAD 没有「隐藏启动」这一说，
窗口会弹到用户脸上）。规矩有两条：

- **一段通道一个 Fact**，只开这段真正需要的那两家。不要把三家捆进一个 Fact ——
  那样跑一次就把三个程序全拉起来，出问题也说不清是谁的。
  现在拆成 `Word 和 Excel 打招呼` / `Excel 和 AutoCAD 打招呼` 两个。
- **AutoCAD 默认不新起**：本机没开就跳过（`Assert.Skip`），要测那条通道显式
  `XUNITHOST_LAUNCH_ACAD=1`。转换、批量改文档这类任务只跑 `-m` 相关的 Fact，别图省事跑 `-a`。

**动笔之前先认人**：核对**文件名 + 最后实质保存时间**，只往「还没落盘的新文档」里写。
落到磁盘上的一律停手 —— 那是用户手上的文件（跟会话层的地址指纹同一套道理）。
`Doc.identify` / `Doc.describe` / `Doc.ensureFresh` 就是干这个的，写之前一律过一次。
自己新建的文档，收尾记得关掉（`Say.word` / `Say.excel` 会把它交回来）——不然碰到用户开着的实例，
就在人家窗口里留了一份垃圾。

禁止：跳过编译直接跑；绕过测试在 PowerShell 里手搓 COM；改完不跑就汇报「已完成」。

## 测试怎么写（统一这个风格）

照 `FSharp.Idioms.Tests/ArrayStackTest.fs`：**一个类，构造函数拿 `ITestOutputHelper`，Fact 是成员方法**。
不要写成模块级 `let` —— 那样拿不到 `output`，报告只能事后另写一遍。

```fsharp
namespace XUnitHost

open Xunit
open FSharp.xUnit      // Should.equal

type CollaborationTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``方式一 一条命令：开-做-关，host 不留``() =
        Report.title output "方式一　一条命令：开 → 做 → 关，host 不留"
        let text = ... 
        Report.line output "结果" (sprintf "「%s」" text)
        Should.equal "一条命令：写完就走" text     // 期望值在前
        Report.conclusion output "通过：一条命令自己起、自己收，host 不留"
```

**`output` 就是返回信息**。Fact 里写的不是调试打印，而是**给人看的那份报告**：

```
■ 方式二　合作完成：host 不关、凭地址回来接着做、按原状态收尾
--------------------------------------------------------------------
  起点      Word / Excel 都没有在跑，从零各起一个
  action1   开 host   word:10288:1793866298981:0:-:w
  action1   结果      「第一句：Word 先开口」
  action2   结果      「第一句：Word 先开口；第二句：大模型让 Word 接着说」
  收尾      关 host   自己起的两个都关了，地址随之失效
  结论      通过：三个 action 串下来，Word / Excel 各一只，收尾干净
```

规则：
- 标题一行（`Report.title`）、事实若干行（`Report.line` / `Report.line2`）、最后一行 `结论`。
- 一行一个事实，**贴真实的值**（地址、读回来的文字、进程号、报错原文），不要写「成功」两个字。
- 断言放在写报告的行**之间**：先写「发生了什么」，再断言，最后写结论——失败时报告已经留在上面了。
- 时间不要摆裸 Unix 毫秒，用 `Report.clock`。
- 报告工具在 `xUnitHost/Common.fs` 的 `module internal Report`（`title` / `line` / `line2` /
  `conclusion` / `clock` / `ok` / `procIds` / `requireClean` / `selfExe`）。写新 Fact 一律用这一套，
  别再各写一份。

## 选模式

```
要动真机吗？
├─ 只读一个值 / 一步做完 / 不需要中途看结果     → 独立模式（方式一）
├─ 两步以上，每步要靠上一步的结果决定下一步      → 合作模式（方式二）
├─ 需要换一个进程接着做（跨进程交接）            → 合作模式（只能这样）
└─ 只是验证环境（有没有装、能不能起）            → 独立模式
```

**独立模式**：开一次、做一段、关掉，host 不留。适合「查个数」「改一处就走」。
**合作模式**：host 开一次、连着用几步，中间把「地址 + 结果」交回，最后按开场前的状态收尾。
适合「先看看有什么，再决定改哪，再回头验一遍」这种要判断的活。

## 独立模式（方式一）

```fsharp
// 只关心成败
Host.execOnce Kind.Word (fun app -> ...)
// 取回一个值
Host.execOnceQuery Kind.Word (fun app -> ...)
```

语义是「一次进程、一个实例、一段活」，跑完自动 Quit + 清残进程。
注意：AutoCAD 的 `exec` 碰到本机已有实例会直接返回 `Error`（不碰你手上的图）。

## 合作模式（方式二）

```
S0  openSession → 拿到 address（第一步之前的现场，决定收尾怎么关）
S1  写代码：execWord/execExcel/execAcad/execByAddress → (结果, session')
S2  把 **session'.Address** 和结果一起交回；下一步用新地址
S3  ... 重复 S1–S2
S4  close session —— 按 S0 之前的状态关
```

地址长这样，带一整套指纹：

```
word:9232:1793863068197:1793863070426:790cb57f:w
     pid  进程启动毫秒   最后保存毫秒    路径指纹 归属
```

- **每步都要用新地址**：`exec` 返回的是重拍过快照的 `Session`。这一步要是保存了文档，
  保存时刻和路径指纹就变了，地址跟着变，**老地址会被移除并拒接**。
- **绝不自拼地址**，也绝不把老地址硬塞回去试探。被拒是在正常工作——说明状态漂移了。
- 被拒时错误信息里会给出「这只 host 现在的地址是 …」，拿那个；实在没有就重新 `openSession`。
- **首次进入 ≠ 按地址回来**：`openSession` 是「对外签约」（可以真的起一个进程、记谁起的）；
  `resolve` 是「同事打招呼」（只接现成的，**绝不新建**）。别在回来的时候用 `GetOrCreateApplication`。
- **收尾必须 `close`**：只有 `Owned = true`（本次起的）才真关（丢改动 + Quit + 清残进程）；
  你原来就开着的那个，一个字都不动。

## 批量转格式（doc → docx / docm、xls → xlsx / xlsm；ppt 走独立入口）

转换实现在**各自 host 库的末尾**，三家形状完全对称，各一份文件放三个入口：

- `WordHost/DocUpgrade.fs`：`DocToDocx`（无宏→.docx）、`DocToDocm`（有宏→.docm）、
  `DocUpgrade`（统一入口：`DocIO.detectMacro` 判宏 → 纯规则 `route` 选路 → 调对应转换）；
- `ExcelHost/XlsUpgrade.fs`：`XlsToXlsx`（无宏→.xlsx）、`XlsToXlsm`（有宏→.xlsm）、
  `XlsUpgrade`（同一个判宏选路形状）；
- `PowerPointHost/PptUpgrade.fs`：`PptToPptx`（无宏→.pptx，格式码 24）、
  `PptToPptm`（有宏→.pptm，格式码 25）、`PptUpgrade.pptUpgrade`（同一个判宏选路形状，
  `convert` 是对齐别名）。**这条暂不接进 `Convert.all`**：批量扫描仍只认 `.doc` / `.xls`，
  要转 .ppt 就在 `Host.execQuery` 里直接调 `pptUpgrade app <路径>`，验证见 `PptUpgradeTest.fs`。

共用的开/关、正文/格子值/幻灯片文字指纹、判宏、临时文件在
`WordHost/DocIO.fs` / `ExcelHost/BookIO.fs` / `PowerPointHost/PptIO.fs`。
**六条转换各自独立**——格式码、目标后缀、核对口径、结果文案各归其主；
host 库只回传 `(目标, 状态, 详情)`，不认识测试项目的 `Outcome`（依赖方向不倒挂）。
测试项目前部的 `ConvertCore.fs` 只剩 Item / Outcome / Paths，`Convert.fs` 只做
`Paths.scan` → `Convert.all` → 调 `DocUpgrade.convert` / `XlsUpgrade.convert` → 组装结果 → 逐份读回核对。
判不出有没有宏就报失败、原件保留，绝不猜「没宏」（猜错会悄悄丢宏）。
**转格式不编命令行程序**：真目录的活是两条 Fact，设环境变量才动，不设就跳过
（起 / 关常驻进程那两条命令在 `OsHost`，不承担转换）。

```bash
exe=xUnitHost/bin/Debug/net10.0/xUnitHost.exe

# 1. 只清点：要转哪些、哪些已有同名新文件（一个文件都不动）
XUNITHOST_CONVERT_DIR='C:\某个目录' "$exe" -noColor -showLiveOutput -method "*真目录*清点*"

# 2. 真转：每份读回来比对，核对通过的才删原件（没通过的原件一个不动）
XUNITHOST_CONVERT_DIR='C:\某个目录' "$exe" -noColor -showLiveOutput -method "*真目录*转*"

# 每批份数默认 10，可以改
XUNITHOST_CONVERT_DIR='C:\某个目录' XUNITHOST_CONVERT_BATCH=5 "$exe" ... 
```

**目标格式由有没有宏决定**：`.doc` 没宏 → `.docx`、**有宏 → `.docm`**；
`.xls` 没宏 → `.xlsx`、**有宏 → `.xlsm`**；`.ppt` 没宏 → `.pptx`、**有宏 → `.pptm`**（VBA 保住）。
这是默认行为，不用问用户；转完还要再确认一遍宏还在（读回来 `HasVBProject` 为真）。

```bash
铁律照旧，而且更严：
- **先跑测试** `-method "*转换*"`（`xUnitHost/ConvertTest.fs` 在临时目录造样例验证），通过才碰真文件。
  真目录的入口也是同一个 exe，只是加环境变量（见上面两条命令）。
- **先备份原件**到目录外交给用户（保持目录结构、逐份比大小），备份好了才允许删。
- **转一份、读回来比一份**：正文/格子值不一致就判失败、原件保留，绝不「先删后转」。
- 删原件只认「核对通过」那批（`Convert.canDelete`），一条一条打印，删完再扫一遍确认只剩该剩的。
- 带 VBA 宏的 `.doc` 转 `.docm`、`.xls` 转 `.xlsm`（不是跳过、也不是硬转无宏格式丢宏）；
  `.docm` / `.xlsm` 都是 ZIP，转完可以直接查 `word/vbaProject.bin` / `xl/vbaProject.bin` 在不在，双重确认。
- 递归扫时会自动跳过 `.idx_cache` 这类缓存目录和 `~$` 临时锁文件，别把它们当文档。

## 代码写在哪

| 东西 | 位置 | 约定 |
|---|---|---|
| 新任务 / 新验证 | `xUnitHost/CollaborationTest.fs` 加 Fact | Fact 是类的成员方法，名字写中文一句话；报告写进 `output` |
| 两条命令的派发验证 | `xUnitHost/CommandTest.fs` | 一条 Fact 两段派发：`ProcHost.startWith` 起 Excel 打印 `Launched`、`ProcHost.closeWith` 按地址关掉打印 `Closed`；Excel 策略在 `module Dispatch`（`-m "*派发*"`） |
| 两条命令的 Rx 编排 | `xUnitHost/ProcRx.fs` + `RxCommandTest.fs` | `ProcessEvent` 判别联合 + `Observable.Create` / `SelectMany` / `Timer` / `Merge` / `Timeout`；`Process.Exited` 事件流 + 轮询兜底；退出检测放 close 一侧、命令前订阅（`-m "*Rx 两条命令*"`） |
| 报告的写法 | `xUnitHost/Common.fs`，`module internal Report` | `title` / `line` / `line2` / `conclusion` / `clock` / `ok` / `requireClean` |
| 认人 / 忙重试 | 同上，`module Doc` / `module Busy` | `Doc.identify` → `Doc.ensureFresh`；`Busy.retry` 扛 COM 忙 |
| 会话层能力 | `xUnitHost/Session.fs`（测试项目**前部**） | 地址 + 指纹、多步执行、按原状态收尾；新能力加在这儿 |
| 操作系统功课 | `OsHost/OsHost.fs` + `Address.fs` + `ProcHost.fs` | STA 线程、进程账本、窗口 → 进程号、时间、COM 底层；**通用进程地址 `ProcAddress` 与起 / 关两条通用命令**（任意进程，不只 Office，2026-11-06 从 ExcelCli 上收）；三家 host 共用 |
| accoreconsole 插件 | `AcadMgdHost/{AcadMgdHost.fsproj,Commands.fs,AssemblyInfo.fs}` + `test/run.cmd` | netstandard2.0 F#；**只引 `acdbmgd.dll` + `accoremgd.dll`（`Private=false`），不引 `acmgd.dll` 等 UI DLL、不引 `OsHost`**；`XH_HELLO` / `XH_DB` 两个 CommandMethod + `Entry : IExtensionApplication`；`run.cmd` 用 `/i` 种子 DWG + 绝对路径 `/s`。与 `AutoCADHost` 的 COM 路线无关，实测命令注册正常、数据库 API 可用；`Console.Out` 被重定向到 `%TEMP%\accc*.tmp`，所以插件另按 `ACADMGDHOST_LOG` 写日志。细节见 `references/api.md` |
| COM 细节 + 转换 | `WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost` | 与 Office 强相关的操作放对应库（六条转换也挂在各自 host 库末尾）；各自会 `identity` 报自己是谁（`PowerPointHost` 只做方式一，没有会话层） |
| ppt 升级函数 | `PowerPointHost/PptIO.fs` + `PptUpgrade.fs`；验证 `xUnitHost/PptUpgradeTest.fs`（`-m "*pptUpgrade*"`） | `pptUpgrade app <路径>`：`PptIO.detectMacro` 判宏 → 纯规则 `route` → `PptToPptx`(24) / `PptToPptm`(25)；转完读回幻灯片文字指纹、已有同名先转临时文件、判不出宏失败且原件保留。PowerPoint 的 Open/SaveAs 是**强类型按值**参数；`WithWindow=msoFalse` 不弹窗口；`Application.Visible` 不许设 msoFalse |
| 批量转格式 | `xUnitHost/ConvertCore.fs` + `Convert.fs`（**前部**）；`WordHost/DocIO.fs` + `DocUpgrade.fs`、`ExcelHost/BookIO.fs` + `XlsUpgrade.fs`（两个 host 库**末尾**） | 两侧对称：各一份文件放三个入口 —— 两条转换各自独立（格式码/后缀/核对/文案各归其主）+ 一个统一入口判宏 → 纯规则 `route` 选路；共用 COM 细节在各自的 `*IO.fs`；host 只回传 `(目标, 状态, 详情)`，`Convert.fs` 组装 Outcome、只留 `Paths.scan` / `Convert.all` / `readWordText` / `readBookPrint` / `canDelete`；真目录入口是两条按环境变量跑的 Fact（`ConvertTest.fs` 末段） |
| 黑板展示文本参考解析 | `xUnitHost/BlackboardText.fs`（测试项目前部） | `module internal BlackboardText`：`caption` / `ofText` 把插件文本填成 `AppInfo` 参数表，只供调用方参考；**`OsHost` 不认文本，只认结构化参数表** |
| 通用参数表（另一条路） | `OsHost/GeneralAppInfo.fs` + `xUnitHost/GeneralAppInfoTest.fs` | 与 `AppKind` 解耦：`GeneralIdentity` 三个参数占位、`attachWith` 的 `acquire` / `work` 获取函数占位；`-method "*通用参数表*"`；12:01 落地时原 `AppInfo.fs` 未动，13:51 两路一起撤了文档级交接检查 |
| Excel 的通用入口 | `ExcelHost/Host.fs` 的 `Host.General`（旧 `GetApplication` 正下方）+ `xUnitHost/ExcelGeneralTest.fs` | `Host.General.GetApplication(info: GeneralAppInfo)`：`Tag="excel"` 公开常量、`GeneralAppInfo.attach` + 强转 Excel.Application + pid 核对；旧入口与其余三家未动；`-method "*Excel 通用参数表*"` 两条 Fact，ROT 探针走完整链路、不启真机 Excel |
| 跨进程 | `xUnitHost/SecondProcessTest.fs` | 第二个进程 = 测试程序自己；地址 / 要说的话 / 回话文件都从环境变量来 |

**项目的增减由用户定**：往现有项目里**加文件随时可以**（新 Fact、新公用模块都行）；但
**不许私自新建项目**。原五个：`OsHost` / `WordHost` / `ExcelHost` / `AutoCADHost` / `xUnitHost`。
第六个 `ExcelCli` 是 2026-11-06 上午讲清理由后问过用户才建的，同日 16:09 又被用户删掉；
同日 18:12 用户**明确要求**「新建 PowerPointHost 项目」，18:48 又要求
「参考样例新建一个 AcadMgdHost 的 fsharp, standard2.0 类库项目」；于是 `xUnitHost.slnx` 现在是
**七个项目**（`PowerPointHost` + `AcadMgdHost`），`OsHost/Address.fs` / `ProcHost.fs` 保留。
再要开新项目，照旧先问用户。

公用代码往上收（操作系统那类进 `OsHost`，测试杂活进 `Common.fs`），其余放测试项目前部。
`HostSession`、`DocConvert`、`CrossProcessWorker` 三个小项目都是这么并掉的。

- 测试**串行**跑（`xUnitHost/xunit.runner.json` 里 `parallelizeTestCollections: false`），别并行抢实例。
- xunit v3 **不能用 `dotnet test`**（VSTest 报不支持），直接跑 exe。
- 测试要求「一开始没有对应进程在跑」时，在 Fact 开头用 `requireClean "WINWORD"` 之类显式声明；
  环境不干净就 **skip**（`Assert.Skip`），不要报失败，更不要为了跑通去关掉用户的 Word/Excel。
  用户开着图纸和表格干活是常态——跳过让整轮跑下来是「通过 + 跳过」，而不是一片红。

## 命令

```bash
dotnet build xUnitHost.slnx -v q

# 单跑一个 Fact（按 Fact 名里的关键字）。-showLiveOutput 让 output 实时打出来
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*关键字*" -noColor -showLiveOutput
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*派发*" -noColor -showLiveOutput          # 同步版两条命令
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*Rx 两条命令*" -noColor -showLiveOutput  # Rx 版两条命令
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*pptUpgrade*" -noColor -showLiveOutput    # ppt 升级（两条 Fact）

# 全量（约 55–60 秒）。注意：会把三类程序都拉起来（acad 除外，见下面的闸门）
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -noColor -showLiveOutput
```

`-showLiveOutput` 不能省：不加的话报告被 xunit 收走，只能看到一行 Summary。
加了之后每行带 `[OUTPUT]` 前缀，抄给用户时把前缀和测试名去掉即可。

或者用脚本（编译 + 跑 + 现场检查，一步到位）：

```bash
bash .workbuddy/skills/office-host-collaboration/scripts/run-facts.sh -m "*关键字*"
bash .workbuddy/skills/office-host-collaboration/scripts/run-facts.sh -a
bash .workbuddy/skills/office-host-collaboration/scripts/run-facts.sh -a -n   # 跳过编译
```

脚本在跑之前先给本机的 `WINWORD` / `EXCEL` / `acad` 拍个快照，跑完再拍一次：
**多出来的**进程算「没关干净」（脚本返回 1），**本来就在跑的**只列出来，不会被误认。

## 黑板模式（插件抄现场 → 按 info 挂回现成实例）

插件把「当前实例」抄到黑板上（采集时间 / PID / ROT 名 / HWND / EXE / 文档 / 路径），
WorkBuddy 只读写、**不负责黑板的创建与销毁**。四家 host 各有一个
`GetApplication(info) : Result<Application,string>`：按 ROT 名字精确附着 + 现场校验，
绝不新建、绝不 Quit、不改应用状态；信息过时 / 种类错配 / 认错进程就 `Error`。
**文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求推迟**（2026-11-07 13:51）：两路库与四家入口都不做，
`DocName` / `DocPath` 字段保留、只是收下的数据，以后再接回来（旧实现在 git 历史 `6be1725` 之前）。
**`AppInfo` 是结构化参数表，调用方按字段填报**（直接构造记录，或用类型上的静态构造成员 `AppInfo.create(...)` 起最小集再 `with` 补齐），`OsHost` 里没有文本解析（`ofText` 已移出生产库）；
把插件展示文本填成参数表的参考写法是测试项目 `xUnitHost/BlackboardText.fs`
（`BlackboardText.ofText`，只供参考、不是库 API），插件改排版只改这一份。
共用件在 `OsHost/AppInfo.fs` + `OsHost.Com.runningAt`，完整 API 见 `references/api.md`。
验证 `xUnitHost/BlackboardTest.fs` 四条 Fact（`-method "*黑板*"`），真机那条靠 `XUNITHOST_BLACKBOARD_FILE`。
返回的 Application 与后续读写都放在 `OsHost.Sta.run` 那条固定 STA 线程上。

**另一条路（2026-11-07 12:01，用户要求「原文件不动，另起炉灶」）**：`OsHost/GeneralAppInfo.fs` 的
`GeneralAppInfo` 与 `AppKind` 完全解耦 —— 进程名 / 显示名 / 种类标记由 `GeneralIdentity` 三个**参数**占位，
「怎么取对象」「拿到 obj 之后怎么转强类型」由 `attachWith` 的 `acquire` / `work` 两个**获取函数**占位。
现场校验口径与 `AppInfo` 一致，两条路并存；**14:05 起 Excel 另有 `ExcelHost.Host.General.GetApplication(info: GeneralAppInfo)`**
（嵌套模块保留函数名，Tag 公开为 `Host.General.Tag`），其余三家 host 仍在旧路上。
验证 `-method "*通用参数表*"`（现五条：`GeneralAppInfoTest.fs` 三条 + `ExcelGeneralTest.fs` 两条），
只跑 Excel 入口用 `-method "*Excel 通用参数表*"`；完整 API 见 `references/api.md`。
生产层不知道 `AppKind` 存在，桥接示例只在测试项目。
13:51 起两路都不做文档级交接检查（见上），`DocName` / `DocPath` 只是收下的数据。

## 汇报格式

做完给用户四段，缺一不可：

1. **做了什么**：一句话说清最终目的（不是「改了很多文件」）。
2. **结果**：直接贴 Fact 的 `output` 报告（去掉 `[OUTPUT]` 前缀和测试类名），别另写一份、别只说「成功」。
3. **留下的代码**：具体到文件与 Fact 名，说明这就是可重跑的记录。
4. **现场干净**：`WINWORD` / `EXCEL` / `POWERPNT` / `acad` 有没有残留（脚本最后会打印 Office 三家；
   跑过 ppt 的活另外单查 `POWERPNT.EXE`）。

## 细节参考

- `references/api.md` —— 会话层 / `OsHost` / 三个 host 库的完整 API、地址与指纹规则、类型要点。
- `references/gotchas.md` —— 已经踩过的坑（STA、保存对话框、RCW 身份缓存、Word 零文档没有 pid…），
  动手前扫一眼，能省掉一半调试时间。
