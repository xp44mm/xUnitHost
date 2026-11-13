---
name: office-host-collaboration
description: 当任务落在 xUnitHost 工作区、需要在本机驱动 Word / Excel / PowerPoint / AutoCAD 时使用——写自动化步骤、加改测试，或真的要在某份文档、工作簿、图纸上干活。这个技巧规定两种工作模式（独立一条命令 / 合作多步传地址）的选用判据，并强制「先写代码 → 编译 → 跑测试 → 再对真机执行 → 汇报」的闭环，把代码留在项目里当记录。访问代码一律用 `task {}` 表达式串 `*Async` 入口；驱动 AutoCAD 一律用 `AutoCADHost.AutoCADTask` 里的现成函数（命令等 EndCommand、LISP 等 EndLisp、按 PID 挂回），不要自己拼 SendCommand 或订阅事件。
agent_created: true
---

# 在 xUnitHost 里驱动 Office / AutoCAD

## 这个技巧解决什么

整个解决方案有**七个项目**：操作系统层 `OsHost`（`Clock` / `Paths` / `Errors` / `Process` / `Win32` / `COM` / `STA` /
`TextFile` / `Rot` / `Identity` 各一份）、四家 COM host 库
（`WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost`）、accoreconsole 插件库 `AcadMgdHost`
（**只引 acdbmgd + accoremgd，不走 COM**）、和测试程序 `xUnitHost`（xunit v3）。
**不许私自新建项目**：往现有项目里加文件随时可以，要开新项目先问用户。
测试项目的**前部**放公用件（顺序即 `xUnitHost.fsproj` 的顺序）：`Common.fs`（报告 / 认人 / 忙重试）、
`Session.fs`（会话层 `Host.*`）、`ConvertCore.fs`（Item / Outcome / Paths）、
`Convert.fs`（只分派）、`BlackboardText.fs`（黑板文本参考解析）；后面全是 Fact。

升级住在**各自 host 库的末尾**，形状完全对称：
`WordHost/DocUpgrade.fs`（.doc 升级；`DocToDocx` / `DocToDocm` / `DocUpgrade` 三个入口合在一份文件里）、
`ExcelHost/XlsUpgrade.fs`（.xls 升级；`XlsToXlsx` / `XlsToXlsm` / `XlsUpgrade`）、
`PowerPointHost/PptUpgrade.fs`（.ppt 升级；`PptToPptx` / `PptToPptm` / `PptUpgrade.pptUpgrade`）；
各自共用的 COM 细节在 `WordHost/DocIO.fs` / `ExcelHost/BookIO.fs` / `PowerPointHost/PptIO.fs`。
前两家接进 `Convert.all` 做批量；`.ppt` 目前是独立函数 `pptUpgrade`，**没接进批量分派**
（避免真目录扫描行为被悄悄扩大），验证在 `xUnitHost/PptUpgradeTest.fs`。

`ExcelHost/SheetCopy.fs` 是另一条路：老 `.xls` 一张一张 `Worksheet.Copy` 抄进副本，再删新格式带不走的东西
（图表 / 图片 / 控件 / OLE / 批注 / 条件格式 / 数据验证 / 超链接 / 查询表 / 智能表 / 数据透视表），只留公式与普通单元格格式。
直接 `SaveAs` 转这些表常常报错甚至弄崩 Excel，逐页抄能定位到具体哪张表；**原件只读、一个字不改**。

「另一个进程」在测试里仍是测试程序自己再起一份、换一条 Fact 跑
（`SecondProcessTest.fs`）；批量转格式不编命令行工具。
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
自己新建的文档，收尾记得关掉（`Host.QuitWithoutSaving` / `CloseAllWorkbooksWithoutSaving`；「打招呼」那两条
Fact 用的 `Say.word` / `Say.excel` 在 `xUnitHost/HostGreetingTest.fs` 里，会把文档交回来）——不然碰到用户开着的实例，
就在人家窗口里留了一份垃圾。

禁止：跳过编译直接跑；绕过测试在 PowerShell 里手搓 COM；改完不跑就汇报「已完成」。

## 测试怎么写（统一这个风格）

**访问代码一律用 `task {}` 表达式**，Fact 本身就写成异步任务定义（`Task` 返回值，xunit v3 直接认）。
别写成同步 `unit` 方法再在里面 `let work = task { … }` 用 `.GetAwaiter().GetResult()` 收尾——
那样异常被包一层、还白占一条线程。存量文件里有旧写法，新写的一律按下面这个来：

- 先拿 `*Async` 入口：`ExcelTask.GetApplicationAsync pid` / `GetOrCreateApplicationAsync()`、
  `WordTask.GetApplicationAsync info`、`AutoCADHost.AutoCADTask.*Async`。
- 库里的同步入口要排进固定 STA 时，用 `let! x = Sta.runAsync(fun () -> …).Unwrap()`。
- `let!` / `do!` 不许省（`let _ = xxxAsync` 是 fire-and-forget，等于没等，「直线没画上」就是这么来的）。
- 取回的 COM 对象（RCW）在使用它的每一步仍要包 `Sta.run` / `Sta.runAsync`。

照 `FSharp.Idioms.Tests/ArrayStackTest.fs`：**一个类，构造函数拿 `ITestOutputHelper`，Fact 是成员方法**。
不要写成模块级 `let` —— 那样拿不到 `output`，报告只能事后另写一遍。

```fsharp
namespace XUnitHost

open System
open System.Threading.Tasks
open Xunit
open FSharp.xUnit      // Should.equal
open OsHost

type ExcelAttachTest(output: ITestOutputHelper) =

    [<Fact>]
    member this.``Excel 挂回真机 按 PID 取到同一只实例``() : Task =
        task {
            Report.title output "Excel　按 PID 挂回（只读，不动表）"

            let pidText = Environment.GetEnvironmentVariable "XUNITHOST_EXCEL_PID"

            if String.IsNullOrWhiteSpace pidText then
                Assert.Skip("没设 XUNITHOST_EXCEL_PID：这条要在活着的 Excel 上做")

            let pid = int pidText
            let! app = ExcelHost.ExcelTask.GetApplicationAsync pid

            // RCW 只在挂 RCW 的那条 STA 线程上用；只读一个值，不碰单元格
            let name = Sta.run(fun () -> app.ActiveWorkbook.Name)

            Report.line output "挂回" (sprintf "PID %d　工作簿「%s」" pid name)
            Should.equal true (name.Length > 0)              // 期望值在前
            Report.conclusion output "通过：按 PID 认到同一只实例，不新建、不 Quit"
        }
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
S1  写代码：execWord/execExcel/execByAddress → (结果, session')   // acad 不进会话层，见「驱动 AutoCAD」一节
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

## 驱动 AutoCAD：直接用 `AutoCADHost.AutoCADTask` 里的函数

acad 的活**不许**自己拼 `doc.SendCommand`、自己 `add_BeginCommand`、或写个一次性脚本草草了事。
`AutoCADHost/AutoCADTask.fs`（`module AutoCADHost.AutoCADTask`）已经按「等的到底是哪个事件」分好了类，
挑一个用；缺什么再加到那儿，别在 Fact 里另起一套。它是 acad 的 COM 驱动入口
（`Application.fs` 只剩 `checkPid` / `findOnSta` / `unwrap` 这些零件；`AutoCADHost` 现在没有 `Host.fs`）。

| 函数 | 签名 | 等什么 | 什么时候用 |
|---|---|---|---|
| `AutoCADTask.GetApplicationAsync` | `int -> Task<IAcadApplication>` | 无（同步闸门 + 固定 STA 取对象） | 按 **PID** 挂回活着的那只，**唯一**挂回入口；不新建、不 Quit |
| `AutoCADTask.sendCommandAsync` | `EventPair.Options -> IAcadDocument -> string -> Task<string>` | `BeginCommand` → `EndCommand` | 内置命令（`_.LINE` / `_.CIRCLE` / `_.ZOOM`…），结果 = 命令名 |
| `AutoCADTask.commandAsync` | `EventPair.Options -> IAcadApplication -> (unit -> unit) -> Task<string>` | 同上 | 启动动作不是「发命令行」而是自己调 COM 方法时 |
| `AutoCADTask.openDrawingAsync` | `EventPair.Options -> IAcadApplication -> (unit -> unit) -> Task<string>` | `BeginOpen` → `EndOpen` | 打开图纸（`start` 里发 `Documents.Open`），比撞忙重试可靠 |
| `AutoCADTask.modalAsync` | `EventPair.Options -> IAcadApplication -> Task<unit>` | `BeginModal` → `EndModal` | 有模态对话框挡着，等它关掉再发下一条 |
| `AutoCADTask.runLispAsync` | `IAcadDocument -> string -> Task<LispOutcome>` | `BeginLisp` → `EndLisp`；Esc 走 `LispCancelled` | 一整段 LISP 要跑完才继续（默认 30 s、认 BeginLisp） |
| `AutoCADTask.sendLispAndWaitAsync` | `EventPair.Options -> IAcadDocument -> string -> Task<LispOutcome>` | 同上 | 要自定义超时 / 取消 / 不认 BeginLisp 时 |
| `AutoCADTask.sendLispAsync` | `IAcadDocument -> string -> Task<unit>` | **不等** | 只交出去，后面没有串行动作 |
| `AutoCADTask.runLispFileAsync` | `IAcadDocument -> string -> Task<LispOutcome>` | `EndLisp` | `(load "…")` 且要等文件里的代码跑完 |
| `AutoCADTask.loadLispAsync` | `IAcadDocument -> string -> Task<unit>` | 不等 | `(load "…")` 只交出去 |

**四条必须先想清楚的语义**（踩过的都在里面）：

1. **`SendCommand` 返回 ≠ 命令跑完**。它同步返回只代表「已经交给 AutoCAD」，所以并发/串行
   一律用 `sendCommandAsync`，等 EndCommand 到了才发下一条。裸 `SendCommand` 之后立刻发下一条
   → 撞 `RPC_E_CALL_REJECTED`。
2. **命令行串必须自带换行收尾**（`"_.ZOOM\n_E\n"`）。少一个换行，最后一个词留在命令行上等输入，
   整只 acad 会一直「忙」，后续所有调用都被拒。
3. **LISP 走的是 Lisp 事件，不是命令事件**：`BeginLisp` / `EndLisp` / `LispCancelled`。
   一整段 `(progn …)` 是一个求值整体，内部 `(command "_.LINE" …)` 的 Begin/EndCommand 先各自走完，
   最后才来 EndLisp —— 所以等命令事件会在第一条内嵌命令结束时**提前返回**。
   真机已验证（R23.1）：`SendCommand` 送进去的 LISP 确实派 `BeginLisp → EndLisp`，默认
   `RequireBegin = true` 可用；万一某版不派开始，传 `{ EventPair.defaults with RequireBegin = false }`。
4. **`EndLisp` = 求值结束，≠ 成功**。AutoLISP 的求值错误只在命令行打印一行、不回报调用方。
   要拿数据或判成败，就让 LISP 自己写结果文件（`open` / `write-line` / `close`），
   F# 侧等「能独占打开且内容非空」—— `File.Exists` 不算，`open` 之后文件就存在了但内容是空的。

LISP 里两个具体坑：`(getvar "PID")` 在 R23.1 上**没有这个系统变量**（返回 `nil`）；
写文件一律 `(vl-princ-to-string …)`，别用 `itoa`（碰到 `nil` / 实数直接报错，只留下 0 字节文件，
句柄还挂在 acad 里把文件锁住）。要证明「跑在哪一只里」，写 `DWGNAME` / `ACADVER`
对照 COM 读到的 `ActiveDocument.Name`。

**线程**：`AutoCADTask.fs` 里所有入口都必须在挂着 RCW 的那条 STA 线程上调用，`start` 也在调用线程上同步跑。
套 `OsHost.Sta.runAsync`，别在原地裸调。`options` 一般直接用 `EventPair.defaults`
（认 Begin、30 秒、不取消），要改就 `{ EventPair.defaults with Timeout = Some (TimeSpan.FromSeconds 60.0) }`：

```fsharp
// 库模块与 System.Threading.Tasks.Task 同名，起个别名
module AcadTaskCheck = AutoCADHost.AutoCADTask

let t =
    task {
        // 1. 按 PID 挂回（PID 不合法 / 不是 acad → 已完成且 faulted 的 Task，绝不新建）
        let! acad = AcadTaskCheck.GetApplicationAsync pid

        // 2. 内置命令：等 EndCommand 才返回
        let! name =
            Sta.runAsync(fun () -> AcadTaskCheck.sendCommandAsync options acad.ActiveDocument "_.ZOOM\n_E\n")
                .Unwrap()

        // 3. LISP：等 EndLisp（LispOutcome = Completed | Cancelled）
        let! outcome = Sta.runAsync(fun () -> AcadTaskCheck.runLispAsync acad.ActiveDocument "(setq probe 4242)").Unwrap()
        return name, outcome
    }
```

`acad` 这个 RCW 也要在 `Sta.run` 里用（同一条固定长活 STA 线程）；`GetApplicationAsync`
只负责把它取出来，不替你切线程。

**真机 Fact 的闸门**：`XUNITHOST_ACAD_PID=<pid>`，没设就 `Assert.Skip`。按 PID 挂回而不是
`GetOrCreateApplication` 是硬要求 —— acad 没有「隐藏启动」，起一只就把窗口弹到用户脸上。
只读验证（`_.DIST`、`(getvar "DWGNAME")`）优先，别往用户的图纸里加实体。

## 批量转格式（doc → docx、xls → xlsx / xlsm；ppt 走独立入口）

转换实现在**各自 host 库的末尾**，三家形状完全对称，各一份文件放三个入口：

- `WordHost/DocUpgrade.fs`：`DocToDocx`（→.docx）、`DocToDocm`（→.docm，**只在 `XUNITHOST_DOC_KEEP_MACRO=1` 时走**）、
  `DocUpgrade`（统一入口：`DocIO.detectMacro` 判宏 → **默认一律 `.docx`**，用户 2026-11-11 定「Word 的宏不留」）；
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
Excel / PowerPoint 侧判不出有没有宏就报失败、原件保留，绝不猜「没宏」（猜错会悄悄丢宏）。
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

**目标格式**：`.doc` **一律 `.docx`（宏不保留，2026-11-11 用户定）**；
`.xls` 没宏 → `.xlsx`、**有宏 → `.xlsm`**；`.ppt` 没宏 → `.pptx`、**有宏 → `.pptm`**（VBA 保住）。
Excel / PowerPoint 转完还要再确认一遍宏还在（读回来 `HasVBProject` 为真）；Word 的 `.docm` 只有显式设
`XUNITHOST_DOC_KEEP_MACRO=1` 才走（老 `.doc` 的 VBA 工程多半被「锁定工程」保护，硬保只会白卡文件）。

铁律照旧，而且更严：
- **先跑测试** `-method "*转换*"`（`xUnitHost/ConvertTest.fs` 在临时目录造样例验证），通过才碰真文件。
  真目录的入口也是同一个 exe，只是加环境变量（见上面两条命令）。
- **先备份原件**到目录外交给用户（保持目录结构、逐份比大小），备份好了才允许删。
- **转一份、读回来比一份**：正文/格子值不一致就判失败、原件保留，绝不「先删后转」。
- 删原件只认「核对通过」那批（`Convert.canDelete`），一条一条打印，删完再扫一遍确认只剩该剩的。
- `.xls` 带宏转 `.xlsm`（不是跳过、也不是硬转无宏格式丢宏）；`.xlsm` / `.docm` 都是 ZIP，
  转完可以直接查 `xl/vbaProject.bin` / `word/vbaProject.bin` 在不在，双重确认。Word 侧默认不保宏（见上）。
- 递归扫时会自动跳过 `.idx_cache` 这类缓存目录和 `~$` 临时锁文件，别把它们当文档。

## 代码写在哪

| 东西 | 位置 | 约定 |
|---|---|---|
| 新任务 / 新验证 | `xUnitHost/CollaborationTest.fs` 加 Fact | Fact 是类的成员方法，名字写中文一句话；报告写进 `output` |
| 报告的写法 | `xUnitHost/Common.fs`，`module internal Report` | `title` / `line` / `line2` / `conclusion` / `clock` / `ok` / `requireClean` |
| 认人 / 忙重试 | 同上，`module Doc` / `module Busy` | `Doc.identify` → `Doc.ensureFresh`；`Busy.retry` 扛 COM 忙 |
| 会话层能力 | `xUnitHost/Session.fs`（测试项目**前部**） | 地址 + 指纹、多步执行、按原状态收尾；新能力加在这儿 |
| 操作系统功课 | `OsHost` 各份（`STA` / `Process` / `Win32` / `COM` / `Clock` / `Paths` / `Errors` / `Rot` / `Identity` / `TextFile`） | STA 线程、进程账本、窗口 → 进程号、时间、COM 底层、现场快照 `Identity`；三家 host 共用 |
| 读文本文件（认编码 / 切行） | `OsHost/TextFile.fs` | `open OsHost` 后 `lines path` / `dec path`（= `TextFile.decodeAllLinesAsync`，`Task<编码名 * 行数组>`）：BOM → ascii → UTF-8/GBK 严格试解码（歧义报 `utf-8\|gbk`）→ 按 `\r\n` / `\n` / `\r` 切行且末尾不多空行。**读任何来路不明的文本一律用它**，别写 `File.ReadAllLines`、别自己猜 Encoding、别手切行；自检 `-m "*文本读取*"`（离线） |
| ROT 名 ↔ 进程号 | 通用切 / 拼 `OsHost/Rot.fs`；前缀固化在各家 `WordHost/Rot.fs`、`PowerPointHost/Rot.fs` | 通用层只有两件纯字符串活儿：`OsHost.Rot.parse`（`!WordVSTO.Instance.27232` → `Some(前缀, 27232)`）与 `OsHost.Rot.build`（前缀 + pid → 名字），**不认识任何具体前缀、不碰 COM**。各家的 `Rot.Prefix` 是 `[<Literal>]`，`Rot.pidOf` 只认自家前缀（别人的登记项一律 None），`Rot.nameOf` 由前缀拼出来。Excel / AutoCAD 现在只按 `Hwnd → pid` 认人，没有名字路径，所以没建各自的前缀。自检 `-m "*ROT 前缀*"`（离线） |
| accoreconsole 插件 | `AcadMgdHost/{AcadMgdHost.fsproj,Commands.fs,Probe.fs,Star.fs,StarGeometry.fs,AssemblyInfo.fs}` + `test/run.cmd` | netstandard2.0 F#；**只引 `acdbmgd.dll` + `accoremgd.dll`（`Private=false`），不引 `acmgd.dll` 等 UI DLL、不引 `OsHost`**；`XH_HELLO` / `XH_DB` / `XH_STAR` 三个 CommandMethod + `Entry : IExtensionApplication`；`run.cmd` 用 `/i` 种子 DWG + 绝对路径 `/s`。与 `AutoCADHost` 的 COM 路线无关；`Console.Out` 被重定向到 `%TEMP%\accc*.tmp`，所以插件另按 `ACADMGDHOST_LOG` 写日志。细节见 `references/api.md` |
| COM 细节 + 转换 | `WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost` | 与 Office 强相关的操作放对应库（六条转换也挂在各自 host 库末尾）；各自有 `Host.SelfImage` 说自己是哪个进程（`PowerPointHost` 只做方式一，没有会话层） |
| acad 驱动（命令 / LISP / 打开 / 模态 / 挂回） | `AutoCADHost/AutoCADTask.fs`（`module AutoCADHost.AutoCADTask`） | **一律用这里的函数，不许手搓 `SendCommand` / 自己订阅事件**；命令等 EndCommand、LISP 等 EndLisp，详见上面「驱动 AutoCAD」一节 |
| 成品 Task 入口（Excel / Word） | `ExcelHost/ExcelTask.fs`（`GetApplicationAsync pid` / `GetOrCreateApplicationAsync()`）、`WordHost/WordTask.fs`（`GetApplicationAsync info`） | 前置闸门同步判 → 排进 `Sta.runAsync` 固定 STA → `Result` 抛成 faulted Task；调用方直接 `let! app = …` |
| 老 .xls 逐页抄 | `ExcelHost/SheetCopy.fs`（`module SheetCopy`） | 一张一张 `Worksheet.Copy` 进副本，再删新格式带不走的东西；原件只读不改 |
| acad 的 Fact | `xUnitHost/AcadAttachTest.fs` + `AcadPdfTextTest.fs` + `AcadMgdStarTest.fs` | 离线 PID 闸门 + 真机（`*挂回 PID*` / `*acad 发内置命令*` / `*发一行 LISP*` / `*EndLisp*` / `*顶点*` / `*PDFSHXTEXT*`），真机全要 `XUNITHOST_ACAD_PID`，写图纸再加 `XUNITHOST_ACAD_TEXT=1` |
| ppt 升级函数 | `PowerPointHost/PptIO.fs` + `PptUpgrade.fs`；验证 `xUnitHost/PptUpgradeTest.fs`（`-m "*pptUpgrade*"`） | `pptUpgrade app <路径>`：`PptIO.detectMacro` 判宏 → 纯规则 `route` → `PptToPptx`(24) / `PptToPptm`(25)；转完读回幻灯片文字指纹、已有同名先转临时文件、判不出宏失败且原件保留。PowerPoint 的 Open/SaveAs 是**强类型按值**参数；`WithWindow=msoFalse` 不弹窗口；`Application.Visible` 不许设 msoFalse |
| 批量转格式 | `xUnitHost/ConvertCore.fs` + `Convert.fs`（**前部**）；`WordHost/DocIO.fs` + `DocUpgrade.fs`、`ExcelHost/BookIO.fs` + `XlsUpgrade.fs`（两个 host 库**末尾**） | 两侧对称：各一份文件放三个入口 —— 两条转换各自独立（格式码/后缀/核对/文案各归其主）+ 一个统一入口判宏 → 纯规则 `route` 选路；共用 COM 细节在各自的 `*IO.fs`；host 只回传 `(目标, 状态, 详情)`，`Convert.fs` 组装 Outcome、只留 `Paths.scan` / `Convert.all` / `readWordText` / `readBookPrint` / `canDelete`；真目录入口是两条按环境变量跑的 Fact（`ConvertTest.fs` 末段） |
| 黑板展示文本参考解析 | `xUnitHost/BlackboardText.fs`（测试项目前部） | `module internal BlackboardText`：`caption` / `ofText` 把插件文本读成 `BoardApp`（`Image` / `Pid` / `Rot`，**只有 `Pid` 参与挂回**），`docOfText` 另给 `BoardDoc`（只进报告）；**`OsHost` 不认文本、也没有「参数表」概念** |
| 挂回入口（四家同形） | `ExcelTask` / `AutoCADTask` / `WordTask` 的 `GetApplicationAsync pid`、`PowerPointHost.Host.GetApplication pid` | 前置闸门同步判（PID>0 且确实是对应进程）→ 排进固定 STA → 取 ROT 现成那只并核 pid / ROT 名后缀；**一律只认 pid** |
| 挂回闸门验证 | `xUnitHost/AttachGateTest.fs`（Word / PPT）+ `ExcelAttachTest.fs`（Excel） | `-method "*不合格*"` 一条（两家就地拒接）、`-method "*Excel 挂回*"` 四条；离线，不启任何 Office 进程 |
| 跨进程 | `xUnitHost/SecondProcessTest.fs` | 第二个进程 = 测试程序自己；地址 / 要说的话 / 回话文件都从环境变量来 |

**项目的增减由用户定**：往现有项目里**加文件随时可以**（新 Fact、新公用模块都行）；但
**不许私自新建项目**。现在的七个：`OsHost` / `WordHost` / `ExcelHost` / `PowerPointHost` /
`AutoCADHost` / `AcadMgdHost` / `xUnitHost`（`xUnitHost.slnx` 就是这七行）。再要开新项目，先问用户。

公用代码往上收（操作系统那类进 `OsHost`，测试杂活进 `Common.fs`），其余放测试项目前部。

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

## 黑板模式（插件抄现场 → 按 pid 挂回现成实例）

插件把「当前实例」抄到黑板上（采集时间 / PID / ROT 名 / HWND / EXE / 文档 / 路径），
WorkBuddy 只读写、**不负责黑板的创建与销毁**。

**四家挂回入口同形，参数只有一个进程号**：`ExcelHost.ExcelTask.GetApplicationAsync pid`、
`AutoCADHost.AutoCADTask.GetApplicationAsync pid`、`WordHost.WordTask.GetApplicationAsync pid`（都是 `Task<_>`）、
`PowerPointHost.Host.GetApplication pid`（同步 `Result<Application,string>`）。
前置闸门同步判（PID > 0，且确实是对应进程）→ 排进固定 STA → 从 ROT 取现成那只并核对：
Excel / AutoCAD / PowerPoint 用 `Hwnd → PID` 硬核；Word 用 `ActiveWindow` 上溯 `OpusApp`，
零文档时退到 ROT 名后缀（`!WordVSTO.Instance.<pid>`）。对不上一律拒绝，**绝不新建、绝不 Quit、不改应用状态**。

2026-11-13 起**没有 `GeneralAppInfo` 这层参数表**（连同 `verify` / `attachWith` / `defaultAcquire` / `GeneralDocInfo` 一并删除）：
当时以为要一张「通用参数表」才能挂回，实际上能唯一确定一只实例的只有 `Pid`；其余字段要么可重算（进程名 = 各库 `SelfImage`），
要么只作佐证（`Rot` 名、`ExePath`）——而 `Rot` 在真实调用里一直是空串，从没用上过。四家因此形状一致，
`OsHost` 也不再认识「参数表」这个概念。

进程映像名只留一个常量：`WordHost.Host.SelfImage = "WINWORD"`、`ExcelHost.Host.SelfImage = "EXCEL"`、
`PowerPointHost.Host.SelfImage = "POWERPNT"`；acad 的 `"acad"` 是 `AutoCADHost/Application.fs` 里的**私有**常量，
在外面要写就写字面量（`checkPid` 用它核实「确实是这个进程」）。

把插件展示文本读成字段的参考写法在测试项目 `xUnitHost/BlackboardText.fs`
（`ofText : string -> Result<BoardApp,string>`、`docOfText : string -> BoardDoc`，只供参考、不是库 API）；
插件改排版只改这一份，`OsHost` 里没有文本解析。
**文档级交接检查（换了文档 / 文档关了 / 采集后存过盘）按用户要求推迟**：`BoardDoc` 只是收下的数据，不参与判断。

验证：`-method "*黑板*"`（`BlackboardTest` 的 ROT 探针）、`-method "*不合格*"`（`AttachGateTest`：Word / PPT 闸门就地拒接）、
`-method "*Excel 挂回*"` 四条（含过时黑板信息按 PID 拒接）。
返回的 Application 与后续读写都放在 `OsHost.Sta.run` 那条固定 STA 线程上。

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
