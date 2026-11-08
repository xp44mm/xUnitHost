# xUnitHost 项目长期笔记

## 铁律
真机 Word / Excel / PowerPoint / AutoCAD 上干活，先读 `.workbuddy/skills/office-host-collaboration/SKILL.md`（含 `references/api.md`、`references/gotchas.md`）。
顺序固定：**写 F# Fact → `dotnet build xUnitHost.slnx -v q`（0 警告 0 错误）→ 单跑 `-m "*关键字*" -showLiveOutput` → 才动真机 → 汇报**。代码留项目当记录，副作用靠它复现，不写一次性脚本。
汇报四段：做了什么 / 直接贴 Fact 的 `output`（去 `[OUTPUT]` 前缀与类名）/ 留下的代码与 Fact 名 / 现场进程是否干净。

## 七个项目
`OsHost`（OS 层 + `Address.fs` 进程地址 + `ProcHost.fs` 起关命令）、`WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost`（四家 COM host）、`AcadMgdHost`（accoreconsole 插件，netstandard2.0，只引 acdbmgd + accoremgd，不走 COM）、`xUnitHost`（xunit v3 测试）。
**往现有项目加文件随时可以；不许私自新建项目**（须先说明理由问用户）。

## 代码位置
| 东西 | 位置 |
|---|---|
| OS 层 / 进程地址 / 起关命令 | `OsHost/{OsHost.fs,Address.fs,ProcHost.fs,GeneralAppInfo.fs}` |
| 四家 COM host（取实例/认自己/收尾/`Security`） | `WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost` |
| 会话层（地址、指纹、两种模式） | `xUnitHost/Session.fs`（前部） |
| 报告 / 认人 / 忙重试 | `xUnitHost/Common.fs`：`module internal Report` / `Doc` / `Busy` |
| 批量转换底座 | `xUnitHost/ConvertCore.fs`（Item / Outcome / Paths），前部 |
| 批量分派 + 事后核对 | `xUnitHost/Convert.fs`（words / books / shows / all / readWordText / readBookPrint / readPptPrint / canDelete / **booksAttached**），前部 |
| 升级实现（各自 host 库末尾，形状对称） | `WordHost/{DocIO.fs,DocUpgrade.fs}`、`ExcelHost/{BookIO.fs,XlsUpgrade.fs}`、`PowerPointHost/{PptIO.fs,PptUpgrade.fs}`（`pptUpgrade`，`convert` 别名） |
| 真目录入口 | `xUnitHost/ConvertTest.fs` 末尾两条 Fact（`module internal RealFolder`），靠 `XUNITHOST_CONVERT_DIR` 叫起 |
| 常驻 Excel 挂回 + 逐份推进 | `xUnitHost/AttachedConvertTest.fs`（`module private Attached` + 3 条 Fact），排在 ConvertTest 之后 |
| P0 安全开关验证 | `xUnitHost/ConvertSafetyTest.fs`（`-m "*安全开关*"`） |
| 其它验证 | `CommandTest.fs`（`*派发*`）、`ProcRx.fs` + `RxCommandTest.fs`（`*Rx 两条命令*`）、`PptUpgradeTest.fs`（`*pptUpgrade*`）、`SecondProcessTest.fs`、`DnTest.fs` / `PriceTest.fs` / `DocSyncTest.fs`（真表靠 `XUNITHOST_DN_DIR`） |
| accoreconsole 插件 | `AcadMgdHost/{Commands.fs,AssemblyInfo.fs,test/run.cmd}`；`XH_HELLO` / `XH_DB` |

## 两种模式
- **独立（方式一）** `Host.execOnce` / `execOnceQuery`：一步做完、只读一个值；开一次、做完关掉。
- **合作（方式二）** `openSession` → `execByAddress` → `close`：多步、每步看上一步结果、跨进程交接。
- 地址 = `种类:pid:进程启动毫秒:最后保存毫秒:路径指纹:归属`；每步用 `exec` 返回的**新地址**；`openSession` 可新建，`resolve` **绝不新建**；只关 `Owned=true` 的，用户手上的实例一个字不动。

## 黑板模式
- 插件（WordVSTO / ExcelVSTO）把当前实例抄到黑板（采集时间 / PID / ROT 名 `!ExcelVSTO.Instance.<pid>` / HWND / EXE / 文档）；本侧只按 info 挂回那一只实例读写，**不创建不销毁**。黑板文本里的采集时间行 / HWND 行都不解析：不是进程的属性。
- `OsHost/GeneralAppInfo.fs`：`GeneralAppInfo`（**只 4 项** `Image`/`Pid`/`Rot`/`ExePath`；`create`/`describe`/`verify`/`defaultAcquire`/`attachWith`/`attach`）+ `GeneralDocInfo`（只作报告数据）；库不认 Office 类型、不解析文本、不新建不 Quit。`OsHost.Com.runningAt` 按 moniker 名取 ROT（不枚举、不退回 GetActiveObject、不新建）。
- 四家 `Host.SelfImage`（Word `WINWORD`、Excel `EXCEL`、PPT `POWERPNT`、CAD `acad`）+ `Host.GetApplication(info)` 走 `GeneralAppInfo.attach`（强转 + pid 核对），失败一律 Error，绝不新建/Quit/改显示开关。
- **`GeneralIdentity` 已整份删除**（11-08，`b567007`）：用户口径「只用来识别进程」→ 表里只留找到进程所需的现场值，`Name`（显示名）/ `Tag`（种类标记）与 `ensureTag` 一起撤掉；进错入口改由 work 强转失败或 pid 不符兜住（也是明确 Error）。
- **文档级交接检查已取消**（13:51）：`checkDoc` 两路都删，`DocName`/`DocPath` 只收数据不参与判断。
- **`GeneralAppInfo` 只装认进程的字段**（11-08，`0dba505`）：文档名/路径搬进同文件的新记录 `GeneralDocInfo`（`Name`/`Path`/`create`/`empty`/`describe`），纯报告数据。文本源用 `BlackboardText.docOfText` 另取。
- **`Hwnd` 与 `borrowProcess` 已删**（11-09，`00cb85a`）：窗口句柄属「谁当时开了什么」，借用进程后必换，留着是摆设 → 表里与 `verify` 里都撤掉。文本里的 `HWND` 行不再解析。`Attached.borrow` 变成恒等函数（保留调用点，名字仍表「借用那只常驻进程」）。
- **`CapturedMs` 已删**（11-09，`349f915`）：采集时刻是黑板抄录动作的时间，不是进程属性 → 表里撤掉，连带 `verify` 的「缺采集时间」「进程号被复用（靠 Proc.startedMs 比采集时刻）」两条检查、`BlackboardText.parseClock` 一起删；`create` 收 4 个参数。插件文本里那行「采集时间」照印但不解析（Fact 用「抽掉这行结果不变」验证）。`verify` 现在只比 PID 有效 / 进程名非空 / PID 存活 / EXE / ROT 名后缀。
- 旧 `AppInfo.fs`（`AppKind`）已 Obsolete 后整份删除，旧实现在 `19d7d0c` 之前；`GeneralBridge` 与旧对照 Fact 已删。

## 测试风格
- 一个类 + 构造函数拿 `ITestOutputHelper`；Fact 写 `member this.``中文一句话``()`，不用模块级 `let`。
- 断言 `Should.equal 期望 实际`（`open FSharp.xUnit`），其余 `Assert.*`。
- 报告就是写进 `output` 的那份：`Report.title` → `line` → `conclusion`，一行一个真实值，别写「成功」两个字；断言夹在报告行之间。
- 必须带 `-showLiveOutput`，否则 xunit 收走 output。测试串行（`xunit.runner.json`）。xunit v3 **不能用 `dotnet test`**，直接跑 `xUnitHost/bin/Debug/net10.0/xUnitHost.exe`。
- `-method` 通配符**只允许在开头/结尾**：`*真目录*清点*` 报 `Filter expression ... is not valid`。

## 命令
```bash
dotnet build xUnitHost.slnx -v q
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*关键字*" -noColor -showLiveOutput
```

## 转换规则（默认，不用问）
`.doc` 无宏→`.docx`、有宏→`.docm`；`.xls` 无宏→`.xlsx`、有宏→`.xlsm`；`.ppt` 无宏→`.pptx`、有宏→`.pptm`（保 VBA）。判不出宏就失败、原件保留。判宏只用 `HasVBProject`（bool，不执行宏）或 OLE/ZIP 目录流（`word/vbaProject.bin`、`xl/vbaProject.bin`），绝不运行宏。
**转一份、读回比一份**（正文 / 格子值 / 幻灯片文字），不一致判失败、原件保留；只删 `canDelete`（「转了」/「已有·一致」）的原件，删前先备份到目录外。
真目录闭环（`RealFolder`）：① 完整备份到**目标目录之外**（`<root>-bak<戳>`，可复用续跑）② 重名让位 `X-bak<戳>.docx` ③ 分块转换（`CHUNK` 默认 200）④ 每份读回核对 ⑤ 只把核对通过的原件送**回收站**（`Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile` + `RecycleOption.SendToRecycleBin`，不用 `File.Delete`），`BATCH` 默认 10。失败的原件一律保留。

## P0 安全开关（2026-11-07）
- Word 用 `Documents.OpenNoRepairDialog`（抑制修复框）+ `ConfirmConversions=false`；Excel `Workbooks.Open` 传满参数、末位 `CorruptLoad=xlNormalLoad`、`AskToUpdateLinks=false`；PPT 保持 `Presentations.Open(..., WithWindow=msoFalse)`。
- 三家 `Host.fs` 各公开 `module Security`（`apply`/`restore`），quieten 设 `msoAutomationSecurityForceDisable`(3)，用完恢复原值。
- 按类型分轮：`XUNITHOST_CONVERT_KIND=doc|xls|ppt`（`Paths.Kind` / `filterByKind`），别一轮混跑三类。
- **kill 型 COM 超时看门狗未做**：COM 跑固定 STA，按份中断要 `Host` 暴露本次 pid 并配重试/重启，风险高于收益；改用 `CHUNK=20` 缩影响面 + 逐份推进。

## 常驻 Excel 挂回式通道（.xls 只走这条）
- 用户架构：Excel 由 ExcelVSTO 插件常驻，本侧只用 `ExcelHost.Host.GetApplication(info)` 挂回，**不新建 / 不 Quit / 不 kill / 不 requireClean**；只逐份开→升级→保存→关自己那本（`Convert.booksAttached`，**绝不用 `CloseAllWorkbooksWithoutSaving`**）。
- **借用进程口径**：表里压根没有窗口句柄 / 采集时刻这些「谁当时开着什么、什么时候抄的」字段，参数表原样交给 `GetApplication` 即可（11-09 起 `Attached.borrow` 只是恒等函数，保留调用点表意）。用户话：「你只管你自己的 excel 文件，其他文件的 hwnd 你不校验，你只是借用 excel 进程」。别断言用户工作簿数量。校验只核 进程名 / PID 存活 / EXE / ROT 名后缀。
- 三个实战坑：① Excel 打开老 .xls 会**改写原件字节**（683520→684032，md5 变）→ 每份失败/送走前从备份还原；② 常驻实例被反复挂回会劣化（曾卡 123 秒），**卡住就重启 Excel 再来**，同一份在新实例上 14 秒通过；③ 老 .xls 若是共享工作簿/带修订历史，SaveAs 被拒（报错原文为准）。
- 逐份推进器 `-m "*逐份升级*"`：一份一送，成功即把原件送回收站；失败打印归类原因、记 `逐份-失败清单.txt`（可直接当 `XUNITHOST_CONVERT_SKIP`）、接着下一份；同一原因连续 N 次或常驻 pid 消失才停。每份计时写 `逐份-进度.log`。开跑前只关「本次队列里」上次残留的工作簿。
- **弹窗策略（2026-11-07 17:3x 用户定）**：不再指望躲掉所有对话框 —— 让界面露出来，卡住就把 Excel 切到前台催人处理。实现：`Attached.bringToFront pid`（`Microsoft.VisualBasic.Interaction.AppActivate`，只激活窗口，不杀进程/不动文件）；默认 45 秒无进展就切前台并写提示，之后每 45 秒再催（`XUNITHOST_CONVERT_POPUP_HINT`，旧名 `XUNITHOST_CONVERT_TIMEOUT` 仍认）；开跑时前置一次（`XUNITHOST_CONVERT_FOCUS=0` 关）；报告打印「催前台 N 次」。**卡住不判失败**，等用户点掉弹窗后 COM 自己返回。
- `C:\崔胜利` 战果：`.doc/.ppt` 全量已转；`.xls` 从 1215 份推进到**剩 11 份**（另 6 个 `.idx_cache/src/*.xls` 属缓存，本就不转）。11 份失败分五类：共享工作簿 1、宏判定 2（`0x800A03EC` / Excel 4.0 宏表）、内容指纹不匹配 6（新文件字符一律更多，疑似公式重算，未证实）、`0x800A03EC` 1、被占用 1。全部原件保留、未删未改，等用户决定 A/B/C/D。

## 约定与坑
- 用户平时开着 Word / Excel / acad：需要从零起实例的 Fact 开头 `requireClean`，不干净就 `Assert.Skip`，**绝不为跑通去关用户实例**。
- 别拉起用不到的程序：只跑 `-m` 相关 Fact，别跑 `-a`；AutoCAD 显式 `XUNITHOST_LAUNCH_ACAD=1` 才起。
- COM 差异：Word Open/SaveAs 是 `obj&`，Excel 是 `object`，**PowerPoint 是强类型按值**；`Presentation.Close()` 无参；`HasVBProject` 是 `bool`；`DisplayAlerts <- ppAlertsNone`。
- Quit 异步：`Gone=true` 后进程可能还有半拍，用 `Proc.waitGone` 复核；`finally` 只 kill 自己起的 pid。
- Rx 版：退出检测放 **close** 一侧，命令执行前订阅 `Process.Exited` + 1000ms 轮询兜底。
- accoreconsole：`Console.Out` 被重定向到 `%TEMP%ccc*.tmp`，插件另按 `ACADMGDHOST_LOG` 写日志。
- 源文件 UTF-8 + BOM、CRLF；子进程输出中文先设 `Console.OutputEncoding <- UTF8`。例外：`AcadMgdHost/test/run.cmd` 纯 ASCII / CRLF / **无 BOM**。
- 递归扫跳过 `.` 开头 / `.idx_cache` 缓存目录 / `~$` 锁文件。
- 项目根已 `git init`（`54db328`），用 git 快照回退，不再用带时间戳的备份副本。
