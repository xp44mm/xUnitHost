# xUnitHost 项目长期笔记

## 铁律
任何在真机 Word / Excel / PowerPoint / AutoCAD 上干活，先加载 `.workbuddy/skills/office-host-collaboration/SKILL.md`（含 `references/api.md`、`references/gotchas.md`）。
顺序固定：**写 F# Fact → `dotnet build xUnitHost.slnx -v q`（0 警告 0 错误）→ 单跑 `-m "*关键字*" -showLiveOutput` → 才对真机执行 → 汇报**。代码留在项目里当记录，副作用靠它复现，不写一次性脚本。
汇报四段：做了什么 / 直接贴 Fact 的 `output` 报告（去 `[OUTPUT]` 前缀和类名，不另写一份）/ 留下的代码文件与 Fact 名 / 现场进程是否干净。

## 七个项目
`OsHost`（OS 层 + `Address.fs` 通用进程地址 + `ProcHost.fs` 起/关两条命令）、`WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost`（四家 COM host）、`AcadMgdHost`（accoreconsole 插件，netstandard2.0，只引 acdbmgd + accoremgd，不走 COM）、`xUnitHost`（xunit v3 测试）。
`ExcelCli` 曾获准新建，2026-11-06 16:09 用户删除，通用件上收到 OsHost。
**往现有项目加文件随时可以；不许私自新建项目**（新建须先说明理由问用户）。`xUnitHost.slnx` 现七个项目。

## 代码位置
| 东西 | 位置 |
|---|---|
| OS 层 / 进程地址 / 起关命令 | `OsHost/{OsHost.fs,Address.fs,ProcHost.fs}` |
| 四家 COM host（取实例/认自己/收尾） | `WordHost` / `ExcelHost` / `PowerPointHost` / `AutoCADHost` |
| 会话层（地址、指纹、两种模式） | `xUnitHost/Session.fs`（前部） |
| 报告 / 认人 / 忙重试 | `xUnitHost/Common.fs`：`module internal Report` / `Doc` / `Busy` |
| 批量转换底座 | `xUnitHost/ConvertCore.fs`（Item / Outcome / Paths），前部 |
| 批量分派 + 事后核对 | `xUnitHost/Convert.fs`（words / books / all / readWordText / readBookPrint / canDelete），前部 |
| 升级实现（各自 host 库末尾，形状对称） | `WordHost/{DocIO.fs,DocUpgrade.fs}`（DocToDocx / DocToDocm / DocUpgrade）、`ExcelHost/{BookIO.fs,XlsUpgrade.fs}`、`PowerPointHost/{PptIO.fs,PptUpgrade.fs}`（PptToPptx 24 / PptToPptm 25 / `pptUpgrade`，`convert` 别名） |
| 真目录入口 | `xUnitHost/ConvertTest.fs` 末尾两条 Fact，靠 `XUNITHOST_CONVERT_DIR` 叫起；不设就跳过 |
| 两条命令验证 | `CommandTest.fs`（`-m "*派发*"`，`module Dispatch` 注入 Excel 策略）、`ProcRx.fs` + `RxCommandTest.fs`（`-m "*Rx 两条命令*"`） |
| ppt 验证 | `PptUpgradeTest.fs`（`-m "*pptUpgrade*"` 两条） |
| 跨进程 | `SecondProcessTest.fs` |
| 管径 / 单价 / 技术方案回填 | `DnTest.fs` / `PriceTest.fs` / `DocSyncTest.fs`，真表靠 `XUNITHOST_DN_DIR` |
| accoreconsole 插件 | `AcadMgdHost/{Commands.fs,AssemblyInfo.fs,test/run.cmd}`；`XH_HELLO` / `XH_DB` + `Entry : IExtensionApplication` |

## 两种模式
- **独立（方式一）** `Host.execOnce` / `execOnceQuery`：一步做完、只读一个值；开一次、做完关掉。
- **合作（方式二）** `openSession` → `execByAddress` → `close`：多步、每步看上一步结果、跨进程交接。
- 地址 = `种类:pid:进程启动毫秒:最后保存毫秒:路径指纹:归属`；每步用 `exec` 返回的**新地址**（保存过就变，老地址作废）；`openSession` 可新建，`resolve` **绝不新建**；只关 `Owned=true` 的，用户手上的实例一个字不动。

## 黑板模式（2026-11-07）
- 场景：插件（WordVSTO / ExcelVSTO）把当前实例抄到黑板上（采集时间 / PID / ROT 名 `!WordVSTO.Instance.<pid>` / HWND / EXE / 文档 / 路径）；WorkBuddy 只按 info 挂回那一只实例读写，**不负责创建销毁**，不并行，交接要完整不过时。
- **`AppInfo` 是结构化参数表**（调用方按字段填报：`Kind/Pid/Rot/Hwnd/ExePath/CapturedMs/DocName/DocPath`），`OsHost/AppInfo.fs` 只有 `ofParts` / `describe` / `verify` / `ensureKind` / `checkDoc` / `attach`；**库不解析展示文本**，`ofText` 与 `Title` 已移出（2026-11-07 11:31 用户定：文本格式紧耦合，不入库）。参考实现 `BlackboardText.ofText` 在测试项目 `xUnitHost/BlackboardText.fs`，只供用户参考。`OsHost.Com.runningAt` 按 moniker 名字取 ROT（不枚举、不退回 GetActiveObject、不新建），`register` / `revoke` 供自测。
- 四家 `Host.GetApplication(info)` 形状对称（Word / Excel / PowerPoint / AutoCAD）：失败一律 `Error`，绝不新建 / Quit / 改 `DisplayAlerts`；返回的 Application 与后续读写都包 `OsHost.Sta.run`（固定长活 STA 线程）。
- 验证 `xUnitHost/BlackboardText.fs` + `BlackboardTest.fs` 五条 Fact：`-method "*黑板*"`；第一条验「手工填报参数表 == 参考解析结果」，真机那条设 `XUNITHOST_BLACKBOARD_FILE`。
- 项目根已 `git init`（初始提交 `54db328`），后续用 git 快照回退，不再用带时间戳的备份副本。

## 测试风格
- 一个类 + 构造函数拿 `ITestOutputHelper`；Fact 写作 `member this.``中文一句话``()`，不用模块级 `let`。
- 断言 `Should.equal 期望 实际`（`open FSharp.xUnit`），其余 `Assert.*`。
- 报告就是写进 `output` 的那份：`Report.title` → `line` / `line2` → `conclusion`，一行一个真实值，别写「成功」两个字；断言夹在报告行之间。
- 必须带 `-showLiveOutput`，否则 xunit 收走 output，只剩 Summary。时间用 `Report.clock`。
- 测试串行（`xunit.runner.json`：`parallelizeTestCollections: false`）。xunit v3 **不能用 `dotnet test`**，直接跑 `xUnitHost/bin/Debug/net10.0/xUnitHost.exe`。
- `-method` 的 filter **通配符只允许在开头/结尾**：`*真目录*清点*` 会报 `Filter expression ... is not valid`。真目录两条改用 `-method "*先清点*"` 和 `-method "*原件送回收站*"`。

## 命令
```bash
dotnet build xUnitHost.slnx -v q
bash .workbuddy/skills/office-host-collaboration/scripts/run-facts.sh -m "*关键字*"
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*关键字*" -noColor -showLiveOutput
```

## 转换规则（默认，不用问）
`.doc` 没宏→`.docx`、有宏→`.docm`；`.xls` 没宏→`.xlsx`、有宏→`.xlsm`；`.ppt` 没宏→`.pptx`、有宏→`.pptm`（保 VBA）。判不出宏就失败、原件保留。
六条转换各自独立（格式码 / 后缀 / 核对口径 / 文案各归其主）；判宏 → 纯规则 `route` 选路 → 调转换；host 只回传 `(目标, 状态, 详情)`，由 `Convert.fs` 组装 Outcome。
**转一份、读回比一份**（正文 / 格子值 / 幻灯片文字），不一致判失败、原件保留；只删 `canDelete`（「转了」/「已有·一致」）的原件，**删前先备份到目录外**。带宏的 `.docm`/`.xlsm` 是 ZIP，可查 `word/vbaProject.bin` / `xl/vbaProject.bin` 双重确认。
```bash
exe=xUnitHost/bin/Debug/net10.0/xUnitHost.exe
XUNITHOST_CONVERT_DIR='C:\目录' "$exe" -showLiveOutput -method "*真目录*清点*"
XUNITHOST_CONVERT_DIR='C:\目录' XUNITHOST_CONVERT_BATCH=10 "$exe" -showLiveOutput -method "*真目录*转*"
```
**2026-11-06 已接通 .ppt 批量**：`Paths.isCandidate/isPpt/targetOf/existingTargetOf` 三类全认；`Convert.shows` 用 `PowerPointHost.Host.execQuery` + `PptUpgrade.convert`，`Convert.all` = words @ books @ shows；`Convert.readPptPrint` 事后核对。验证 `ConvertTest.fs` 末尾 `ConvertBatchTest` 两条 Fact（`.ppt` 批量分派真机、备份/改名/回收站纯文件演练）。

**真目录闭环（2026-11-06 加固）**：`RealFolder` 依次做 ① 完整备份到**目标目录之外**（`<root>-bak<戳>`，保持目录结构、逐份比大小、可复用续跑）② 重名让位：现有同名新格式文件改名 `X-bak<戳>.docx`（`Paths.renameOf`），旧文件照常转，改名失败的旧文件跳过 ③ 分块转换（`XUNITHOST_CONVERT_CHUNK` 默认 200）④ 每份读回核对 ⑤ **只把核对通过的原件送回收站**（`Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile` + `RecycleOption.SendToRecycleBin`，不是 `File.Delete`），每批 `XUNITHOST_CONVERT_BATCH` 默认 10。备份目录可用 `XUNITHOST_CONVERT_BACKUP` 覆盖。失败的原件一律保留，Fact 不再断言 `bad.Length = 0`。

## 续跑前必改（2026-11-07 用户现场反馈，未实现）

`C:\崔胜利` 全量转换跑到 777/4907 时用户叫停，反馈三个会卡住、需要人工关弹窗的问题。**下次续跑前先改代码，别直接重跑**。

1. **Word「文档已损坏，是否修复」弹窗**：老 .doc 打开时 Word 弹修复对话框，COM 调用一直挂到有人去点掉。改 `WordHost/DocIO.fs` 的 `openDoc`：用 `Documents.OpenNoRepairDialog(...)` 取代 `Documents.Open`（这个方法专门抑制修复对话框），并设 `app.DisplayAlerts <- WdAlertLevel.wdAlertsNone`、`app.Options.ConfirmConversions <- false`；Excel 侧 `Workbooks.Open` 带 `CorruptLoad = XlCorruptLoad.xlNormalLoad`；PPT 侧保持 `Presentations.Open(..., WithWindow=msoFalse)` + `DisplayAlerts <- ppAlertsNone`。
2. **必须以不启用宏的方式打开**：带宏文件（`AutoOpen` / `Workbook_Open`）会弹宏提示，同样挂住。Word / Excel / PowerPoint 三家 host 在起实例后统一设 `app.AutomationSecurity <- MsoAutomationSecurity.msoAutomationSecurityForceDisable`（值 3），**另存完恢复原值**。判宏只靠 `HasVBProject`（bool，不执行宏）或 OLE 目录流，绝不靠运行宏；`route` 选 `.docm/.xlsm/.pptm` 的规则不变。
3. **按类型分三轮跑**：别再用 `Convert.all` 一轮混跑三类。给 `RealFolder` 加 `XUNITHOST_CONVERT_KIND=doc|xls|ppt`（`Paths.scan` 之后按后缀过滤），分三次调用。好处：某类弹窗/卡死不拖累另外两类、可按类调超时与实例重启策略、失败归因清楚。顺带给每份 COM 调用加超时，超时只 kill 自己起的 pid 再重启该实例，彻底免人工介入。

补充：1/2 的弹窗很可能就是本次「日志 22 分钟无输出、23 份失败」的成因之一，但**未经证实**，别当结论写。

## 约定与坑
- 用户平时开着 Word / Excel / acad（真图纸表格）：需要「从零起实例」的 Fact 开头 `requireClean "WINWORD"`，环境不干净 `Assert.Skip`，**绝不为跑通去关用户实例**；被跳过是正常现象。
- 别拉起用不到的程序：转换类任务只跑 `-m` 相关 Fact，别跑 `-a`；AutoCAD 默认不新起，要测显式 `XUNITHOST_LAUNCH_ACAD=1`。
- COM 差异：Word Open/SaveAs 是 `obj&`，Excel 是 `object`，**PowerPoint 是强类型按值**；PPT `Presentations.Open(path, ro, msoFalse, msoFalse)` 的 `WithWindow=msoFalse` 才不弹窗，`Application.Visible <- msoFalse` 不允许；`Presentation.Close()` 无参；`HasVBProject` 是 `bool`；`DisplayAlerts <- ppAlertsNone`。
- Quit 异步：`Gone=true` 后进程可能还有半拍，用 `Proc.waitGone` 复核；`finally` 只 kill 自己起的 pid。
- Rx 版：退出检测放 **close** 一侧、命令执行前订阅 `Process.Exited` + 1000ms 轮询兜底；放 start 一侧会误判。`Disposable.Create` 退订只摘 handler，绝不 kill。
- accoreconsole：`Console.Out` 被重定向到 `%TEMP%ccc*.tmp`，插件另按 `ACADMGDHOST_LOG` 写日志；netstandard2.0 能被 AutoCAD 2020 加载、CommandMethod 正常注册，前提是**不引 acmgd.dll / UI DLL**。
- 源文件 UTF-8 + BOM、CRLF；子进程输出中文先设 `Console.OutputEncoding <- UTF8`。**例外**：`AcadMgdHost/test/run.cmd` 纯 ASCII / CRLF / **无 BOM**（cmd.exe 把 BOM 当命令），且 rem 行不能出现 `%~fI` 这类参数运算符。
- 递归扫跳过 `.` 开头 / `.idx_cache` 缓存目录和 `~$` 锁文件。
