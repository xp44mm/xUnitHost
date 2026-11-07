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
- **14:11 收束，只此一条路**：用户把 `AppKind` / `AppInfo` 标为 Obsolete 后整份删除 `OsHost/AppInfo.fs`（含 `verify`/`ensureKind`/`attach`），范围选「原来的 GetApplication 被 GeneralAppInfo 参数的替代」。Obsolete 标注快照 `19d7d0c`，旧实现在更早 git 历史。
- `OsHost/GeneralAppInfo.fs`：`GeneralIdentity`（Image/Name/Tag 三参数）+ `GeneralAppInfo`（`create`/`describe()`/`verify`/`ensureTag`/`defaultAcquire`/`attachWith`/`attach`）；库不认 Office 类型、不解析文本、不新建不 Quit。`OsHost.Com.runningAt` 按 moniker 名字取 ROT（不枚举、不退回 GetActiveObject、不新建），`register`/`revoke` 供自测。
- **kind 里有用的代码搬到专用 Host**：四家各有 `Host.SelfIdentity`（Word `WINWORD/Word/word`、Excel `EXCEL/Excel/excel`、PPT `POWERPNT/PowerPoint/ppt`、CAD `acad/AutoCAD/acad`）；四家 `Host.GetApplication(info: GeneralAppInfo)` 走 `GeneralAppInfo.attach info SelfIdentity.Tag progIds work`（强转 + pid 核对），形状对称、失败一律 Error、绝不新建/Quit/改 DisplayAlerts；Excel 14:05 的 `Host.General` 子模块已撤。返回 Application 与后续读写包 `OsHost.Sta.run`。
- 测试：`xUnitHost/BlackboardText.fs` 参考解析产出 `GeneralAppInfo`（身份取各家 SelfIdentity，不是库 API）；`-method "*黑板*"` 4 条、`-method "*通用参数表*"` 5 条（GeneralAppInfoTest 3 + ExcelGeneralTest 2）、`-method "*Excel 通用参数表*"` 2 条。真机那条设 `XUNITHOST_BLACKBOARD_FILE`。`GeneralBridge` 与旧对照 Fact 已删，第三条改成「四家 SelfIdentity + 入口拒接」。
- **文档级交接检查推迟（13:51 用户定）**：`checkDoc` 两路都删、文档级 Fact 删、四家入口只做强转 + pid 核对；`DocName`/`DocPath` 只收数据、不参与判断。旧实现在 `6be1725` 之前。
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

## P0 安全开关（2026-11-07 用户反馈；15:2x 已实现 1/2 与分轮，超时看门狗未做）

`C:\崔胜利` 全量转换跑到 777/4907 时用户叫停，反馈三个会卡住、需要人工关弹窗的问题。
前两条和「按类型分三轮」已于 15:2x 实现并验证；**kill 型超时看门狗未做**（理由见本节末尾）。

1. **Word「文档已损坏，是否修复」弹窗**：老 .doc 打开时 Word 弹修复对话框，COM 调用一直挂到有人去点掉。改 `WordHost/DocIO.fs` 的 `openDoc`：用 `Documents.OpenNoRepairDialog(...)` 取代 `Documents.Open`（这个方法专门抑制修复对话框），并设 `app.DisplayAlerts <- WdAlertLevel.wdAlertsNone`、`app.Options.ConfirmConversions <- false`；Excel 侧 `Workbooks.Open` 带 `CorruptLoad = XlCorruptLoad.xlNormalLoad`；PPT 侧保持 `Presentations.Open(..., WithWindow=msoFalse)` + `DisplayAlerts <- ppAlertsNone`。
2. **必须以不启用宏的方式打开**：带宏文件（`AutoOpen` / `Workbook_Open`）会弹宏提示，同样挂住。Word / Excel / PowerPoint 三家 host 在起实例后统一设 `app.AutomationSecurity <- MsoAutomationSecurity.msoAutomationSecurityForceDisable`（值 3），**另存完恢复原值**。判宏只靠 `HasVBProject`（bool，不执行宏）或 OLE 目录流，绝不靠运行宏；`route` 选 `.docm/.xlsm/.pptm` 的规则不变。
3. **按类型分三轮跑**：别再用 `Convert.all` 一轮混跑三类。给 `RealFolder` 加 `XUNITHOST_CONVERT_KIND=doc|xls|ppt`（`Paths.scan` 之后按后缀过滤），分三次调用。好处：某类弹窗/卡死不拖累另外两类、可按类调超时与实例重启策略、失败归因清楚。顺带给每份 COM 调用加超时，超时只 kill 自己起的 pid 再重启该实例，彻底免人工介入。

补充：1/2 的弹窗很可能就是本次「日志 22 分钟无输出、23 份失败」的成因之一，但**未经证实**，别当结论写。

**实现记录（2026-11-07 15:2x）**
- `WordHost/DocIO.fs` → `Documents.OpenNoRepairDialog`；`WordHost/Host.fs` 的 quieten 加 `Options.ConfirmConversions=false`。
- `ExcelHost/BookIO.fs` → `Workbooks.Open` 传满 15 参、末位 `CorruptLoad=xlNormalLoad`；quieten 加 `AskToUpdateLinks=false`。
- 三家 `Host.fs` 各加公开 `module Security`（`apply` / `restore MsoAutomationSecurity`），quieten 统一设 `msoAutomationSecurityForceDisable`(3)，`runOnce` 在 work 之后 `Security.restore app previousSecurity`。判宏仍只用 `HasVBProject`，`route` 规则不变；PPT 的 `WithWindow=msoFalse` 保持。
- `ConvertCore.fs` 的 `Paths` 加 `Kind` / `parseKind` / `kindFromEnv` / `kindLabel` / `filterByKind`；`ConvertTest.fs` 的 `RealFolder.items` 按 `XUNITHOST_CONVERT_KIND=doc|xls|ppt` 过滤，两条真目录 Fact 报告加「本轮类型」行，备份/改名/回收站只作用于本轮那一类。
- 新 Fact 文件 `xUnitHost/ConvertSafetyTest.fs` 四条（`-method "*安全开关*"`）：三家真机各读回 DisplayAlerts / AutomationSecurity / 恢复值 + KIND 纯逻辑。编译 0 警告 0 错误；`*转换*` 4/4、`*pptUpgrade*` 2/2、真目录 `KIND=doc` 清点 2803 份且一个文件没动；现场 Office 残留 0。
- 续跑分三轮（建议 ppt → xls → doc）：
  `XUNITHOST_CONVERT_DIR='C:\崔胜利' XUNITHOST_CONVERT_KIND=ppt XUNITHOST_CONVERT_CHUNK=20 XUNITHOST_CONVERT_BATCH=10 xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*原件送回收站*" -noColor -showLiveOutput`
- **未做**：每份 COM 调用超时 + kill 自己起的 pid 重启实例。COM 跑在固定 STA 线程，按份中断要改 `Host` 暴露本次 pid 并配套重试/重启，风险高于收益；弹窗根因（1/2）已消除，先用 `CHUNK=20` 缩小单块影响面。要做再单开一轮。

## 常驻 Excel 挂回式通道（2026-11-07 15:3x，验一张未通过）

- 用户改架构：Excel 由 ExcelVSTO 插件常驻（`!ExcelVSTO.Instance.<pid>`），本侧只用 `ExcelHost.Host.GetApplication(info: GeneralAppInfo)` 挂回，**不新建 / 不 Quit / 不 kill / 不 requireClean**；只逐份开→升级→保存→关自己那本。范围：只 .xls 走常驻。
- 代码：`Convert.fs` 的 `booksAttached app items`（**绝不用 `CloseAllWorkbooksWithoutSaving`**，那会关掉用户所有表）；`ConvertTest.fs` 的 `RealFolder` 改 internal；新增 `xUnitHost/AttachedConvertTest.fs` 两条 Fact（`-method "*只读核对*"` / `*挂回式升级一张*`）；黑板文本 `TestResults/崔胜利-Excel黑板.txt` 靠 `XUNITHOST_BLACKBOARD_FILE`。所有 COM 调用包 `OsHost.Sta.run`。
- 只读挂回 1/1 通过：按 ROT 精确挂回 pid 30816、`Ok System.__ComObject`、工作簿数 1、原值 DisplayAlerts=True / AutomationSecurity=1、现场进程前后 `[30816]`。
- **验一张 FAIL，三个坑（未修，等用户定）**：
  1. 常驻实例里唯一的未保存空「工作簿1」在第一次 `Workbooks.Open` 时被 Excel 关掉 → 工作簿数 1→0，窗口 HWND `0x5f0594`→`0x1b04bc`，**黑板信息随即过时**，`verify` 拒接「HWND 已不属于 PID」。翻倍法会在第一轮后断链。
  2. 老 .xls 被打开后**原件被改写**（683520→684032 B、md5 变了）。目录外备份与 `TestResults/验一张-转换前原件.xls` 都保住了转换前的 683520 字节，可还原。
  3. `.xlsx` 没生成（转换失败），但**失败原因没打印**：`Should.equal booksBefore booksAfter` 排在逐份 outcome 报告之前，Fact 先 FAIL。下次把报告挪到断言前。
- 断言在 `deleteVerified` 之前 → **原件一份都没进回收站**；目录残留仍是 xls 1237 / doc 2803 / ppt 90。现场 EXCEL.EXE pid 30816 仍在（工作簿数 0），WINWORD / POWERPNT / xUnitHost 全 0。
- **15:5x 用户拍板（只答第 ① 项）**：「你只管你自己的 excel 文件，其他文件的 hwnd 你不校验，你只是借用 excel 进程」→ **不重抄黑板、插件也不用每轮重抄**；挂回只认进程身份，HWND 是别人文档窗口的句柄，不参与判断。

**借用进程口径（已实现并验证，2026-11-07 15:5x）**
- `OsHost/GeneralAppInfo.fs` 加公开 `borrowProcess (info) = { info with Hwnd = 0n }`：verify 本来就把 `Hwnd=0n` 定义为「没填窗口、不校验」；PID 存活 / 采集时间 / 启动时刻 / EXE / ROT 名后缀一条不少。**库不改校验规则**，只给「借用进程」一个正式名字；原表 HWND 留在调用方当数据，报告照实打印。
- `AttachedConvertTest.fs`：`Attached.borrow = GeneralAppInfo.borrowProcess`；两条挂回 Fact 都用借用表；新增 `ownBooksLeft`（按 FullName 只数自己本轮打开的文件）。**去掉 `Should.equal booksBefore booksAfter`**：用户开着几本工作簿不归这条通道管，不再因 Excel 关掉用户空表而 FAIL；改成只断言「自己开的都关了」+「应用开关恢复原值」+「常驻 pid 还在、进程没多没少」。逐份 outcome 报告挪到所有断言之前，失败原因不会再被断言挡住。
- 新 Fact `-method "*借用进程*"`（纯逻辑）：换了 HWND 的旧表被拒接、借用表放行；PID 不在 / ROT 后缀不符 / EXE 不符 / 采集时间缺失 / Image 为空一律仍拒接。
- 验证：`dotnet build` 0 警告 0 错误；`*借用进程*` 1/1、`*只读核对*` 1/1（**HWND 已过时的旧黑板现在能挂回 pid 30816**）、`*通用参数表*` 5/5、`*黑板*` 4 条（3 通过 1 Skip）。`-method "*黑板*"` 曾误捞新 Fact（名字带「黑板」），已改名 `*只读核对*`，避免污染既有回归口径。
- 本轮**磁盘一个文件没动**：原件仍 684032 B（上次被 Excel 改写）、备份/留档 683520 B、`.xlsx/.xlsm` 都未生成、目录残留 4130 份。git 快照 `16d3080`（改动前）。
- **仍未处理（用户未答）**：② 原件被 Excel 改写要不要还原成 683520 B；③ 只读打开 / 转副本避免改写原件；④ 重跑验一张查清 `.xlsx` 为什么没生成（报告前置已就绪，重跑就能看到 Detail）。

## 约定与坑
- 用户平时开着 Word / Excel / acad（真图纸表格）：需要「从零起实例」的 Fact 开头 `requireClean "WINWORD"`，环境不干净 `Assert.Skip`，**绝不为跑通去关用户实例**；被跳过是正常现象。
- 别拉起用不到的程序：转换类任务只跑 `-m` 相关 Fact，别跑 `-a`；AutoCAD 默认不新起，要测显式 `XUNITHOST_LAUNCH_ACAD=1`。
- COM 差异：Word Open/SaveAs 是 `obj&`，Excel 是 `object`，**PowerPoint 是强类型按值**；PPT `Presentations.Open(path, ro, msoFalse, msoFalse)` 的 `WithWindow=msoFalse` 才不弹窗，`Application.Visible <- msoFalse` 不允许；`Presentation.Close()` 无参；`HasVBProject` 是 `bool`；`DisplayAlerts <- ppAlertsNone`。
- Quit 异步：`Gone=true` 后进程可能还有半拍，用 `Proc.waitGone` 复核；`finally` 只 kill 自己起的 pid。
- Rx 版：退出检测放 **close** 一侧、命令执行前订阅 `Process.Exited` + 1000ms 轮询兜底；放 start 一侧会误判。`Disposable.Create` 退订只摘 handler，绝不 kill。
- accoreconsole：`Console.Out` 被重定向到 `%TEMP%ccc*.tmp`，插件另按 `ACADMGDHOST_LOG` 写日志；netstandard2.0 能被 AutoCAD 2020 加载、CommandMethod 正常注册，前提是**不引 acmgd.dll / UI DLL**。
- 源文件 UTF-8 + BOM、CRLF；子进程输出中文先设 `Console.OutputEncoding <- UTF8`。**例外**：`AcadMgdHost/test/run.cmd` 纯 ASCII / CRLF / **无 BOM**（cmd.exe 把 BOM 当命令），且 rem 行不能出现 `%~fI` 这类参数运算符。
- 递归扫跳过 `.` 开头 / `.idx_cache` 缓存目录和 `~$` 锁文件。
