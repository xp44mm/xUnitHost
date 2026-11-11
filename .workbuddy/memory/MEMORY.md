# xUnitHost 项目长期笔记

## 铁律
- 真机 Word/Excel/PPT/acad 干活前，先读 `.workbuddy/skills/office-host-collaboration/SKILL.md`（含 `references/api.md`、`gotchas.md`）。
- **acad 驱动只用 `AutoCADHost.Task` 的现成函数**（命令等 EndCommand、LISP 等 EndLisp、按 PID 挂回），不许手拼 `SendCommand` / 自订阅事件。
- 顺序：**写 Fact → `dotnet build xUnitHost.slnx -v q`（0 警告 0 错误）→ 单跑 `-m "*关键字*" -showLiveOutput` → 才动真机 → 汇报四段**（做了什么 / 贴 output / 代码与 Fact 名 / 现场进程是否干净）。代码留项目当记录。
- 用户平时开着 Office/acad：只跑 `-m`，别跑 `-a`；从零起实例的 Fact 用 `requireClean`，不干净就 `Assert.Skip`，**绝不为跑通去关用户实例**。

## 项目与命令
`OsHost`（OS 层 / `Address` / `ProcHost` / `GeneralAppInfo` / `TextFile`）、`WordHost`、`ExcelHost`、`PowerPointHost`、`AutoCADHost`、`AcadMgdHost`、`xUnitHost`（xunit v3，**直接跑 exe，不能用 `dotnet test`**）。**不许私自新建项目**。

```bash
dotnet build xUnitHost.slnx -v q
xUnitHost/bin/Debug/net10.0/xUnitHost.exe -method "*关键字*" -noColor -showLiveOutput   # 通配符只许在首尾
```

## 四个挂回入口（只认这些）
| 程序 | 入口 | 闸门 / 注意 |
|---|---|---|
| Word | `WordHost/WordTask.fs` → `WordTask.GetApplicationAsync info`（Task 版，2026-11-11 从 `Host.GetApplication` 挪来，**唯一**） | `GeneralAppInfo.verify` 纯查进程表、同步 faulted；`pidOf` 要 `ActiveWindow` 上溯 `OpusApp`，没开文档返回 None（不反驳）。Fact `-m "*通用参数表*"`，真机 `XUNITHOST_WORD_PID` |
| Excel | `ExcelHost/ExcelTask.fs` → `ExcelTask.GetApplicationAsync pid`（**唯一**，info 版 11-10 已删） | `Application.Hwnd`→PID；Fact `ExcelAttachTest.fs`，真机 `XUNITHOST_EXCEL_PID` |
| PPT | `PowerPointHost.Host.GetApplication info`（`: Result<_,_>`） | 只做方式一，无会话层 |
| acad | `AutoCADHost/Task.fs`（`module AutoCADHost.Task`）：`GetApplicationAsync pid` / `commandAsync` / `sendCommandAsync` / `modalAsync` / `openDrawingAsync` / `sendLispAsync`(不等) / `sendLispAndWaitAsync`·`runLispAsync`·`runLispFileAsync`(等 EndLisp) | 参数收 `IAcadApplication`/`IAcadDocument`；`EventPair` 不认 COM，`AcadBeginEnd` 只剩接线；**`Ready.fs` 已废弃**（IsQuiescent 挡不住 RPC_E_CALL_REJECTED） |

- 判定链 `Host.checkPid / findOnSta / unwrap`、`matchApp` 已公开，留在各 `Host.fs`。
- 黑板（VSTO 抄 PID/ROT/EXE/文档）只挂回那一只，不新建不 Quit。**黑板常量写死在测试文件里**（`BlackboardText.ofText`），用户换进程/重开文件必须同步改（11-10 遇过 28884→14052）。
- 独立 `Host.exec*`（要求本机没有同类进程）vs 合作 `openSession`→`execByAddress`→`close`；每步用新地址，绝不自拼；只有 `Owned=true` 才真关。
- 改用户**已打开的真表/真文档**：备份到项目目录外 → 逐行核对现场 → 只改点名处 → `Save()` → 读回 → 另用 zip+XML 独立核磁盘字节。

## 测试风格
一个类 + 构造函数拿 `ITestOutputHelper`，Fact 是成员方法 `member this.``中文一句话``()`，纯函数放模块里且在类之前。报告写进 `output`：`Report.title` → `line` / `line2 output 标签 子标签 内容` → `conclusion`，一行一个真实值。断言 `Should.equal 期望 实际`（`FSharp.xUnit`）。必须 `-showLiveOutput`。

## Excel / XLL 的坑
- **`Range.Formula` 会给自定义函数补 `@`，`Formula2` 不补**；动态数组必须 `Formula2` 才能 spill。批量写公式一律 `Formula2`。
- XLL 名：`Evaluate`/公式栏不带前缀，磁盘 XML 才存 `_xll.`；数组函数 = 单锚点 + `<f t="array" ref="C7:C9">`。
- `-2146826259` / `#NAME?` 往往只是**那个 Excel 进程没加载插件**：先 `Evaluate` 探测。
- 起 Excel 用 `ExcelHost.Host.GetOrCreateApplication()`（会把 XLL 一起加载），别自己 `Com.create` 隐藏实例；拿到 app 先 `Evaluate` 探 `空气体积成分` / `gasMoleEnthalpy`。
- **`0x800A03EC` / 静默 `#NAME?` 常见于以 `=` 开头的说明文本**（`Formula` 与 `Value2` 都一样，11-11 D11 踩过）。写完扫一遍 `Range(addr).Text` 有无 `#` 错值（整块取会抛，逐格取）。
- 批量写：先收集 (地址,值) → 校验现有非空地址都在清单内 → 统一落笔；spill 锚点写前 `ClearContents`。
- **同会话起不出两本同名工作簿**：`Workbooks.Open` 第二个同名文件静默失败（返回 null、`Worksheets.Count` 为 0）。先 `File.Copy` 成带来源标记的副本（`BurnerCatalogTest.Item.Alias`）。
- `Workbooks.Add()` 偶发 `0x800AC472` → 局部重试；加表先 `Worksheets.Add()` 再 `Move(Missing, 目标表)`，`Move` 包 try。

## 约定与坑
- COM 差异：Word Open/SaveAs 是 `obj&`，Excel 是 `object`，**PPT 强类型按值**；`Presentation.Close()` 无参；`DisplayAlerts <- ppAlertsNone`。
- 源文件 UTF-8 + BOM、CRLF；子进程输出中文先 `Console.OutputEncoding <- UTF8`；例外 `AcadMgdHost/test/run.cmd` 纯 ASCII 无 BOM。
- **`.lsp` 必须 GBK + CRLF**（AutoCAD 2020 R23.1 按本地 ANSI 读，UTF-8 带 BOM 的 `(load ...)` 直接失败）；生成前检查 GBK 字节不含 `"` `\` `(` `)`。
- **`(getvar "PID")` 在 R23.1 返回 nil**；证明跑在哪只 acad 用 `(getvar "DWGNAME")` / `ACADVER`。写文件一律 `(vl-princ-to-string …)`，别用 `itoa`（nil/实数报错 → 0 字节文件 + 文件锁）。
- **LISP 有事件**：`BeginLisp`/`EndLisp`/`LispCancelled`；`(progn …)` 内嵌 `(command …)` 的 Begin/EndCommand 先走完、**最后才 EndLisp** —— 等「跑完」只等 EndLisp。`EndLisp` ≠ 成功（求值错误只打印在命令行），要数据让代码自己写结果文件。
- `task` CE 里 `let _ = xxxAsync` = fire-and-forget（「直线没画上」的根因），一律 `let!` / `do!`。
- 递归扫跳过 `.` 开头 / `.idx_cache` / `~$`；项目根已 `git init`。
- 读来路不明的文本一律 `OsHost.TextFile.decodeAllLinesAsync`（`lines` / `dec`），别 `File.ReadAllLines`。

## 焦炉煤气燃烧计算书（`C:\崔胜利\群有\中冶南方都市环保\焦炉煤气燃烧计算书.xlsx`）
细节见 `2026-11-10.md` / `2026-11-11.md`，只留定稿口径：
- **定值**（`焦炉煤气`!E2:E11，用户定的保守自洽）：4 kPa / 4000 kcal/Nm³ 取下限；成分取中值、N2 平衡到 100%（N2 6.45 / CO2 2.5 / CO 6.5 / H2 57 / CH4 24.5 / O2 0.55 / C2+ 2.5）；干基。别抄区间端点。
- **表序**：焦炉煤气 → 需要加热的烟气参数 → 烟气成分计算 → 烟气吸热量计算 → 补热燃烧计算 → 燃烧器选型。成分表是唯一口径。
- **烟气成分**（`GasCompositionTest.fs`，`XUNITHOST_GAS_COMPOSITION=1`）：仙桃 100760 Pa/16.5℃/80% → α=**1.67323**；湿基 H2O 15.635 / O2 7.593 / N2 71.821 / CO2 4.951；干基 9.00 / 85.13 / 5.87。α 不能抄天然气燃烧器的 1.15。
- **吸热**（`GasHeatTest.fs`，`XUNITHOST_GAS_HEAT=1`）：57000 Nm³/h × 175→500℃ = **637.45 万大卡/hr**；业主 530 **别改**。F/G 列有 NIST Shomate 独立复算（偏差 −0.35%）。
- **放热量口径（11-11 用户定稿）**：`C84 = C80 − C79` = LHV − 产物终温显热；T_ad 二分**已删**。α=1.15 是补燃自身助燃空气，与「第一股烟气含氧 8~10%」**互不相干**，别互相反推。现表：LHV 16747.2、显热 3995.84、放热 12751.355 kJ/Nm³（3045.609 kcal/Nm³）、煤气 2093.016 Nm³/h。Fact `-m "*去二分*"`（`XUNITHOST_GAS_TRIM=1`）。
- **燃烧器选型**（第 6 张，`BurnerSelectTest.fs`，`XUNITHOST_BURNER_SELECT=1`）：燃料侧全活链接补热表；**不乘热效率**，散热由 `C28 裕量 1.10` 覆盖；额定负荷按 LHV 输入热功率。现表：设计 837.21 万大卡/h → 额定 920.93、额定煤气 2302.3 Nm³/h。**表已存在就地写、绝不重建**（会冲掉用户输入格与 `C76` 的 `_xll.电机额定功率(C75)`）。
- **助燃风机**（同表 67~102 行）：`C68=C36*(273.15+C71)/273.15`、`C69=ROUNDUP(E69,-2)`；选型 16500 m³/h、4.5 kPa、60% → 轴功率 34.375 kW → 电机 45 kW；风压分项合计 4419 Pa（差 −81）。待厂家确认：`C87` 本体阻力、`C82` 管长、`C83` λ、`C85` Σζ。

## 转换规则（默认不问）
**Word 的宏一律不保留，统一存 `.docx`**（用户 2026-11-11 定）——老 `.doc` 的 VBA 工程多半被「锁定工程」保护，Word 自己 `SaveAs .docm` 也会丢（ForceDisable / ByUI / 开了 `AccessVBOM` 都试过），保宏只会白卡文件；`DocToDocm` 留着，`XUNITHOST_DOC_KEEP_MACRO=1` 才走。`.xls`→`.xlsx`/`.xlsm`；`.ppt`→`.pptx`/`.pptm`。**转一份读回比一份**；只删 `canDelete`，删前备份。
常驻 Excel 挂回式通道（`.xls` 只走这条）：只用 `ExcelTask.GetApplicationAsync pid`，不新建不 Quit 不 kill，只关自己那本；打开老 `.xls` 会改写原件 → 失败前从备份还原；常驻实例反复挂回会劣化，**卡住就重启 Excel**。
