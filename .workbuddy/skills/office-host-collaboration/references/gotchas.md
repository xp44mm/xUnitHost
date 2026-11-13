# 踩过的坑（动手前扫一眼）

## 测试框架

- **xunit v3 不能用 `dotnet test`**：报「VSTest 不支持」。直接跑
  `xUnitHost/bin/Debug/net10.0/xUnitHost.exe`。
- **测试并行关不掉**：`[<assembly: CollectionBehavior(DisableTestParallelization=true)>]` 无效，
  必须放 `xUnitHost/xunit.runner.json`（`parallelizeTestCollections: false`）并
  `CopyToOutputDirectory`。否则两个 COM 测试并行抢同一个实例。
- xunit v3 对 string 的 `Assert.NotEqual(a, b)` 有二义性（`'T` vs `'T seq`），用 `Assert.True(a <> b)`。

## COM 与线程

- **必须 STA**。会话层把所有 COM 调用排在**同一个固定 STA 线程**（`BlockingCollection` 队列）：
  host 是长活的，RCW 跨公寓要封送，而封送要求创建它的 STA 线程在泵消息。
- **RCW 按 IUnknown 身份缓存**：`Marshal.GetObjectForIUnknown` 对同一个指针返回**同一个 RCW**。
  所以拿到 app 后随手 `FinalReleaseComObject`，会把地址簿里那只**一起拆掉**，
  之后再用就是 `InvalidComObjectException`。释放前先判断「这只我们手上是不是已经有了」。
- `Marshal.GetActiveObject` / `Type.GetTypeFromProgID` 在 netstandard2.0 里没有 →
  直接 P/Invoke `ole32` / `oleaut32`（`CLSIDFromProgID` + `GetActiveObject`）。
- **ROT 的 `Register` 是给服务端用的**：客户端拿别人的代理去登记，flags 0/1/2/3 全 `E_INVALIDARG`。
  Office 自己已经登记好了，客户端只用 `GetActiveObject`。
- **黑板模式按名字取 ROT**：VSTO 插件用 `CreateItemMoniker("!", "WordVSTO.Instance.<pid>")` 登记；
  消费侧造同一个 moniker 用 `IRunningObjectTable.GetObject` 取回（`OsHost.Com.runningAt`）。
  别枚举 ROT，更别退回 `GetActiveObject` —— 多实例机器上那等于认错人。自测要登记就登记自己 new 的托管对象，
  并记得 `Revoke`。
- **返回的 Application 不能跨临时 STA 线程**：`GetApplicationAsync` 只负责把对象取出来，
  后续读写都包 `Sta.run` / `Sta.runAsync`；用 `Sta.here` 的临时线程一退出，就没人给这个 RCW 泵消息了。
- **`AccessibleObjectFromWindow` + `OBJID_NATIVEOM` 要求最内层文档窗口**：Word 是 `_WwG`、
  Excel 是 `EXCEL7`（XLDESK 的子窗口），敲在顶层 `OpusApp`/`XLMAIN` 上一律 `E_FAIL`。

## 退出时弹框 / 进程不退

- **Word / Excel**：只要还有「改过没存」的文档，`Quit()` 就弹「是否保存」，自动化场景没人点 →
  卡死。必须先 `Document.Close(false)` / `Workbook.Close(SaveChanges=false)` 丢改动，再 `Quit`。
- **Excel** 即使 Quit 了也不一定退：客户端还捏着 RCW 时它赖着。做法是
  `FinalReleaseComObject` + `GC.Collect/WaitForPendingFinalizers`，等 5 秒，还不走就 `Kill` 自己起的那些 pid。
- **AutoCAD**：`Documents.Item(n).Close(false)` 会直接 `E_FAIL 0x80004005`（2020 实测），
  必须用 `Application.ActiveDocument.Close(box false, Type.Missing)`。用 `Documents.Item` 关不掉 →
  Quit 照样弹框。
- **只碰自己起的进程**：起实例前后各扫一次同名进程，差集才是这次起的。用户手上那只一个字不动。

## AutoCAD 冷启动

- **别靠轮询 `GetAcadState().IsQuiescent` 判就绪**：它刚起来时会短暂假报 true，这时动它就
  `RPC_E_SERVERFAULT` 甚至崩；而「连中 N 次 + 静置」那套（原 `Ready.fs`）**挡不住 `RPC_E_CALL_REJECTED`，已废弃删除**。
  等就绪一律用事件：`AutoCADTask.sendCommandAsync` / `openDrawingAsync` / `runLispAsync` 各自等自己的 End 事件。
- 冷启动分钟级，别在没有实例的机器上反复试；本机通常开着一只，按 PID 挂回就行。

## AutoCAD：命令 / LISP 的完成判据

- **`SendCommand` 返回 ≠ 命令跑完**：它同步返回只代表「已交给 AutoCAD」。等 `EndCommand`
  用 `AutoCADHost.AutoCADTask.sendCommandAsync`；裸发之后立刻发下一条会撞 `RPC_E_CALL_REJECTED`。
- **命令行串必须自带换行**（`"_.ZOOM\n_E\n"`）：少一个换行，最后一个词留在命令行上等输入，
  整只 acad 一直「忙」，后面所有调用都被拒。
- **LISP 有事件**：`BeginLisp` / `EndLisp` / `LispCancelled`（互操作里就是这个大小写，
  不是全大写 LISP）。一整段 `(progn …)` 是一个求值整体 —— 内部 `(command "_.LINE" …)` 的
  Begin/EndCommand 先各自走完，最后才来 EndLisp；等命令事件会**提前返回**。
  真机（R23.1）已验证 `SendCommand` 送进去的 LISP 确实派 `BeginLisp → EndLisp`。
- **`EndLisp` = 求值结束，≠ 成功**：AutoLISP 的求值错误只在命令行打印一行、不回报调用方。
  要数据/判成败就让 LISP 自己写结果文件；`File.Exists` 不算判据（`open` 之后文件就存在、内容是空的），
  要「能独占打开且内容非空」。
- **`(getvar "PID")` 在 R23.1 上不存在**（返回 `nil`），要证明「跑在哪只 acad 里」写
  `(getvar "DWGNAME")` / `(getvar "ACADVER")`，对照 COM 的 `ActiveDocument.Name`。
- **写文件用 `(vl-princ-to-string …)`，别用 `itoa`**：碰到 `nil` / 实数直接报错，只留下 0 字节文件，
  而且句柄还挂在 acad 里把文件锁住（F# 侧读会一直超时）。上一次没 close 的句柄要在下次代码开头
  `(if bet-f (close bet-f))` 兜一下。
- 测试里跑 LISP 优先选**只读**命令（`_.DIST`），别用 `_.LINE` / `_.CIRCLE` 往用户图纸里加实体。

## 地址 / 指纹

- **每步用新地址**：`exec` 会重拍快照。这一步保存了文档，保存毫秒和路径指纹就变，地址也变；
  老地址会被移除并被拒接。被拒是正常行为，不是 bug。
- 被拒时**用错误信息里给出的新地址**，别自己拼。
- **Word 冷启动是 0 份文档**，`ActiveWindow` 直接抛 → 从主窗口反查拿不到 pid。三级回退：
  主窗口 → 调用方已知 pid → 机器上这种进程只有一个就是它；都不成立记 0，核对时跳过 pid。
- **按地址回来绝不新建**：`resolve` 只接现成的那只（`TryGetRunningApplication`）。
  用 `GetOrCreateApplication` 会在没人接的时候静默起一个新进程，调用方还以为接回了原来那只。
- 多实例时 ROT 只暴露一只，pid 对不上就老实报错，不要乱动。

## F# / 互操作

- **byref 参数不能在 lambda 里取地址**（FS0421）：`let mutable x: obj = box v` 再 `f(&x)`。
- **Word 的参数是 `obj&`（byref），Excel 的参数是按值 `object`**：`doc.SaveAs(&name, &fmt)` 对，
  `book.SaveAs(box path, box 51)` 才对；给 Excel 传 `&x` 会报 FS0421/FS0412。
  **但 Word 集合的 `.Item(i)` 例外**：`Tables.Item(i)` / `Rows.Item(i)` / `Cells.Item(i)` 收的是
  **int 按值**，`doc.Tables.Item(&idx)` 会报 FS0193（byref&lt;obj&gt; 与 int 不兼容）；
  只有 `Rows.Add(&beforeRow)` 仍是 `obj&`。别照 `Documents.Open(&...)` 一律取地址。
- **Word 表格插行**：`table.Rows.Add(&beforeRow)` 在指定行之前插一行，新行继承那一行的格式；
  分类行常是合并行，所以要在**编号数据行**之前插，并断言 `newRow.Cells.Count = 8`。
  单元格文本用 `cell.Range.Text <- "..."`，读回要去掉尾部 `\r\a`（`DocIO.normalize`）。
- **Excel 的 `UsedRange.Value2` 返回的二维数组下界是 1，不是 0**：`grid.[0,0]` 直接越界，
  用 `grid.GetLowerBound 0 .. grid.GetUpperBound 0` 遍历。
- **F# 不允许 `try/with/finally` 写在一起**：先 `let outcome = try body() with ex -> ...`，再单独收尾。
- `Documents.Add()` / `Workbooks.Add()` 可以不传参；`Documents.Add(&template, …)` 才需要四个 mutable。
- 中文在子进程重定向时变乱码：子进程里先 `Console.OutputEncoding <- Encoding.UTF8`，
  父进程给 `StandardOutputEncoding` 也设 UTF-8。控制台程序也一样，不然中文路径全成乱码。
- 项目文件统一 **UTF-8 + BOM、CRLF**，否则 F# 编译器对中文注释/字符串可能报错。

## `task {}` 与异步 Fact

- **Fact 直接写成 `member this.``中文一句话``() : Task = task { … }`**（xunit v3 认 `Task` 返回值）。
  别写 `let work = task { … }` 再 `work.GetAwaiter().GetResult()` —— 异常被包成 `AggregateException`、还白占一条线程；
  存量文件里有这种写法，新写的一律直接返回 `Task`。
- **`let!` / `do!` 不许省**：`let _ = xxxAsync` 是 fire-and-forget，等于没等（「直线没画上」就是这么来的）。
- **`Report.await` 只能在 STA 之外用**：它阻塞等 Task，在 `Sta.run` 的活儿里等就是自己等自己，直接死锁。
  要在固定 STA 上串下一段，用 `let! x = Sta.runAsync(fun () -> …).Unwrap()`。
- 库里的同步入口（`Host.GetOrCreateApplication` / `Host.exec*` / `DocIO.*`）不会替你切线程：
  排进固定 STA 一律 `Sta.runAsync(fun () -> …).Unwrap()`，或干脆用 `*Async` 成品入口。
- RCW 拿到手不等于随处能用：`await` 之后每一步读写仍要包 `Sta.run`。

## 格式转换（doc→docx / docm、xls→xlsx / xlsm）

- **`UsedRange` 的行列数在 .xls 与 .xlsx 之间会不一样**（只有格式的空列算不算），拿尺寸当指纹会假报警。
  比内容就只比「表名 + 格子里的值」。
- **空表不算内容**：.xls 里常多出 Sheet2/Sheet3（只有格式没值），转成 xlsx 后可能就没了，不必报警。
- **`Workbook.HasVBProject` / `Document.HasVBProject` 判宏**：`.xls` 有宏 → `.xlsm`，硬转无宏格式等于删掉宏；
  **判不出宏就报失败、原件保留**，绝不猜「没宏」。转成 `.xlsm` 后还要再读一遍 `HasVBProject` 复核宏还在。
- **Word 反过来：`.doc` 一律存 `.docx`，宏不保留**（用户 2026-11-11 定）——老 `.doc` 的 VBA 工程多半被
  「锁定工程」保护，`SaveAs .docm` 也会丢（ForceDisable / ByUI / 开 `AccessVBOM` 都试过），硬保只会白卡文件；
  `XUNITHOST_DOC_KEEP_MACRO=1` 才走 `DocToDocm`。
- **删之前先清只读属性**：压缩包/邮件来的文件带 ReadOnly，`File.Delete` 直接 `Access denied`。
- 转换要写成「转完立刻读回来比对，一致才算成功」；不一致就保留原件，绝不允许「先删后转」。
- 删原件单独一条命令，只认得到「核对通过」的清单文件；清单别写 BOM，否则按 `"OK\t"` 匹配会漏第一条。

## 读文本文件：别自己猜编码、别手切行

- **一律走 `OsHost.TextFile.decodeAllLinesAsync`（短名 `lines`）**，它一次解决两件事：
  编码（BOM → ascii → UTF-8/GBK 严格试解码，歧义标 `utf-8|gbk`）和行末（`\r\n`/`\n`/`\r` 都切，末尾不多空行）。
- `File.ReadAllLines` 默认 UTF-8：GBK 的 `.lsp` / 老日志 / 插件写的结果文件一读就是一片 `�` 或者干脆抛
  `DecoderFallbackException`；自己 `Encoding.GetEncoding "gb2312"` 又会把 UTF-8 读坏。两边都"看着像对的"，
  最容易把错数据当真值写进报告。
- 手切行常犯两个错：只认 `\r\n`（Linux/Mac 来的文件整篇算一行）、`split` 之后不剥末尾空行
  （行数比真实多 1，按行号取值的断言全跟着偏）。
- **无 BOM 且 UTF-8 / GBK 都能解时它报 `utf-8|gbk`**：这是设计如此 —— 别在 Fact 里断言"等于 gbk"，
  要么接受两个值，要么拿一条已知内容（比如 ASCII 的标记行）当锚。
- 反过来，**写**文件仍按项目约定：UTF-8 + BOM + CRLF；别拿读函数的宽松去放松写入。

## 测试卫生

- 测试里「一开始没有进程在跑」是隐含前提。要依赖它就在 Fact 开头 `requireClean "WINWORD"`，
  环境不干净就 `Assert.Skip`（不是 `failwith`：用户开着图纸/表格干活是常态，别把整轮跑红）。
- 中途失败要 `try/finally` 把 host 收掉，否则残留的 Word 会让后面每个测试都跟着挂。
- 建临时文件放 `%TEMP%` 下独立目录，`finally` 里删干净。
- **报告必须写进 `output`**（构造函数注入的 `ITestOutputHelper`）：模块级 `let` 的 Fact 拿不到它。
- **`-showLiveOutput` 不加就什么都看不见**：xunit 会把 output 收走，只剩一行 Summary。
  这是「报告即返回信息」的前提条件。
- **别顺手把 AutoCAD 拉起来**。`GetOrCreateApplication()` 是「没有就新建」，而 AutoCAD 没有隐藏启动，
  窗口会真的弹到用户脸上。以前那个「三个主机互相打招呼」一个 Fact 把三家全开一遍，用户只是让转个 doc，
  结果弹出一个 CAD。现在**一段通道一个 Fact**（Word→Excel / Excel→CAD），AutoCAD 默认不新起，
  要测那条通道得显式 `XUNITHOST_LAUNCH_ACAD=1`。转换、改文档这类任务只跑 `-m` 相关的 Fact，别跑 `-a`。
- 在用户已经开着的 AutoCAD 里做测试，**别往 `ActiveDocument` 里写东西**（那是他的图纸）。
  一律 `Documents.Add()` 开一张临时图，`finally` 里 `Close(box false, Type.Missing)` 关掉。
- **写之前先认人**：核对文件名 + 最后实质保存时间（`Doc.identify` → `Doc.ensureFresh`）。
  新建的文档 `Path` 是空串、还没落盘；一旦 `Path` 非空就是磁盘上的真文件，直接 `failwithf` 停手。
  这条同样适用于 `xUnitHost/Convert.fs` 之外的任何「打开→改→存」动作。
- 别为了跑通去关用户的 Word / Excel / AutoCAD——宁可跳过。
- **「第二个进程」不用另编程序**：跨进程那条 Fact 用 `Report.selfExe()` 把自己再起一份，
  `-method "*第二个进程*"` 只跑那一条。子进程的回话**写文件**（`XUNITHOST_SAY_AT`），
  别去挑 stdout —— 那里面混着跑框架的输出。读 stdout 和 stderr 要一路并行抽，
  两头都读才不会把管道写满僵住。
- **黑板展示文本不进生产库**：`OsHost` 里没有 `ofText`、没有展示用的标题字段，2026-11-13 起也没有「参数表」了
  ——挂回只认 pid。把「`KEY : value` 的排版」写进库，等于把插件的显示格式变成隐含契约，插件一改文案库就跟着改。
  参考解析只在测试项目 `xUnitHost/BlackboardText.fs`（`ofText → BoardApp`、`docOfText → BoardDoc`），
  格式紧耦合就锁在那儿。
- **Word 零文档时认不出 pid 就拒绝，不要退回「随便接一只」**：`pidOf` 靠 `ActiveWindow` 上溯 `OpusApp`，
  一份文档都没开时是 None。此时按 ROT 名后缀（`!WordVSTO.Instance.<pid>`）认；两样都对不上就报错。
  以前 `matchApp` 在 None 时直接 `Ok app`（认不出就不反驳），多实例机器上会认错人 —— 现在是拒绝。
