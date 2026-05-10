# 任务执行说明书：移除 VPN 模块 + 全项目审计、改错、优化、完善、扩展、外观美化

## 【角色定义】

你是一个专门负责对 MyTools（.NET Framework 4.8 / WPF / MVVM / Costura 单 exe 打包）进行"一次性大清洗与升级"的助手。你只在仓库 `c:\ToolOfAjun` 中工作，严格遵守 `AGENTS.md` 全部约束。你的工作**不包括**执行其它说明书（见下"协作说明"）里已经分配的实现任务，**仅**负责：①拆除 VPN 模块；②对全项目做一次静态审计并修复明确 Bug；③按本说明书白名单做结构优化与外观美化。

### 协作说明（避免与其它说明书冲突）

仓库内已存在 3 份说明书，本次**禁止重复实现**其内容，只做交叉验证与留位：
- `docs\规划\任务_乱码修复与多模块扩展.md`：乱码修复、Codex 导入导出、截图 NRE、Postgres/MySQL、录像/录音。
- `docs\规划\任务_系统优化与微信清理备份.md`：自动优化、垃圾清理、微信清理/备份/恢复。
- `docs\规划\任务_移除VPN与全局审计优化.md`：**即本文件**。

如果本次改动与它们在同一文件同一段落发生冲突，**以它们为准**；你只负责把自己的改动挪到不冲突的位置，并把冲突点记录到 `docs\规划\冲突待合并.txt`。

## 【任务目标】

交付 6 类结果：
1. **拆除 VPN 模块**：删除"VPN（WireGuard）"左侧导航项、主内容面板、所有绑定命令、服务类、原生二进制与文档描述，使项目编译通过且启动后再无该入口与相关日志。
2. **改错**：修复本说明书"执行步骤 2"列出的 6 个明确 Bug；对不明确、需要运行时验证的问题，仅记录不改代码。
3. **优化**：按本说明书"执行步骤 3"白名单完成结构与性能优化；禁止超范围重构。
4. **完善**：按本说明书"执行步骤 4"补齐缺失的防御性代码与友好提示。
5. **扩展**：按本说明书"执行步骤 5"新增的 4 项**轻量**功能（主题切换、左侧导航折叠、功能搜索、底部状态栏）；**禁止**新增导航模块。
6. **外观美化**：按本说明书"执行步骤 6"统一颜色/圆角/阴影资源字典，不得影响已有功能可用性。

## 【输入说明】

必须先读的文件（绝对路径）：
- `c:\ToolOfAjun\AGENTS.md`
- `c:\ToolOfAjun\src\MyTools\MyTools.csproj`
- `c:\ToolOfAjun\src\MyTools\App.xaml`、`App.xaml.cs`
- `c:\ToolOfAjun\src\MyTools\MainWindow.xaml`（1410 行）、`MainWindow.xaml.cs`
- `c:\ToolOfAjun\src\MyTools\ViewModels\MainViewModel.cs`（约 1500 行，含 69 处 VPN 引用）
- `c:\ToolOfAjun\src\MyTools\Services\WireGuardService.cs`
- `c:\ToolOfAjun\src\MyTools\NativeBinaries\README.txt`
- 三份先行说明书（见"协作说明"），**只读不改**。

## 【执行步骤】

### 步骤 1 — 拆除 VPN 模块（任务目标 1）

1.1 按如下清单**删除**或**整行删除区块**。执行前用 `grep_search` 统计命中数做基线，执行后再次统计必须降为 0：
- 文件整个删除：
  - `c:\ToolOfAjun\src\MyTools\Services\WireGuardService.cs`
- `MainWindow.xaml` 内删除：
  - 左侧导航按钮 `Command="{Binding ShowWireGuardCommand}"` 所在 `<Button>...</Button>` 区块（当前约第 159–173 行，TextBlock 文字为 `VPN`）。
  - `CurrentModule == "WireGuard"` 的主内容 `<Grid>`（用 `grep_search` 定位 `Value="WireGuard"` 确定起止 `<Grid>...</Grid>`，整段删除）。
- `MainViewModel.cs` 内删除：
  - 所有 `Wg*` 私有字段与公开属性（`_wgInterfaceName`、`_wgConfig`、`_isWgConnected`、`_wgStatusText`、`_wgEndpoint`、`_wgAddress`、`_wgServerPublicKey` 等；`WgInterfaceName`、`WgConfig`、`IsWgConnected`、`WgStatusText`、`WgEndpoint`、`WgAddress`、`WgServerPublicKey` 等）。
  - 所有 `ShowWireGuardCommand`、`ToggleWireGuardCommand`、`RefreshWireGuardStatusCommand`（若存在）及其 `new RelayCommand(...)` 构造行。
  - 所有 `using` 到 `MyTools.Services.WireGuardService` 的引用（改成删除，不保留）。
  - 所有包含 `WireGuardService.*` 的方法体。
  - 状态轮询定时器相关代码（若仅为 VPN 服务）。
- `NativeBinaries\README.txt`：删除其中 WireGuard 相关段落；若删除后文件空白，整个文件删除。
- `docs\功能说明.md`：删除"WireGuard 连接"章节（**仅此一节**，其它章节不得改）。
- `AGENTS.md` §七"现有功能模块"首条"WireGuard 连接"：**不要改动**（该文件由用户维护，留提醒给用户自行同步）；在 `docs\开发记录.txt` 末尾注明"AGENTS.md §七 需用户同步删除 WireGuard 行"。

1.2 启动模块变更：
- `MainViewModel.cs` 构造函数中原默认 `CurrentModule = "Network"`（或其它）保持不变；**禁止**默认值改为 `"WireGuard"`。
- 若删除后 `ShowWireGuardCommand` 在 XAML 绑定残留导致 `BindingExpression` 报错，用 `grep_search` 全量扫描 `ShowWireGuardCommand`、`WgStatusText`、`IsWgConnected` 等名字，确认 0 命中。

1.3 编译验证：`dotnet build src\MyTools\MyTools.csproj -c Debug` 必须通过；`dotnet build -c Release` 必须通过且产物仍是单一 `MyTools.exe`。

### 步骤 2 — 改错（任务目标 2）

**只修复以下 6 个确定性 Bug**，不要擅自扩展：

2.1 `MainWindow.xaml.cs` 第 32 行 `Assembly.GetEntryAssembly()?.Location`：Costura 打包后可能返回空字符串；改为：
```csharp
var executablePath = Process.GetCurrentProcess().MainModule?.FileName;
```
并在顶部 `using System.Diagnostics;`。

2.2 `MainWindow.xaml.cs` 的 `OnClosing` 只隐藏、`OnClosed` 里才 `TrayIcon.Dispose()`：当前只要未 `IsExiting` 就 `e.Cancel=true` 直接 `Hide()`，导致 `OnClosed` 永不执行，托盘图标与热键在"退出程序"时才释放。把清理逻辑改由"退出程序"命令 (`ExitCommand`) 集中完成（调用 `HotkeyService.Unregister()` + `TrayIcon.Dispose()` + `Application.Shutdown()`）。`OnClosed` 保留原逻辑作兜底。

2.3 `App.xaml.cs` `Mutex` 创建使用 `initiallyOwned=true` 且在非新实例分支未释放：当 `isNewInstance=false` 时不得调用 `ReleaseMutex()`；当前代码在 `OnExit` 里无条件 `ReleaseMutex()`，可能抛 `ApplicationException`。改为：
```csharp
if (_singleInstanceMutex != null && _ownsMutex)
{
    try { _singleInstanceMutex.ReleaseMutex(); } catch (ApplicationException) { }
}
```
新增 `private static bool _ownsMutex;`，仅在 `isNewInstance==true` 分支内 `_ownsMutex = true;`。

2.4 与"任务_乱码修复与多模块扩展.md"步骤 2 交叉：如果乱码任务**尚未执行**，本次**不**做乱码替换（避免与另一 brief 冲突）；但必须在本次步骤 1.1 删除 VPN 时，一并删掉 `_wgStatusText = "鏈繛鎺?"` 这类行（因字段整体删除）。

2.5 `HotkeyService.Initialize` 在 `MainWindow.OnSourceInitialized` 中调用后，立即 `vm.ReRegisterHotkey()`。若 `ReRegisterHotkey` 方法在 VM 构造时读取了尚未加载的 `AppSettingsService`，会得到默认值而非用户值。静态检查：用 `grep_search` 读 `ReRegisterHotkey` 实现；若方法开头未 `await LoadScreenshotSettingsAsync()` 或其等价物，则保持原状并在日志中提醒（**不要**擅改，因为截图相关改动属另一 brief）。

2.6 所有 `catch { }` 空吞异常审计：用 `grep_search` 正则 `catch\s*\(?.*\)?\s*\{\s*\}`，对 `App.xaml.cs` `MainWindow.xaml.cs` 内命中项，改为 `catch (Exception ex) { AppLogService.Error(ex, "Silent catch at <位置描述>"); }`。**仅**改这两个文件，避免波及其它被另一 brief 负责的文件。

### 步骤 3 — 优化（任务目标 3）

**只做以下项，禁止扩展：**

3.1 资源字典抽取（落实 `AGENTS.md §4.1`）：新建 `Resources\Theme.xaml`，定义以下 `<SolidColorBrush>` 与 `<sys:Double>`：
- 画板色 `BrushSidebar=#172030`、`BrushSidebarDivider=#253347`、`BrushSidebarHover=#1F3347`、`BrushSidebarPressed=#253F58`、`BrushSidebarActive=#1A3048`、`BrushSidebarText=#8FA8C0`、`BrushSidebarTextHover=#C8DCF0`、`BrushSidebarTextActive=White`。
- 圆角 `CornerRadiusCard=8`、`CornerRadiusButton=6`。
- 字号 `FontSizeTitle=20`、`FontSizeSubtitle=15`、`FontSizeBody=13`。
在 `App.xaml` 的 `Application.Resources` 的 `ResourceDictionary.MergedDictionaries` 中追加合并。`MainWindow.xaml` 中所有硬编码 `#172030` / `#253347` / `#1F3347` / `#253F58` / `#1A3048` / `#8FA8C0` / `#C8DCF0` 改为 `{DynamicResource BrushSidebar*}`。**仅替换已列出的颜色**，不要触碰 MaterialDesign 自带资源键。

3.2 按需懒加载：`MainViewModel` 构造函数内不再立刻 `await LoadCodexProfilesAsync()` / `LoadSqlConnectionHistoryAsync()` 等磁盘 IO 操作；改为在第一次切到对应模块（`CurrentModule` setter 中）时 `_ = LoadXxxAsync()`（fire-and-forget，加 `Interlocked` 防重复）。若发现这些方法已经是此模式，跳过。

3.3 Serilog 日志 rolling：`AppLogService.Initialize` 中对 `WriteTo.File` 配置 `rollingInterval: RollingInterval.Day, retainedFileCountLimit: 14, fileSizeLimitBytes: 10 * 1024 * 1024, rollOnFileSizeLimit: true`。若当前已经是这样跳过。

3.4 `MainWindow.xaml` 93KB 单文件：**本次不拆**（拆分风险高）；仅在文件顶部添加 `<!-- region: ... --> ... <!-- endregion -->` 注释分段，共 8 段：`Navigation` / `HomeView` / `NetworkPanel` / `StartupPanel` / `SystemPanel` / `SqlExportPanel` / `ScreenshotPanel` / `CodexProfilesPanel`。若实际模块名不同，以实际 XAML 段名为准，但必须一一对应且总数 = 8。便于后续 brief 拆分。

3.5 `grep_search` 所有 `File.ReadAllBytes`、`File.ReadAllText`、`File.WriteAllBytes`、`File.WriteAllText` 同步 IO 调用；若调用位于 `async` 方法中，改为 `await File.ReadAllBytesAsync` 等（**仅**在 net48 没有异步 API 时才改成 `FileStream` + `ReadAsync`）。**不在** async 方法中的保持不变。

### 步骤 4 — 完善（任务目标 4）

4.1 全局 `MessageBox.Show` 审计：必须提供 4 参数版（带 `MessageBoxButton` 与 `MessageBoxImage`）。仅修复 `App.xaml.cs` 与 `MainWindow.xaml.cs`；其它文件由各自 brief 负责。

4.2 异常友好文案：`App.xaml.cs::BuildUserMessage` 添加"日志文件打开按钮"——把 `MessageBox` 改为在日志路径下方加一行 `点击 [确定] 关闭本窗口，可在此路径查看完整日志`（仅文案变更，不引入新窗口）。

4.3 `NativeBinaries\README.txt` 改为新版内容：
```
此目录用于放置随 MyTools.exe 一起分发的原生二进制依赖与脚本。
当前依赖：
- LockWin10_22H2.ps1（锁屏脚本）
- （预留）ffmpeg\ffmpeg.exe 用于录像/录音功能
```

4.4 `docs\功能说明.md` 顶部加一条"模块总览"表（模块名、入口、负责 Service、主要文件），并把"WireGuard 连接"章节**删除**。

### 步骤 5 — 扩展（任务目标 5）

**仅**新增以下 4 项轻量功能。**禁止**新增导航模块。

5.1 **主题切换（浅/深）**：标题栏右侧加一个 `ToggleButton`（`PackIcon Kind="ThemeLightDark"`），切换 MaterialDesignThemes `IBaseTheme.Light / Dark`（通过 `PaletteHelper().SetTheme`）。持久化到 `AppSettingsService` 的新属性 `bool IsDarkMode`；启动时应用。

5.2 **左侧导航折叠**：标题栏左上角加一个 `ToggleButton`（`PackIcon Kind="Menu"`），把左侧 `ColumnDefinition Width="220"` 通过绑定切换到 `64`；折叠态隐藏 `TextBlock`，仅显示图标。状态持久化到 `AppSettingsService.IsSidebarCollapsed`。

5.3 **功能搜索**：标题栏中部加一个 `TextBox` + `PackIcon Kind="Magnify"`（`materialDesign:HintAssist.Hint="搜索模块…"`）。输入关键字匹配模块显示名（`当前网络`、`启动管理`、`系统优化`、`SQL 导出`、`截图工具`、`Codex 配置`），回车跳转到第一个匹配项（设置 `CurrentModule`）。

5.4 **底部状态栏**：在 `MainWindow` 最外层 `Grid` 底部新增 `RowDefinition Height="28"` + `Border Grid.RowSpan="1"`，显示三段：
- 左：`CPU {0:F0}%`
- 中：`内存 {usedGB:F1}/{totalGB:F1} GB`
- 右：`系统盘剩余 {freeGB:F1} GB`
数据由新 `Services\SystemMetricsService.cs` 提供（`PerformanceCounter("Processor", "% Processor Time", "_Total")` + `GlobalMemoryStatusEx` + `DriveInfo`）。`DispatcherTimer` 每 2 秒刷新；VM 析构或 `OnExit` 时停止。若 `PerformanceCounter` 初始化抛异常，退化为仅显示内存与磁盘，不抛出。

### 步骤 6 — 外观美化（任务目标 6）

6.1 卡片统一：所有 `materialDesign:Card` 未设 `UniformCornerRadius` 的，统一加 `UniformCornerRadius="8"`；未设 `materialDesign:ShadowAssist.ShadowDepth` 的设为 `Depth2`。**仅**修改 `MainWindow.xaml`。

6.2 顶部导航头部：把左侧 `Border Background="#172030"` 顶部 Logo 区渐变改为 `LinearGradientBrush StartPoint="0,0" EndPoint="0,1"`，`#1E2A3D → #172030`；字体 `FontSize` 从 17 → 18；图标大小从 26 → 28。

6.3 活动导航项：`NavItemStyle` 的 `ActiveBar` 宽度从 3 → 4；添加 `RadiusX="2" RadiusY="2"`。

6.4 按钮：给 `MaterialDesignRaisedButton` 之外的普通按钮加入 `CornerRadius="6"` 的 Setter（在 Theme.xaml 中集中设置 `<Style TargetType="Button" BasedOn="{StaticResource MaterialDesignRaisedButton}">`，**仅作为 Key 样式，不覆盖默认**）。

6.5 窗口标题栏：`MetroWindow` 的 `TitleForeground` 绑定到主题文字色，`WindowTitleBrush` 使用 `BrushSidebar`。保证深色模式下标题栏清晰。

6.6 `HomeView` 空白页增强：把当前"欢迎使用 MyTools"文案下方加 4 个快捷入口卡片（当前网络、系统优化、SQL 导出、截图工具），每卡片 120x120，`Command="{Binding ShowXxxCommand}"`。

### 步骤 7 — 文档与构建收尾

7.1 `docs\开发记录.txt` 追加 1 条 `## [YYYY-MM-DD] 移除 VPN 模块与全项目审计优化`，列出被删除/新增/修改的主要文件。
7.2 运行 `dotnet build src\MyTools\MyTools.csproj -c Release`，把输出末尾 30 行追加到 `docs\规划\构建基线日志.txt`（标记 `=== 任务_移除VPN与全局审计优化 最终构建 ===`）。
7.3 打开 `MyTools.exe` 通过手工流程：①左侧无 VPN 入口 ②主题切换可用 ③导航折叠可用 ④搜索"SQL"能跳转 ⑤状态栏显示三段指标 ⑥所有原模块（启动管理、系统优化、SQL 导出、截图工具、Codex 配置、当前网络）均能正常打开并渲染。任何一项失败记录到 `docs\规划\手工验证失败.txt`。

## 【输出要求】

必须产出：
- 删除：`Services\WireGuardService.cs`；`MainWindow.xaml` 的 VPN 导航项与面板；`MainViewModel.cs` 的全部 `Wg*` 代码；`NativeBinaries\README.txt` 的 WireGuard 段；`docs\功能说明.md` 的 WireGuard 章节。
- 新增：`Resources\Theme.xaml`；`Services\SystemMetricsService.cs`；`docs\规划\构建基线日志.txt`（若不存在创建）。
- 修改：`App.xaml`（合并 Theme 字典）、`App.xaml.cs`（Mutex 修复、友好异常文案）、`MainWindow.xaml`（导航删除、区域注释、卡片圆角、Logo 渐变、快捷入口、主题/折叠/搜索/状态栏）、`MainWindow.xaml.cs`（`MainModule.FileName`、清理逻辑集中、空 catch 改造）、`MainViewModel.cs`（VPN 清除、懒加载、主题/折叠属性、搜索命令）、`AppSettingsService.cs`（新字段 `IsDarkMode`、`IsSidebarCollapsed`）、`AppLogService.cs`（rolling 配置）。
- 日志：`docs\开发记录.txt` 追加 1 条。

**明确禁止**：
- 禁止新增任何 NuGet 包。
- 禁止把 VPN 改名或折叠起来；必须物理删除。
- 禁止重写 / 大范围重构 `MainViewModel.cs`；本次只做 VPN 删除 + 懒加载 + 新增 2 个扩展属性 + 搜索命令（4 种改动）。
- 禁止更改目标框架（保持 `net48`）。
- 禁止触碰另两份 brief 里的目标文件段落（若需要修改那些段落，留下 TODO 注释引用对应 brief 文件名）。
- 禁止删除全局异常挂接（`DispatcherUnhandledException`、`AppDomain.UnhandledException`、`TaskScheduler.UnobservedTaskException`）。
- 禁止删除单实例 Mutex 机制。
- 禁止使用 `DrawingImage` 作为 `NotifyIcon` 图标源。

## 【边界与限制】

- 如果 `grep_search` 统计的 `WireGuard|Wg|wireguard|VPN` 命中数在步骤 1 执行后未归 0，**停止**后续步骤并把残留清单写入 `docs\规划\VPN残留.txt`，由用户确认后再续。
- 如果某处 `MainViewModel.cs` 的改动在另一份 brief 的目标段落（例如乱码修复涉及的 `_sqlStatusMessage` 行），**不要同时改**：只做 VPN 字段删除，其它字面量保留原样。
- 如果 `PerformanceCounter("Processor","% Processor Time","_Total")` 在首次启动抛 `InvalidOperationException`（系统英文计数器被本地化），状态栏 CPU 段显示 `CPU --`，**不要**尝试多语言兜底（体积/复杂度不值当）。
- 如果合并 Theme.xaml 导致 MaterialDesignThemes 默认键被意外覆盖（编译成功但 UI 崩塌），**立即回滚**并在 `docs\规划\主题合并问题.txt` 记录。
- 如果 Debug 构建通过但 Release 构建 Costura 报 "could not embed"，先检查是否有未知新依赖被引入；若无，则 Release 构建失败记录到 `docs\规划\构建基线日志.txt` 的 `[待确认]` 段，**不要**降级 Costura 版本或改其 config。
- 如果 `MainWindow.xaml` 中 VPN 面板的 `<Grid>` 与其它面板 `<Grid>` 之间存在共用样式资源，**只**删除面板本身，共享资源留着。
- 如果发现其它疑似 Bug（超出本清单 6 项），记录到 `docs\规划\疑似Bug.txt`，**不要**自行修复。

## 【示例】

**正例 1**（VPN 导航按钮删除）：用 `grep_search` 命中 `ShowWireGuardCommand` 所在 `<Button>` → `</Button>` 区块，整段删除；相邻两个 `<Button>` 之间的空行保持 1 行。

**反例 1**：只删除 TextBlock 文字 `VPN` 改成 `保留`，留下按钮空壳 —— 没有达到"物理删除"目标，**禁止**。

**正例 2**（Mutex 修复）：新增 `_ownsMutex` 字段，仅在 `isNewInstance==true` 时置 `true`，`OnExit` 中条件释放。

**反例 2**：把 `new Mutex(true, ...)` 改成 `new Mutex(false, ...)` —— 虽然也能避免释放异常，但失去了单实例互锁语义，**禁止**。

**正例 3**（Theme.xaml 抽取）：只替换 `#172030` / `#253347` 等 8 个已列出的颜色。

**反例 3**：把 `MaterialDesignPaper` / `PrimaryHueMidBrush` 等 MaterialDesign 自带键也硬编码到 Theme.xaml —— 会破坏主题切换，**禁止**。

**正例 4**（底部状态栏 CPU 失败兜底）：`try { _cpuCounter = new PerformanceCounter(...); } catch { _cpuCounter = null; }`，刷新时 `CpuText = _cpuCounter == null ? "CPU --" : $"CPU {_cpuCounter.NextValue():F0}%"`。

**反例 4**：在刷新 tick 内 `throw` 让 `DispatcherTimer` 被吞异常 —— UI 线程崩。

**正例 5**（卡片圆角批量化）：只改未显式写 `UniformCornerRadius` 的 Card。

**反例 5**：替换所有 `Card` 的 `Margin`、`Padding` 统一化 —— 超出"仅圆角/阴影"范围，**禁止**。

## 【自检清单】

- [ ] `dotnet build src\MyTools\MyTools.csproj -c Release` 通过；产物仍是单 `MyTools.exe`。
- [ ] `grep_search` 搜索 `WireGuard`、`Wg[A-Z]`、`ShowWireGuardCommand`、`wireguard.exe`，主工程内命中数 = 0（docs 旧记录、AGENTS.md、其它 brief 除外）。
- [ ] 启动后左侧导航无 VPN 入口；主界面渲染正常；Home 页显示 4 个快捷入口。
- [ ] 标题栏主题切换按钮可用，深浅色切换 ≤ 300ms，且次启动记忆上次主题。
- [ ] 左侧导航折叠按钮可用，宽度在 220 ↔ 64 之间切换；折叠态只显示图标，悬停显示 Tooltip。
- [ ] 搜索框输入"SQL"并回车，立即切到 SQL 导出模块。
- [ ] 底部状态栏三段均显示合理数值；禁用 `PerformanceCounter` 模拟故障时 CPU 段显示 `CPU --`，其余仍正常。
- [ ] `App.xaml.cs` Mutex 非新实例分支不再抛 `ApplicationException`（可用第二个进程实例手动验证）。
- [ ] `MainWindow.xaml.cs` 空 `catch {}` 均已带 `AppLogService.Error` 记录（`App.xaml.cs` 与 `MainWindow.xaml.cs` 内）。
- [ ] `Resources\Theme.xaml` 中的 8 个画板颜色键存在且被 `MainWindow.xaml` 引用。
- [ ] `docs\开发记录.txt` 已追加 1 条 dated 记录。
- [ ] 没有新增 NuGet 包；`.csproj` 依赖清单未膨胀。
- [ ] `docs\规划\构建基线日志.txt` 含本任务最终 Release 构建输出。
