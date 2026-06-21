# MyTools - Windows 个人实用工具集

本项目是一个基于 .NET Framework 4.8 的原生 Windows 桌面工具，旨在提供高效、美观、且在 Win7/10/11 环境下均可直接运行的个人常用功能（如系统增强、轻量数据处理等）。

## 一、项目原则
1. **高兼容性**：必须支持 Windows 7 SP1 以上所有系统，严禁引入 .NET Core 或高版本运行时依赖。
2. **极简分发**：最终产物必须是单一的 `.exe` 文件（利用 Costura.Fody 打包），严禁产生零散 DLL。
3. **视觉 Premium**：界面必须现代、流畅，符合 Material Design 审美。
4. **本地优先**：数据与日志均存放在程序同级目录，不依赖云端。当前轻量配置使用 JSON + DPAPI；SQLite 为预留方案，需要结构化存储时再启用。

## 二、核心技术栈
- **框架**: .NET Framework 4.8 / WPF（MVVM 模式），csproj 采用 `Microsoft.NET.Sdk.WindowsDesktop` SDK 风格。
- **UI 组件**:
  - **MaterialDesignThemes**：全局风格。
  - **MahApps.Metro**：窗体容器（`MetroWindow`）。
  - **FluentWPF**：动态毛玻璃效果。
  - **Hardcodet.NotifyIcon.Wpf**：系统托盘图标与右键菜单。
- **数据层**:
  - **System.Data.SqlClient**：SQL Server 连接与查询（导出模块使用）。
  - **System.Data.SQLite + Dapper**：预留依赖，业务真正需要本地结构化存储时再启用。
  - **Newtonsoft.Json**：本地配置序列化（如 SQL 连接历史 `MyTools.sqlhistory.json`）。
- **工具**:
  - **Serilog + Serilog.Sinks.File**：异步文件日志。
  - **Costura.Fody**：单文件打包，并嵌入 `SQLite.Interop.dll`（x86/x64）。

## 三、项目结构
- `src/MyTools/`：主工程根目录。
  - `App.xaml(.cs)`：应用入口、全局异常处理、显式创建主窗口（不使用 `StartupUri`）。
  - `MainWindow.xaml(.cs)`：主窗体与托盘宿主。
  - `Services/`：业务服务（`WireGuardService`、`SqlExportService`、`SqlConnectionHistoryService`、`AppLogService`、`StartupService`、`NetworkService` 等）。
  - `ViewModels/`：MVVM 视图模型与转换器。
  - `Resources/`：图标、图片等嵌入资源（`AppIcon.ico` / `AppIcon.png`）。
  - `NativeBinaries/`：随产物输出的原生依赖与脚本（如 `LockWin10_22H2.ps1`）。
- `docs/`：`功能说明.md` 与 `开发记录.txt`。
- `.dotnet/`：仓库内置的 .NET SDK，用于离线构建。

## 四、开发规范
### 4.1 UI/UX
- 必须支持 **Per-Monitor V2 高 DPI** 自适应。
- 使用 `MetroWindow` 作为主窗体，并启用 `MaterialDesign` 主题集成。
- 所有颜色、字体大小必须定义在 `ResourceDictionary` 中，方便全局调整。
- **托盘图标禁用 `DrawingImage`**：必须使用位图资源（如 `Resources/AppIcon.png`），否则 `Hardcodet.NotifyIcon.Wpf` 在 XAML 解析阶段会触发 `UriFormatException` → `XamlParseException`，导致启动闪退。

### 4.2 编码
- **异步原则**：所有磁盘 IO、网络请求、数据库调用必须使用 `async/await`，禁止卡顿 UI。
- **依赖管理**：尽量减少 NuGet 包，优先选择轻量级、无二次依赖的库。
- **白名单查询**：执行 SQL 时，库名/架构名/表名必须来自程序加载的白名单，并使用方括号转义，禁止拼接用户自定义 SQL。

### 4.3 启动健壮性
- 同时挂接 `DispatcherUnhandledException`、`AppDomain.UnhandledException`、`TaskScheduler.UnobservedTaskException`。
- 启动失败时写入程序同目录的 `MyTools.startup.log` 并弹窗提示用户日志位置。
- 运行期异常通过 Serilog 写入文件日志。

### 4.4 安全
- 敏感凭据（如 SQL 密码）必须使用 **Windows DPAPI（当前用户范围）** 加密后再写入本地配置，严禁明文落盘。
- 日志严禁记录密码或完整连接字符串，仅记录服务器/用户名/操作步骤等非敏感信息。

## 五、构建与分发
- **推荐构建命令**：
  ```powershell
  dotnet build src\MyTools\MyTools.csproj -c Release
  ```
- Release 产物为单一 `MyTools.exe`，须确认 Costura.Fody 已合并全部托管 DLL，且 `SQLite.Interop.dll` 通过 `EmbeddedResource` 嵌入。
- 应用图标固定为 `src/MyTools/Resources/AppIcon.ico`（在 csproj 中通过 `<ApplicationIcon>` 指定）。

## 六、工作流程
1. **需求定义**：每次新增功能先在 `docs/功能说明.md` 中简述逻辑、核心步骤与涉及文件。
2. **高效开发**：优先编写 Service 与 ViewModel，最后打磨 UI。
3. **分发打包**：编译 Release 时确认 Costura.Fody 已合并资源，产物为单一 `.exe`。
4. **日志维护**：每次更新后，在 `docs/开发记录.txt` 中按 `## [YYYY-MM-DD] 标题` 格式追加修改点。
5. **模块逻辑同步**：功能模块有变动时，须同步更新 `docs/程序逻辑.md` 中对应的描述（新增功能增加章节，删除功能移除章节，逻辑变更在原章节内更新）。

### 6.1 场景驱动开发范式（开qg）
- 用户要求模拟真实使用流程、跑一遍功能场景、按用户操作方式完善工具，或输入 `开qg` 触发工作场景准备时，必须先读取 `docs/场景驱动开发/场景驱动开发范式.md`。
- 场景准备态只收集目标、使用角色和入口，不立即修改代码；信息齐全后生成确认卡片，用户回复 `确认执行` 后才进入读取、实施、文档同步和验证。
- 场景执行时必须优先检查现有 WPF 模块、Service、ViewModel、命令、配置和通用能力；已有承载点不足时在原处补全，确认没有承载点后才允许新增模块。
- 场景相关文档统一维护在 `docs/场景驱动开发/`：`业务流程清单.md` 记录真实使用步骤，`场景实现核查台账.md` 记录实现状态和证据，`控件交互逻辑说明.md` 记录页面、弹框、按钮、校验、提示和权限展示。
- 原有 `docs/功能说明.md`、`docs/程序逻辑.md`、`docs/开发记录.txt` 仍按本项目规则同步维护；场景驱动文档用于补充真实流程、核查证据和控件事实。

### 6.2 Loop Engineering 闭环规则
- 本项目使用 `.agents/skills/mytools-loop-engineering/SKILL.md` 作为 Codex 闭环执行流程：Inspect → Plan → Implement → Validate → Repair Loop → Review → Deposit。
- 常规开发、修复、重构、发布准备、安装包重打和评审任务默认按该 skill 执行；一次性问答、只读解释、用户明确要求不改代码的任务除外。
- 默认验证入口为 `scripts/codex-eval.ps1`；普通小任务优先使用 `powershell -ExecutionPolicy Bypass -File scripts\codex-eval.ps1 -Quick`，发布准备或安装包相关改动使用不带范围参数的全量验证；也可按任务范围追加 `-Build` / `-Installer` 定向验证。
- 验证失败后必须读取具体失败日志，做最小修复并重跑相关验证；同一问题最多进行 5 次修复验证循环。超过 5 次仍未解决时，停止继续试错，说明已执行命令、关键日志、已尝试修复、剩余缺口和需要用户或外部状态补充的内容。
- Bugfix 如能用自动化测试、构建脚本或可重复场景稳定复现，必须优先补充或更新回归测试；若暂不适合补测，交付说明中必须说明原因。
- 非微小代码或规则改动在交付前必须使用独立审查 agent `.codex/agents/mytools-reviewer.toml`；该 agent 只读审查 diff、验证结果和项目规则，输出阻断问题、建议、待确认项和沉淀建议，阻断问题未处理前不得宣称完成。
- 问题解决后必须判断经验沉淀位置：可执行约束优先沉淀为测试或 `scripts/codex-eval.ps1`；稳定项目规则沉淀到 `AGENTS.md`；重复流程沉淀到 `.agents/skills/`；模块知识、运行手册和业务事实沉淀到 `docs/`、`docs/规划/` 或 `docs/场景驱动开发/`。不得为了记录进度另建总结类 Markdown，除非用户明确要求。

## 七、现有功能模块
- **WireGuard 连接**：`WireGuardService` 调用本地/系统 `wireguard.exe`，通过 `installtunnelservice` / `uninstalltunnelservice` 控制隧道，结合网卡状态判断连通性。
- **系统托盘**：关闭按钮默认隐藏到托盘，托盘提供"显示窗口 / 退出程序"菜单及双击恢复。
- **SQL Server 导出 Excel**：仅支持 SQL 身份验证；自动降级 `sys.databases` → `sp_databases`；自写 OpenXML `.xlsx`，无需安装 Office；导出前校验 `1,048,576` 行上限。
- **SQL 连接历史**：服务器/用户名/密码本地保存（密码 DPAPI 加密），启动回填最后一次成功连接。
- **应用日志**：`AppLogService` 记录 SQL 导出等关键步骤，遵循 4.4 安全规范。

## 八、任务说明书生成规则（Task Brief Generator）

**触发条件**：用户说"不动标："或类似表达。

**职责**：不执行该任务本身，而是生成一份结构化的「任务执行说明书」放在 `docs/规划/` 里，供能力较弱的 AI 直接使用。

**说明书必须包含**：

**【角色定义】** 用"你是一个专门负责……的助手"开头。

**【任务目标】** 一句话清晰描述最终要交付什么。

**【执行步骤】** 分步编号，每步只做一件事。禁止"适当""合理""一些"等模糊词，全部替换为具体数字或可验证描述。

**【输入说明】** 描述该 AI 将收到什么内容（格式、来源、示例）。

**【输出要求】** 列出必须包含的元素与明确禁止出现的内容。

**【边界与限制】** 列出不能做的事，遇到不确定时的处理：
> 如果遇到 X，则执行 Y；无法判断则标注 [待确认] 并说明原因。

**【示例】** 每个关键判断点至少一个正例一个反例。

**【自检清单】** 可勾选的检查项列表。

生成完毕后，自动扮演一个能力较弱的 AI 通读说明书找出歧义并修改，只输出说明书本身，不加前言/解释/元注释。Markdown 格式，可直接复制使用。

## 九、性能与启动规则（不可违反）

### 9.1 启动期黄金 500 ms 原则
- `App.OnStartup` 与 `MainWindow` 构造函数中严禁执行任何耗时超过 50 ms 的磁盘 IO、注册表查询、WMI 调用或大对象反序列化。
- 启动期允许的同步操作仅限：单实例 Mutex、全局异常处理器挂接、合并资源字典加载、命令对象创建、空 `ObservableCollection<T>` 创建。
- 加载历史、扫描磁盘、查询硬件、读取注册表、枚举启动项、枚举已安装程序等任务，必须通过 `Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, ...)` 延后，或在用户点击对应模块时按需加载。

### 9.2 ViewModel 构造约束
- `MainViewModel` 构造函数禁止直接创建重型子 ViewModel；`ScheduleViewModel`、`SystemSettingsViewModel` 等必须以懒加载属性形式暴露。
- 构造函数末尾禁止直接启动批量后台任务；统一通过私有方法 `ScheduleStartupBackgroundLoads()` 在 `ApplicationIdle` 优先级派发。
- 任何 `ObservableCollection<T>` 在构造时禁止预填充大量数据；数据必须由延迟加载方法填入。

### 9.3 面板可见性约束
- `MainWindow.xaml` 中非 Home 模块面板默认使用 `Visibility="Collapsed"`，仅在 `CurrentModule` 匹配时显示。
- 切换可见性必须通过 `DataTrigger Binding=CurrentModule` 实现；禁止用 `Hidden` 代替 `Collapsed`。
- 单个面板 XAML 行数超过 800 行时，必须抽出到 `src\MyTools\Views\<Module>View.xaml` 作为 `UserControl`，并保持继承主窗口 `DataContext`。

### 9.4 日志与磁盘 IO 约束
- Serilog `File` Sink 必须使用缓冲写入，`flushToDiskInterval` 不得超过 2 秒。
- 启动期 `MyTools.startup.log` 的追加写入必须延后到主窗口可见之后。
- 单条日志严禁记录完整连接字符串、密码或超过 1 KB 的大文本；只记录关键状态、文件名、服务器名、用户名等非敏感摘要。

### 9.5 内存与 GC 约束
- `App.config` 的 `<runtime>` 节点必须保留 `<gcServer enabled="true" />` 与 `<gcConcurrent enabled="true" />`。
- 长期持有的 `ObservableCollection` 元素数超过 5000 项时，必须启用 WPF 虚拟化列表，不能依赖无虚拟化布局渲染全部元素。
- 新增事件订阅时必须同步规划退订位置；长期对象订阅短生命周期对象事件时必须实现 `IDisposable` 或在 `Unloaded` 中解绑。

### 9.6 性能基线维护
- 修改 `App.OnStartup`、`MainWindow` 构造函数、`MainViewModel` 构造函数时，必须在 `docs/规划/` 下记录启动基线或说明未测原因。
- 启动首帧、稳态私有工作集、单 exe 体积任一指标回退超过 5% 时，必须回滚或说明不可回避原因。
- 新增静态资源导致单 exe 体积增长超过 1% 时，必须在 `docs/开发记录.txt` 说明文件名、大小和用途。

### 9.7 禁止事项（红线）
- 禁止在启动期同步调用 WMI。
- 禁止在构造函数中调用 `Task.Run(...).Result` 或 `.Wait()`。
- 禁止用 `Application.Current.Dispatcher.Invoke` 同步等待 UI 线程结果。
- 禁止为了追求速度关闭全局异常处理、DPAPI 加密、SQL 白名单校验或日志脱敏。

---
