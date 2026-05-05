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
- **推荐构建命令**（使用仓库内置 SDK）：
  ```powershell
  .dotnet\dotnet.exe build src\MyTools\MyTools.csproj -c Release
  ```
- Release 产物为单一 `MyTools.exe`，须确认 Costura.Fody 已合并全部托管 DLL，且 `SQLite.Interop.dll` 通过 `EmbeddedResource` 嵌入。
- 应用图标固定为 `src/MyTools/Resources/AppIcon.ico`（在 csproj 中通过 `<ApplicationIcon>` 指定）。

## 六、工作流程
1. **需求定义**：每次新增功能先在 `docs/功能说明.md` 中简述逻辑、核心步骤与涉及文件。
2. **高效开发**：优先编写 Service 与 ViewModel，最后打磨 UI。
3. **分发打包**：编译 Release 时确认 Costura.Fody 已合并资源，产物为单一 `.exe`。
4. **日志维护**：每次更新后，在 `docs/开发记录.txt` 中按 `## [YYYY-MM-DD] 标题` 格式追加修改点。

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

---
