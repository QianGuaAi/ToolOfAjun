# frp 隧道穿透功能任务说明书（按当前项目实际结构修正版）

---

## 【角色定义】

你是一个专门负责在 MyTools（.NET Framework 4.8 / WPF / MVVM）中新增 frp 隧道穿透功能的助手。你必须严格遵守当前项目的 `AGENTS.md`：保持 Windows 7 SP1+ 兼容、单 exe 分发、Token 使用 DPAPI 加密、启动期不做重 IO、不引入 .NET Core 或高版本运行时依赖。

---

## 【任务目标】

在 MyTools 左侧导航栏新增独立的「穿透」模块，让用户在软件中配置 frp 服务器、认证 Token、本地端口与远程端口，并通过一键启动/停止 `frpc.exe` 把当前电脑的指定端口暴露到阿里云 Ubuntu 服务器公网端口。

---

## 【当前项目事实】

执行前必须先确认以下事实，不要按旧说明书照抄：

1. 项目根目录是 `c:\exe`。
2. 主工程是 `src\MyTools\MyTools.csproj`。
3. 目标框架是 `net48`，使用 `Microsoft.NET.Sdk.WindowsDesktop`，`UseWPF=true`。
4. `MyTools.csproj` 是 SDK 风格项目，默认自动包含同目录下的 `.cs` 和 `.xaml` 文件；不要手动添加普通 `<Compile Include="...">` 或 `<Page Include="...">`，否则可能触发重复包含错误。
5. 当前 `MainViewModel` 在 `src\MyTools\ViewModels\MainViewModel.cs` 中声明为：
   ```csharp
   public class MainViewModel : INotifyPropertyChanged, IDisposable
   ```
   它不是 `partial`。本任务不创建 `MainViewModel.Frp.cs`，直接修改 `MainViewModel.cs`。
6. `MainWindow.xaml` 已有：
   ```xml
   xmlns:views="clr-namespace:MyTools.Views"
   ```
   不需要重复添加。
7. 已有独立模块页示例：
   - `views:MultimediaPage DataContext="{Binding Multimedia}"`
   - `views:SchedulePage DataContext="{Binding Schedule}"`
8. 新模块应使用同样模式：
   ```xml
   <views:FrpView DataContext="{Binding Frp}">
   ```
9. `ScheduleStartupBackgroundLoads()` 已存在，位置约在 `MainViewModel.cs` 第 563 行，内部使用 `DispatcherPriority.ApplicationIdle` 延后加载。
10. `AppLogService.Warning` 的签名是：
    ```csharp
    AppLogService.Warning(string messageTemplate, params object[] propertyValues)
    ```
    不存在 `Warning(Exception, ...)` 重载。
11. 当前 `NativeBinaries` 已有 `ffmpeg\ffmpeg.exe`，本任务不得修改 ffmpeg 相关文件。
12. `NativeBinaries\**\*.*` 当前被 `<None Update=... CopyToOutputDirectory=PreserveNewest>` 复制到输出目录。为了保持 frpc 单 exe 嵌入，必须对 `NativeBinaries\frpc.exe` 单独 `None Remove`，再作为 `EmbeddedResource` 嵌入。

---

## 【重要版本决策】

### 默认 frp 版本

由于 MyTools 必须支持 Windows 7 SP1+，不要默认使用 `frp v0.61.0` 作为嵌入版。较新的 frp 通常由新版 Go 编译，存在不兼容 Windows 7 的风险。

本任务默认使用：

```text
frp v0.50.0 windows_amd64
```

下载地址：

```text
https://github.com/fatedier/frp/releases/download/v0.50.0/frp_0.50.0_windows_amd64.zip
```

服务器端 Ubuntu 24.04 也建议使用同版本：

```text
https://github.com/fatedier/frp/releases/download/v0.50.0/frp_0.50.0_linux_amd64.tar.gz
```

### 配置格式

`frp v0.50.0` 使用 `.ini` 配置。不要在默认实现中生成 `.toml`。

客户端配置文件名：

```text
%TEMP%\MyTools\frpc.ini
```

服务器配置文件名建议：

```text
/etc/frps.ini
```

---

## 【服务器端前置条件】

MyTools 只负责 Windows 客户端侧的 `frpc` 管理，不负责自动安装阿里云服务器端 `frps`。

用户需要在阿里云 Ubuntu 24.04 上完成：

1. 安装同版本 `frps`。
2. 配置 `/etc/frps.ini`。
3. 使用 systemd 启动并设置开机自启。
4. 在阿里云安全组放行：
   - `7000/tcp`：frpc 连接 frps 的控制端口。
   - 每个远程端口：例如 `8081/tcp`、`8000/tcp`、`33890/tcp`。

服务器端示例：

```ini
[common]
bind_port = 7000
token = 这里填写强密码Token
```

systemd 示例：

```ini
[Unit]
Description=frps
After=network.target

[Service]
ExecStart=/usr/local/bin/frps -c /etc/frps.ini
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
```

---

## 【最终用户使用场景】

用户在每台 Windows 电脑上运行 MyTools，各自配置：

| 电脑 | 本地服务 | 本地端口 | 远程端口 | 外部访问方式 |
|---|---:|---:|---:|---|
| A | Web | 80 | 8081 | 浏览器访问 `http://服务器IP:8081` |
| B | Web | 8000 | 8000 | 浏览器访问 `http://服务器IP:8000` |
| C | 远程桌面 RDP | 3389 | 33890 | `mstsc /v:服务器IP:33890` |

规则：

1. 同一台 frps 上的 `remote_port` 不能重复。
2. RDP 不是 HTTP，不能用浏览器访问。
3. 暴露 RDP 前必须确认 Windows 账户有强密码。
4. 服务器安全组必须放行对应 `remote_port`。

---

## 【文件变更清单】

必须新增：

```text
src\MyTools\NativeBinaries\frpc.exe
src\MyTools\Services\FrpService.cs
src\MyTools\ViewModels\FrpViewModel.cs
src\MyTools\Views\FrpView.xaml
src\MyTools\Views\FrpView.xaml.cs
```

必须修改：

```text
src\MyTools\MyTools.csproj
src\MyTools\ViewModels\MainViewModel.cs
src\MyTools\MainWindow.xaml
docs\开发记录.txt
```

禁止新增：

```text
src\MyTools\ViewModels\MainViewModel.Frp.cs
```

禁止修改：

```text
src\MyTools\Services\WireGuardService.cs
src\MyTools\Services\SqlExportService.cs
src\MyTools\NativeBinaries\ffmpeg\*
```

---

## 【执行步骤】

### 步骤 1：下载并嵌入 frpc.exe

1. 下载：
   ```text
   https://github.com/fatedier/frp/releases/download/v0.50.0/frp_0.50.0_windows_amd64.zip
   ```
2. 解压后取出：
   ```text
   frpc.exe
   ```
3. 放入：
   ```text
   src\MyTools\NativeBinaries\frpc.exe
   ```
4. 修改 `src\MyTools\MyTools.csproj`。
5. 在现有 `<ItemGroup>` 中添加：
   ```xml
   <None Remove="NativeBinaries\frpc.exe" />
   <EmbeddedResource Include="NativeBinaries\frpc.exe" />
   ```
6. 不要添加普通 `<None Include="NativeBinaries\frpc.exe">`。
7. 不要让 `frpc.exe` 复制到输出目录。

自检标准：

```text
Release 输出目录中不应出现独立的 frpc.exe。
MyTools.exe 体积会增加约 10-13 MB。
```

---

### 步骤 2：新增 FrpService.cs

路径：

```text
src\MyTools\Services\FrpService.cs
```

命名空间：

```csharp
namespace MyTools.Services
```

必须使用的 using：

```csharp
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
```

#### 2.1 数据模型

在同一文件中定义：

```csharp
public enum FrpState
{
    Stopped,
    Starting,
    Running,
    Error
}

public sealed class FrpServerConfig
{
    public string ServerAddress { get; set; } = string.Empty;
    public int ServerPort { get; set; } = 7000;
    public string EncryptedToken { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
}

public sealed class FrpTunnelRule : INotifyPropertyChanged
{
    private string _name = string.Empty;
    private string _type = "tcp";
    private int _localPort;
    private int _remotePort;
    private string _description = string.Empty;
    private bool _isEnabled = true;

    public event PropertyChangedEventHandler PropertyChanged;

    public string Name
    {
        get => _name;
        set { if (_name == value) return; _name = value ?? string.Empty; OnPropertyChanged(); }
    }

    public string Type
    {
        get => _type;
        set { if (_type == value) return; _type = string.IsNullOrWhiteSpace(value) ? "tcp" : value; OnPropertyChanged(); }
    }

    public int LocalPort
    {
        get => _localPort;
        set { if (_localPort == value) return; _localPort = value; OnPropertyChanged(); }
    }

    public int RemotePort
    {
        get => _remotePort;
        set { if (_remotePort == value) return; _remotePort = value; OnPropertyChanged(); }
    }

    public string Description
    {
        get => _description;
        set { if (_description == value) return; _description = value ?? string.Empty; OnPropertyChanged(); }
    }

    public bool IsEnabled
    {
        get => _isEnabled;
        set { if (_isEnabled == value) return; _isEnabled = value; OnPropertyChanged(); }
    }

    private void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
```

说明：

- `ClientId` 用于生成全局唯一代理名，避免多台电脑使用同一端口时代理名冲突。
- `Name` 可以为空；为空时由服务自动生成。
- `Type` 第一版只允许 `tcp`。
- `FrpTunnelRule` 必须实现 `INotifyPropertyChanged`，否则 UI 中勾选/取消启用规则后，`CanStart` 与 `PublicAddressPreview` 不会及时刷新。

#### 2.2 FrpService 静态类

必须实现以下成员：

```csharp
public static class FrpService
{
    public static string ConfigPath { get; }
    public static string RulesPath { get; }

    public static Task<string> EnsureFrpcExtractedAsync();
    public static string BuildFrpcIni(FrpServerConfig config, string plainToken, IEnumerable<FrpTunnelRule> rules);
    public static string EncryptToken(string plainText);
    public static string DecryptToken(string cipherText);
    public static Task SaveConfigAsync(FrpServerConfig config);
    public static Task<FrpServerConfig> LoadConfigAsync();
    public static Task SaveRulesAsync(IEnumerable<FrpTunnelRule> rules);
    public static Task<List<FrpTunnelRule>> LoadRulesAsync();
    public static bool IsValidPort(int port);
}
```

#### 2.3 文件路径

```csharp
ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.frpconfig.json");
RulesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.frprules.json");
```

临时运行目录：

```csharp
Path.Combine(Path.GetTempPath(), "MyTools")
```

解压目标：

```csharp
%TEMP%\MyTools\frpc.exe
```

临时配置：

```csharp
%TEMP%\MyTools\frpc.ini
```

#### 2.4 EnsureFrpcExtractedAsync

行为要求：

1. 创建 `%TEMP%\MyTools`。
2. 如果 `%TEMP%\MyTools\frpc.exe` 已存在且大小大于 `1024 * 1024`，直接返回路径。
3. 从嵌入资源中读取 `frpc.exe`。
4. 资源名优先使用：
   ```text
   MyTools.NativeBinaries.frpc.exe
   ```
5. 如果资源名找不到，遍历 `Assembly.GetExecutingAssembly().GetManifestResourceNames()`，取以 `.frpc.exe` 结尾的资源。
6. 使用异步流写入，禁止同步大文件写入。
7. 解压失败时抛出异常，异常消息必须包含「frpc.exe 解压失败」。

#### 2.5 BuildFrpcIni

输出必须是 ini 格式。

示例：

```ini
[common]
server_addr = 1.2.3.4
server_port = 7000
token = secret

[mytools_pc_8f3a2c1d_3389_33890]
type = tcp
local_ip = 127.0.0.1
local_port = 3389
remote_port = 33890
```

要求：

1. `server_addr` 来自 `FrpServerConfig.ServerAddress.Trim()`。
2. `server_port` 来自 `FrpServerConfig.ServerPort`。
3. `token` 来自 `plainToken`，不从 `EncryptedToken` 解密。
4. 只输出 `IsEnabled == true` 的规则。
5. 只允许 `Type == "tcp"`。
6. local_ip 固定为 `127.0.0.1`。
7. 端口必须在 `1-65535`。
8. 代理名生成规则：
   ```text
   mytools_{MachineName}_{ClientId前8位}_{LocalPort}_{RemotePort}
   ```
9. 代理名必须只包含字母、数字、下划线，其他字符替换为 `_`。
10. 如果没有可用规则，抛出异常，消息为：
    ```text
    请至少添加并启用一条隧道规则。
    ```

#### 2.6 Token 加密

使用 DPAPI 当前用户范围：

```csharp
ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser)
ProtectedData.Unprotect(bytes, null, DataProtectionScope.CurrentUser)
```

要求：

1. 空字符串加密后仍返回空字符串。
2. 解密失败返回空字符串，不抛出。
3. 日志中禁止出现明文 Token。

#### 2.7 配置读写

所有 JSON 读写必须异步：

- `FileStream(..., useAsync: true)`
- `StreamReader.ReadToEndAsync()`
- `StreamWriter.WriteAsync()`

序列化使用项目已有 `Newtonsoft.Json`。

异常处理：

```csharp
catch (Exception ex)
{
    AppLogService.Warning("FRP config load failed: {Msg}", ex.Message);
    return new FrpServerConfig();
}
```

不要使用：

```csharp
File.ReadAllText(...)
File.WriteAllText(...)
```

---

### 步骤 3：在 FrpService.cs 中实现 FrpProcessManager

仍放在：

```text
src\MyTools\Services\FrpService.cs
```

类签名：

```csharp
public sealed class FrpProcessManager : IDisposable
```

必须包含：

```csharp
public FrpState State { get; private set; } = FrpState.Stopped;
public string StatusMessage { get; private set; } = "未运行";
public event EventHandler StateChanged;
```

内部字段至少包含：

```csharp
private Process _process;
private string _plainTokenForSanitize = string.Empty;
```

必须实现：

```csharp
public Task StartAsync(string frpcExePath, string iniContent)
public void Stop()
public void Dispose()
```

#### StartAsync 行为

1. 如果已有 frpc 进程在运行，先调用 `Stop()`。
2. 设置：
   ```csharp
   State = FrpState.Starting;
   StatusMessage = "正在连接...";
   ```
3. 从 `iniContent` 中提取 `token = ...` 的值保存到 `_plainTokenForSanitize`，仅用于日志脱敏，不写入日志。
4. 将 `iniContent` 异步写入：
   ```text
   %TEMP%\MyTools\frpc.ini
   ```
5. 使用 `ProcessStartInfo`：
   ```csharp
   FileName = frpcExePath
   Arguments = "-c \"" + iniPath + "\""
   UseShellExecute = false
   CreateNoWindow = true
   RedirectStandardOutput = true
   RedirectStandardError = true
   StandardOutputEncoding = Encoding.UTF8
   StandardErrorEncoding = Encoding.UTF8
   ```
6. 启动后调用：
   ```csharp
   BeginOutputReadLine();
   BeginErrorReadLine();
   ```
7. 逐行读取 stdout/stderr。
8. 如果任一行包含以下文本，则认为连接成功：
   ```text
   login to server success
   start proxy success
   proxy added
   ```
9. 成功后设置：
   ```csharp
   State = FrpState.Running;
   StatusMessage = "已连接";
   ```
10. 如果任一行包含以下文本，则认为失败：
   ```text
   login to server failed
   authorization failed
   port unavailable
   EOF
   ```
11. 失败后设置：
    ```csharp
    State = FrpState.Error;
    StatusMessage = "连接失败：" + SanitizeLogLine(line);
    ```
12. `StartAsync` 最多等待 15 秒确认成功或失败；15 秒内没有明确失败且进程未退出，则设置为 Running，状态文字为：
    ```text
    已启动，等待服务器确认
    ```
13. 进程退出时，如果当前不是 `Stopped`，设置为：
    ```csharp
    State = FrpState.Error;
    StatusMessage = "frpc 已退出";
    ```

#### Stop 行为

1. 如果进程存在且未退出，先 `CloseMainWindow()`。
2. 等待 1500 ms。
3. 如果仍未退出，调用 `Kill()`。
4. 最终设置：
   ```csharp
   State = FrpState.Stopped;
   StatusMessage = "未运行";
   ```

#### 日志脱敏

`SanitizeLogLine` 必须：

1. 删除 Token 明文。
2. 截断超过 300 字符的行。
3. 不记录完整 ini 内容。

---

### 步骤 4：新增 FrpViewModel.cs

路径：

```text
src\MyTools\ViewModels\FrpViewModel.cs
```

命名空间：

```csharp
namespace MyTools.ViewModels
```

必须实现：

```csharp
public sealed class FrpViewModel : INotifyPropertyChanged, IDisposable
```

必须使用的 using：

```csharp
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using MyTools.Services;
```

#### 4.1 字段

```csharp
private readonly MainViewModel _owner;
private readonly FrpProcessManager _manager = new FrpProcessManager();
private string _frpServerAddress = string.Empty;
private int _frpServerPort = 7000;
private string _frpToken = string.Empty;
private string _statusHint = "填写服务器信息和端口规则后启动隧道。";
private FrpTunnelRule _draftRule;
private bool _isLoadingConfig;
private string _clientId = string.Empty;
```

#### 4.2 构造函数

```csharp
public FrpViewModel(MainViewModel owner)
```

要求：

1. 只创建空集合、命令对象、默认 DraftRule。
2. 禁止读取磁盘。
3. 禁止解压 frpc。
4. 禁止启动进程。
5. 订阅 `_manager.StateChanged`。

#### 4.3 属性

必须包含：

```csharp
public string FrpServerAddress { get; set; }
public int FrpServerPort { get; set; }
public string FrpToken { get; set; }
public string StatusHint { get; private set; }
public ObservableCollection<FrpTunnelRule> FrpRules { get; }
public FrpTunnelRule DraftRule { get; set; }
public FrpState ConnectionState => _manager.State;
public string ConnectionStatusText => _manager.StatusMessage;
public bool IsRunning => _manager.State == FrpState.Starting || _manager.State == FrpState.Running;
public bool CanStart => !IsRunning && HasRequiredConfig && HasEnabledRules;
public bool CanStop => IsRunning;
public bool HasEnabledRules => FrpRules.Any(r => r.IsEnabled);
public bool HasRules => FrpRules.Count > 0;
public bool HasRequiredConfig => !string.IsNullOrWhiteSpace(FrpServerAddress) && FrpServerPort >= 1 && FrpServerPort <= 65535 && !string.IsNullOrWhiteSpace(FrpToken);
public string PublicAddressPreview { get; }
```

`PublicAddressPreview` 规则：

- 无规则：`未添加隧道规则`
- 有 1 条启用规则：`服务器地址:远程端口`
- 有多条启用规则：`已启用 N 条隧道`

#### 4.4 命令

必须包含：

```csharp
public ICommand StartTunnelCommand { get; }
public ICommand StopTunnelCommand { get; }
public ICommand AddRuleCommand { get; }
public ICommand RemoveRuleCommand { get; }
public ICommand SaveConfigCommand { get; }
public ICommand LoadConfigCommand { get; }
```

命令类型：

- `StartTunnelCommand`：`AsyncRelayCommand(StartTunnelAsync, () => CanStart)`
- `StopTunnelCommand`：`RelayCommand(StopTunnel, () => CanStop)`
- `AddRuleCommand`：`RelayCommand(AddRule, CanAddDraftRule)`
- `RemoveRuleCommand`：`RelayParameterCommand(RemoveRule, p => p is FrpTunnelRule)`
- `SaveConfigCommand`：`AsyncRelayCommand(SaveConfigAsync, () => !_isLoadingConfig)`
- `LoadConfigCommand`：`AsyncRelayCommand(LoadConfigAsync, () => !_isLoadingConfig)`

#### 4.5 LoadConfigAsync

步骤：

1. `_isLoadingConfig = true`。
2. 调用 `await FrpService.LoadConfigAsync()`。
3. 如果 `ClientId` 为空，生成：
   ```csharp
   Guid.NewGuid().ToString("N")
   ```
4. 解密 Token：
   ```csharp
   FrpToken = FrpService.DecryptToken(config.EncryptedToken);
   ```
5. 调用 `await FrpService.LoadRulesAsync()`。
6. 清空 `FrpRules`，加入读取到的规则。
7. 如果没有规则，保持空列表，不自动添加默认规则。
8. `_isLoadingConfig = false`。
9. 调用 `NotifyAll()` 和 `CommandManager.InvalidateRequerySuggested()`。

#### 4.6 SaveConfigAsync

步骤：

1. 校验服务器地址非空。
2. 校验服务器端口 `1-65535`。
3. 取当前 `_clientId`；如果为空，生成并持有。
4. 保存 `FrpServerConfig`，其中：
   ```csharp
   EncryptedToken = FrpService.EncryptToken(FrpToken)
   ```
5. 保存 `FrpRules`。
6. 设置 `StatusHint = "配置已保存。"`。
7. 日志只记录服务器地址、端口、规则数量，不记录 Token。

#### 4.7 AddRule

校验：

1. `DraftRule.LocalPort` 必须在 `1-65535`。
2. `DraftRule.RemotePort` 必须在 `1-65535`。
3. `DraftRule.RemotePort` 在当前 `FrpRules` 中不能重复。
4. `DraftRule.Type` 固定为 `tcp`。

成功后：

1. 将新规则加入 `FrpRules`。
2. 重置 DraftRule：
   ```csharp
   new FrpTunnelRule { Type = "tcp", IsEnabled = true }
   ```
3. 刷新 `HasRules`、`HasEnabledRules`、`CanStart`、`PublicAddressPreview`。

#### 4.8 StartTunnelAsync

步骤：

1. 调用 `await SaveConfigAsync()`，确保配置先落盘。
2. 调用 `await FrpService.EnsureFrpcExtractedAsync()`。
3. 构造 `FrpServerConfig`，其中 `EncryptedToken` 可为空，因为生成 ini 使用明文 `FrpToken`。
4. 调用：
   ```csharp
   var ini = FrpService.BuildFrpcIni(config, FrpToken, FrpRules);
   ```
5. 调用：
   ```csharp
   await _manager.StartAsync(frpcExePath, ini);
   ```
6. 刷新状态属性。

#### 4.9 Dispose

必须：

1. 取消订阅 `_manager.StateChanged`。
2. 调用 `_manager.Dispose()`。
3. 不使用 `.Wait()` 或 `.Result`。

---

### 步骤 5：新增 FrpView.xaml

路径：

```text
src\MyTools\Views\FrpView.xaml
```

根节点：

```xml
<UserControl x:Class="MyTools.Views.FrpView"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
             xmlns:viewModels="clr-namespace:MyTools.ViewModels">
```

DataContext 由 `MainWindow.xaml` 传入，不在 FrpView 内设置。

#### 5.1 页面布局

使用：

```xml
<Grid Margin="16">
```

行结构：

1. 标题区：`Auto`
2. 主内容：`*`
3. 底部操作区：`Auto`

#### 5.2 标题区

标题区参考 `MultimediaPage.xaml` 的标题样式：

- 外层 `Border` 使用 `CornerRadius="{StaticResource CornerRadiusCard}"`。
- 背景使用当前项目已有标题渐变：
  ```xml
  BrushHeaderGradientStartColor
  BrushHeaderGradientEndColor
  ```
- 图标使用：
  ```xml
  Kind="VpnKey"
  ```
- 标题：
  ```text
  隧道穿透
  ```
- 副标题：
  ```text
  frp 内网穿透 · 指定本机端口映射到公网服务器
  ```
- 右侧状态 Pill 绑定：
  ```xml
  Text="{Binding ConnectionStatusText}"
  ```

#### 5.3 服务器配置卡片

使用：

```xml
<materialDesign:Card Style="{StaticResource SectionCardOutlined}">
```

字段：

| 控件 | 绑定 | 说明 |
|---|---|---|
| TextBox | `FrpServerAddress` | 服务器公网 IP 或域名 |
| TextBox | `FrpServerPort` | frps 监听端口，默认 7000 |
| PasswordBox | 代码后置写入 `FrpToken` | Token |
| Button | `SaveConfigCommand` | 保存配置 |

Token 输入框必须使用 `PasswordBox`，不要用普通 `TextBox`。

提示文字：

```text
Token 已使用 Windows DPAPI 加密保存，仅当前 Windows 用户可解密。
```

#### 5.4 添加规则卡片

字段：

| 控件 | 绑定 | 说明 |
|---|---|---|
| TextBox | `DraftRule.LocalPort` | 本机端口，如 3389 |
| TextBox | `DraftRule.RemotePort` | 服务器公网端口，如 33890 |
| TextBox | `DraftRule.Description` | 说明，如 远程桌面 |
| Button | `AddRuleCommand` | 添加规则 |

说明文字必须包含：

```text
服务器安全组必须放行远程端口；远程端口不能与其他电脑重复。
```

#### 5.5 规则列表

使用 `ItemsControl ItemsSource="{Binding FrpRules}"`。

每行显示：

1. 启用复选框：`IsEnabled`
2. `127.0.0.1:{LocalPort}`
3. 箭头 `→`
4. `{FrpServerAddress}:{RemotePort}`
5. 描述
6. 删除按钮：`RemoveRuleCommand`

如果没有规则，显示：

```text
暂无隧道规则，请先添加本地端口和远程端口。
```

空状态 Visibility 使用 `BoolToVis` 或新建局部反向样式，不使用 `Hidden`。

#### 5.6 底部操作区

按钮：

1. 启动隧道：
   ```xml
   Command="{Binding StartTunnelCommand}"
   IsEnabled="{Binding CanStart}"
   ```
2. 停止隧道：
   ```xml
   Command="{Binding StopTunnelCommand}"
   IsEnabled="{Binding CanStop}"
   ```

同时显示：

```xml
Text="{Binding PublicAddressPreview}"
Text="{Binding StatusHint}"
```

#### 5.7 行数限制

`FrpView.xaml` 必须少于 800 行。

---

### 步骤 6：新增 FrpView.xaml.cs

路径：

```text
src\MyTools\Views\FrpView.xaml.cs
```

必须使用：

```csharp
using System.Windows.Controls;
using MyTools.ViewModels;
```

实现：

```csharp
namespace MyTools.Views
{
    public partial class FrpView : UserControl
    {
        public FrpView()
        {
            InitializeComponent();
        }

        private void FrpTokenBox_OnPasswordChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is FrpViewModel viewModel && sender is PasswordBox box)
            {
                viewModel.FrpToken = box.Password;
            }
        }
    }
}
```

如果希望加载已保存 Token 后同步到 PasswordBox，可追加 `Loaded` 事件：

```csharp
private void FrpView_OnLoaded(object sender, System.Windows.RoutedEventArgs e)
{
    if (DataContext is FrpViewModel viewModel && FrpTokenBox.Password != viewModel.FrpToken)
    {
        FrpTokenBox.Password = viewModel.FrpToken ?? string.Empty;
    }
}
```

注意：

- 不要在代码后置设置 `DataContext`。
- 不要在代码后置读取配置文件。
- 不要在代码后置启动 frpc。

---

### 步骤 7：修改 MainViewModel.cs

路径：

```text
src\MyTools\ViewModels\MainViewModel.cs
```

#### 7.1 添加字段

在现有字段：

```csharp
private MultimediaViewModel _multimedia;
```

附近添加：

```csharp
private FrpViewModel _frp;
```

#### 7.2 添加命令初始化

在构造函数命令初始化区，现有代码附近：

```csharp
ShowSqlExportCommand = new RelayCommand(() => { SwitchModule("SqlExport"); Refresh(); });
ShowCodexProfilesCommand = new RelayCommand(() => SwitchModule("CodexProfiles"));
ShowSystemInfoCommand = new RelayCommand(() => ShowSystemSection("SystemInfo"));
ShowFileVerifyCommand = new RelayCommand(() => SwitchModule("FileVerify"));
ShowMultimediaCommand = new RelayCommand(() => ShowMultimedia(MultimediaPreferredFilter.All));
```

添加：

```csharp
ShowFrpCommand = new RelayCommand(() => SwitchModule("Frp"));
```

建议放在 `ShowSqlExportCommand` 后面或 `ShowMultimediaCommand` 后面。

#### 7.3 添加命令属性

在现有命令属性：

```csharp
public ICommand ShowSqlExportCommand { get; }
public ICommand ShowCodexProfilesCommand { get; }
public ICommand ShowMultimediaCommand { get; }
```

附近添加：

```csharp
public ICommand ShowFrpCommand { get; }
```

#### 7.4 添加懒加载属性

参考 `Schedule`、`SystemSettings`、`Multimedia` 属性，在同一区域添加：

```csharp
public FrpViewModel Frp
{
    get
    {
        if (_frp == null)
        {
            _frp = new FrpViewModel(this);
            OnPropertyChanged();
        }

        return _frp;
    }
}
```

#### 7.5 延后加载配置

在 `ScheduleStartupBackgroundLoads()` 的 `ApplicationIdle` 块内添加：

```csharp
SafeFireAndForget(Frp.LoadConfigAsync());
```

必须放在已有 `SafeFireAndForget(...)` 同级位置，例如：

```csharp
SafeFireAndForget(LoadCodexProfilesAsync());
SafeFireAndForget(LoadOptimizationReportsAsync());
SafeFireAndForget(LoadWeChatRootsAsync());
SafeFireAndForget(LoadRecentWeChatBackupsAsync());
SafeFireAndForget(Frp.LoadConfigAsync());
```

不要在构造函数中直接调用 `Frp.LoadConfigAsync().Wait()` 或 `.Result`。

#### 7.6 SwitchModule 最近使用记录

在 `SwitchModule(string module)` 中追加：

```csharp
else if (string.Equals(module, "Frp", StringComparison.Ordinal))
{
    AddHomeRecentItem("隧道穿透", "最近进入穿透模块", module, string.Empty, "打开");
}
```

#### 7.7 Dispose

在 `Dispose()` 中，建议放在 `_sensorService = null;` 后、`CloseOwnedWindowsForShutdown();` 前：

```csharp
_frp?.Dispose();
```

不要在 Dispose 中等待异步任务。

---

### 步骤 8：修改 MainWindow.xaml

路径：

```text
src\MyTools\MainWindow.xaml
```

#### 8.1 NavItemStyle 添加 ActiveFrp

在 `<Style x:Key="NavItemStyle" TargetType="Button">` 的 `ControlTemplate.Triggers` 中追加：

```xml
<Trigger Property="Tag" Value="ActiveFrp">
    <Setter TargetName="Root" Property="Background" Value="{DynamicResource BrushSidebarActive}" />
    <Setter Property="Foreground" Value="{DynamicResource BrushSidebarTextActive}" />
    <Setter TargetName="ActiveBar" Property="Visibility" Value="Visible" />
    <Setter TargetName="ActiveBar" Property="Fill" Value="{DynamicResource BrushAccentAmber}" />
</Trigger>
```

不要破坏现有 `ActiveMultimedia` 与 `ActiveCodexProfiles` 触发器。

#### 8.2 左侧导航添加按钮

建议放在 `SQL 导出` 后、`多媒体` 前：

```xml
<Button Command="{Binding ShowFrpCommand}">
    <Button.Style>
        <Style TargetType="Button" BasedOn="{StaticResource NavItemStyle}">
            <Style.Triggers>
                <DataTrigger Binding="{Binding CurrentModule}" Value="Frp">
                    <Setter Property="Tag" Value="ActiveFrp" />
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Button.Style>
    <StackPanel Orientation="Horizontal">
        <materialDesign:PackIcon Kind="VpnKey" Width="17" Height="17" Margin="0,0,9,0" VerticalAlignment="Center" Foreground="{DynamicResource BrushAccentAmber}" />
        <TextBlock Text="穿透" VerticalAlignment="Center" FontSize="12" Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}" />
    </StackPanel>
</Button>
```

#### 8.3 主内容区添加 FrpView

在主内容 `Grid Grid.Column="1"` 中，参考 `MultimediaPage` 和 `SchedulePage` 的写法添加：

```xml
<views:FrpView DataContext="{Binding Frp}">
    <views:FrpView.Style>
        <Style TargetType="views:FrpView">
            <Setter Property="Visibility" Value="Collapsed" />
            <Style.Triggers>
                <DataTrigger Binding="{Binding DataContext.CurrentModule, RelativeSource={RelativeSource AncestorType=mah:MetroWindow}}" Value="Frp">
                    <Setter Property="Visibility" Value="Visible" />
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </views:FrpView.Style>
</views:FrpView>
```

关键点：

- `FrpView` 自己的 DataContext 是 `FrpViewModel`。
- 显示/隐藏触发器必须从 `mah:MetroWindow` 取主 DataContext 的 `CurrentModule`。
- 默认必须是 `Collapsed`，不能是 `Hidden`。

---

### 步骤 9：修改 MyTools.csproj

路径：

```text
src\MyTools\MyTools.csproj
```

当前项目是 SDK 风格，不要添加：

```xml
<Compile Include="ViewModels\FrpViewModel.cs" />
<Page Include="Views\FrpView.xaml" />
```

只需要添加 frpc 资源嵌入：

```xml
<ItemGroup>
  <None Remove="NativeBinaries\frpc.exe" />
  <EmbeddedResource Include="NativeBinaries\frpc.exe" />
</ItemGroup>
```

可以放在现有 SQLite `EmbeddedResource` 的 `<ItemGroup>` 后面。

---

### 步骤 10：更新 docs/开发记录.txt

路径：

```text
docs\开发记录.txt
```

在文件末尾追加：

```markdown
## [YYYY-MM-DD] 新增 frp 隧道穿透模块

- 新增 `Services/FrpService.cs`：负责 frpc.exe 嵌入资源解压、INI 配置生成、DPAPI Token 加解密、配置/规则异步持久化、frpc 进程启动与停止。
- 新增 `ViewModels/FrpViewModel.cs`：负责服务器配置、端口规则列表、启动/停止命令、连接状态与公网访问地址预览。
- 新增 `Views/FrpView.xaml(.cs)`：新增「穿透」页面，支持服务器地址、端口、Token、端口映射规则和一键启停。
- 修改 `MainWindow.xaml`：新增左侧「穿透」导航按钮和 `views:FrpView` 内容面板，默认 `Collapsed`。
- 修改 `MainViewModel.cs`：新增 `Frp` 懒加载属性、`ShowFrpCommand`、启动空闲期配置加载、退出时释放 frpc 进程。
- 修改 `MyTools.csproj`：将 `NativeBinaries\frpc.exe` 从 None 项移除并作为 `EmbeddedResource` 嵌入，保持单 exe 分发。
- 新增嵌入资源 `NativeBinaries\frpc.exe`（frp v0.50.0 windows_amd64，约 10-13 MB），用途为 frp 客户端内网穿透；单 exe 体积预计增加约 10-13 MB。
```

日期使用执行当天。

---

### 步骤 11：构建验证

执行：

```powershell
dotnet build src\MyTools\MyTools.csproj -c Release --no-incremental
```

必须满足：

1. 输出包含：
   ```text
   已成功
   ```
2. 没有 `error`。
3. Release 产物存在：
   ```text
   src\MyTools\bin\Release\net48\MyTools.exe
   ```
4. Release 目录中不应出现：
   ```text
   frpc.exe
   ```

---

## 【输入说明】

执行 AI 将收到：

1. 本说明书全文。
2. 当前项目源码目录 `c:\exe`。
3. 阿里云服务器信息由用户在 MyTools UI 中填写，不硬编码在代码中。

示例用户输入：

```text
服务器地址：1.2.3.4
服务器端口：7000
Token：用户自己填写
本地端口：3389
远程端口：33890
说明：远程桌面
```

---

## 【输出要求】

最终代码必须满足：

1. MyTools 左侧出现「穿透」入口。
2. 点击「穿透」显示 `FrpView`。
3. 可以添加至少 1 条 TCP 端口映射规则。
4. Token 使用 PasswordBox 输入。
5. Token 落盘使用 DPAPI 加密。
6. 点击「启动隧道」后启动 `%TEMP%\MyTools\frpc.exe -c %TEMP%\MyTools\frpc.ini`。
7. 点击「停止隧道」后终止 frpc 进程。
8. 软件退出时释放 frpc 进程。
9. 日志不包含明文 Token。
10. 构建 Release 成功。

明确禁止：

1. 禁止引入新 NuGet 包。
2. 禁止将 Token 明文写入 JSON。
3. 禁止将 frpc.ini 写入程序目录。
4. 禁止在 `MainViewModel` 构造函数中读取 frp 配置。
5. 禁止使用 `.Wait()` 或 `.Result`。
6. 禁止修改 WireGuard 模块逻辑。
7. 禁止使用 TOML 作为默认配置格式。
8. 禁止让 `frpc.exe` 作为零散文件复制到 Release 输出目录。

---

## 【边界与限制】

1. 如果 `frpc.exe` 下载失败，则标注 `[待确认] frpc.exe 下载失败`，不要伪造二进制文件。
2. 如果用户要求使用 frp v0.61.0，则必须提示：该版本可能不兼容 Windows 7，需用户确认放弃 Windows 7 兼容后才能改用 TOML。
3. 如果 `remote_port` 被服务器占用，frpc 会失败，UI 显示 `连接失败：...`，不要自动换端口。
4. 如果服务器安全组未放行远程端口，frpc 可能显示已连接，但外部访问失败；UI 中必须有文字提示用户检查阿里云安全组。
5. 如果本机端口没有服务监听，frpc 可能仍能启动，但外部连接会失败；UI 中必须提示用户确认本机服务已运行。
6. 如果本机休眠、关机或断网，外部访问会中断；这是正常限制，不要写成 bug。
7. 如果多台电脑同时连接同一 frps，必须保证每条规则的远程端口不同。
8. 如果 `FrpView.xaml` 超过 800 行，必须拆分为更小的 UserControl 或在 `docs/开发记录.txt` 写明原因。

---

## 【示例】

### 正例：客户端 INI

```ini
[common]
server_addr = 1.2.3.4
server_port = 7000
token = strong-token

[mytools_pc01_8f3a2c1d_3389_33890]
type = tcp
local_ip = 127.0.0.1
local_port = 3389
remote_port = 33890
```

### 反例：默认使用 TOML

```toml
serverAddr = "1.2.3.4"
serverPort = 7000
auth.token = "strong-token"
```

原因：默认方案要求兼容 Windows 7，因此使用 frp v0.50.0 + ini。

### 正例：MainWindow 中 FrpView 的 DataContext

```xml
<views:FrpView DataContext="{Binding Frp}">
```

### 反例：让 FrpView 继承 MainWindow DataContext 后再用 Frp 前缀

```xml
<views:FrpView>
```

然后在内部写：

```xml
Text="{Binding Frp.FrpServerAddress}"
```

原因：当前项目已有 `MultimediaPage`、`SchedulePage` 均使用子 ViewModel 作为页面 DataContext，新模块应保持一致。

### 正例：ApplicationIdle 延后加载

```csharp
SafeFireAndForget(Frp.LoadConfigAsync());
```

### 反例：构造函数同步加载

```csharp
Frp.LoadConfigAsync().Wait();
```

原因：违反启动期规则和 `.Wait()` 禁令。

---

## 【自检清单】

- [ ] `MainViewModel` 未改成 partial，未新增 `MainViewModel.Frp.cs`
- [ ] `FrpViewModel` 构造函数没有磁盘 IO、网络 IO、进程启动
- [ ] `ScheduleStartupBackgroundLoads()` 中通过 `SafeFireAndForget(Frp.LoadConfigAsync())` 延后加载配置
- [ ] `FrpService` 的 JSON 读写使用异步 FileStream
- [ ] Token 使用 DPAPI 当前用户范围加密
- [ ] 日志不包含明文 Token
- [ ] `frpc.exe` 作为 EmbeddedResource 嵌入
- [ ] `MyTools.csproj` 对 `NativeBinaries\frpc.exe` 使用了 `<None Remove=...>`
- [ ] Release 输出目录没有零散 `frpc.exe`
- [ ] `MainWindow.xaml` 中 `FrpView` 默认 `Collapsed`
- [ ] `FrpView` 的 `DataContext="{Binding Frp}"`
- [ ] 显示/隐藏触发器通过 `mah:MetroWindow` 绑定 `CurrentModule`
- [ ] 远程端口重复时禁止添加规则
- [ ] Stop/Dispose 能终止 frpc 进程
- [ ] `docs\开发记录.txt` 已说明 frpc.exe 体积增长和用途
- [ ] Release 构建成功
