# frp 预设端口下拉列表任务说明书

---

## 【角色定义】

你是一个专门负责在 MyTools（.NET Framework 4.8 / WPF / MVVM）中为已有「隧道穿透」模块增加「预设端口下拉列表」功能的助手。你必须严格遵守 `c:\exe\AGENTS.md`：保持 Windows 7 SP1+ 兼容、单 exe 分发、不引入 .NET Core 或高版本运行时依赖、所有 IO 异步、启动期零重 IO、保留 Token DPAPI 加密、日志不记录敏感信息。

---

## 【任务目标】

在 `FrpView.xaml`「添加规则」卡片顶部新增一个「常用预设」下拉列表（`ComboBox`），用户选中条目后立即把该预设的本机端口、公网端口、备注三个值填入下方的 DraftRule 输入框，再让用户点击已有的「添加规则」按钮提交；同时把第一次启动时的默认服务器地址改为 `120.26.50.234`，已保存非空服务器地址不覆盖。

---

## 【当前项目事实】

执行前必须先确认以下事实，不要凭旧记忆改：

1. 项目根目录是 `c:\exe`。
2. 主工程是 `c:\exe\src\MyTools\MyTools.csproj`，目标框架 `net48`，SDK 风格自动包含 `.cs` / `.xaml`，不要手动 `<Compile Include>`。
3. FRP 相关现有文件（已实现，本任务在其基础上扩展，禁止重写）：
   - `c:\exe\src\MyTools\Services\FrpService.cs`：定义 `FrpServerConfig`、`FrpTunnelRule`、`FrpService`、`FrpProcessManager`。
   - `c:\exe\src\MyTools\ViewModels\FrpViewModel.cs`：定义 `FrpViewModel`、属性 `FrpServerAddress` / `FrpServerPort` / `FrpToken` / `FrpRules` / `DraftRule` 等，命令 `AddRuleCommand` 等。
   - `c:\exe\src\MyTools\Views\FrpView.xaml`：UI 布局，「添加规则」卡片在 `Grid.Row="0" Grid.Column="2"`，绑定 `DraftRule.LocalPort` / `DraftRule.RemotePort` / `DraftRule.Description`。
   - `c:\exe\src\MyTools\Views\FrpView.xaml.cs`：含 `FrpTokenBox_OnPasswordChanged` 和 `FrpView_OnLoaded`。
4. `FrpTunnelRule` 已实现 `INotifyPropertyChanged`，三个属性 `LocalPort` / `RemotePort` / `Description` 均会在 set 时触发通知。
5. `FrpViewModel.LoadConfigAsync()` 在 `_isLoadingConfig = false` 阶段才会触发 `NotifyAll()`，配置加载默认值的最佳时机是在该方法内部完成 `FrpServerAddress` 赋值之后。
6. `AppLogService.Warning` 签名是 `Warning(string template, params object[] propertyValues)`，不存在 `Warning(Exception, ...)` 重载。
7. 现有「添加规则」卡片 XAML 区间：`c:\exe\src\MyTools\Views\FrpView.xaml:112-161`，根 `materialDesign:Card` 的 `Padding="14,12"`，内部为 `StackPanel`。
8. 项目已有 `RelayCommand` 与 `RelayParameterCommand` 类（`MyTools.ViewModels` 命名空间），不要新建命令基类。
9. `Newtonsoft.Json`、`MaterialDesignThemes` 已经引用，不要新增 NuGet 包。
10. 不要修改 `FrpService.cs` 的现有公共 API；本任务允许在该文件**追加**一个新的静态嵌套类型 `FrpPortPreset`（数据载体）。

---

## 【文件变更清单】

必须修改：

```text
c:\exe\src\MyTools\Services\FrpService.cs            （仅追加 FrpPortPreset 类型与静态预设列表）
c:\exe\src\MyTools\ViewModels\FrpViewModel.cs        （新增 PortPresets / SelectedPortPreset / ApplyPortPreset）
c:\exe\src\MyTools\Views\FrpView.xaml                （在「添加规则」卡片顶部插入 ComboBox 行）
c:\exe\docs\开发记录.txt                              （按 ## [YYYY-MM-DD] 标题追加修改条目）
```

允许修改：

```text
c:\exe\docs\功能说明.md                               （如果存在则追加一段说明，不存在则不创建）
```

禁止新增：

```text
任何独立的 FrpPortPresetService.cs / FrpPortPresetViewModel.cs
任何独立的 *.xaml 子用户控件
```

禁止修改：

```text
c:\exe\src\MyTools\Services\FrpService.cs 中已有的方法签名
c:\exe\src\MyTools\ViewModels\FrpViewModel.cs 中已有的命令与属性签名
c:\exe\src\MyTools\Views\FrpView.xaml 中已有的「服务器配置」「隧道规则」「服务器端配置说明」三张卡片的结构
c:\exe\src\MyTools.Installer\*
c:\exe\src\MyTools\NativeBinaries\*
```

---

## 【执行步骤】

### 步骤 1：在 FrpService.cs 末尾追加 FrpPortPreset 类型与预设列表

打开 `c:\exe\src\MyTools\Services\FrpService.cs`，在 `namespace MyTools.Services { ... }` 大括号内部、所有现有类型之后追加：

```csharp
public sealed class FrpPortPreset
{
    public string DisplayName { get; }
    public int LocalPort { get; }
    public int RemotePort { get; }
    public string Description { get; }

    public FrpPortPreset(string displayName, int localPort, int remotePort, string description)
    {
        DisplayName = displayName ?? string.Empty;
        LocalPort = localPort;
        RemotePort = remotePort;
        Description = description ?? string.Empty;
    }

    public override string ToString() => DisplayName;
}

public static class FrpPortPresetCatalog
{
    public static IReadOnlyList<FrpPortPreset> All { get; } = new[]
    {
        new FrpPortPreset("远程桌面 (RDP)", 3389, 33890, "Windows 远程桌面"),
        new FrpPortPreset("网页 HTTP 80", 80, 8081, "本机 80 端口网页服务"),
        new FrpPortPreset("网页开发 8080", 8080, 8082, "本机 8080 开发服务器"),
        new FrpPortPreset("网页开发 8000", 8000, 8003, "本机 8000 开发服务器"),
        new FrpPortPreset("SSH 远程", 22, 2222, "OpenSSH 服务"),
        new FrpPortPreset("MySQL 数据库", 3306, 3307, "MySQL 服务"),
        new FrpPortPreset("PostgreSQL 数据库", 5432, 5433, "PostgreSQL 服务"),
        new FrpPortPreset("Redis 缓存", 6379, 6380, "Redis 服务"),
        new FrpPortPreset("SMB 文件共享", 445, 4450, "Windows 文件共享"),
        new FrpPortPreset("VNC 远程桌面", 5900, 5901, "VNC 服务")
    };
}

public static class FrpDefaults
{
    public const string DefaultServerAddress = "120.26.50.234";
    public const int DefaultServerPort = 7000;
}
```

自检标准：

- 编译 `dotnet build c:\exe\src\MyTools\MyTools.csproj -c Release` 应为 0 error，可有 NETSDK1057 message。
- 文件总行数较改动前增加在 35–50 行区间。

---

### 步骤 2：修改 FrpViewModel.cs 暴露预设列表与"应用预设"逻辑

打开 `c:\exe\src\MyTools\ViewModels\FrpViewModel.cs`。

#### 2.1 在 `private bool _disposed;` 这一行下面追加字段：

```csharp
private FrpPortPreset _selectedPortPreset;
```

#### 2.2 在 `public ObservableCollection<FrpTunnelRule> FrpRules { get; }` 这一行下面追加只读属性：

```csharp
public IReadOnlyList<FrpPortPreset> PortPresets { get; } = FrpPortPresetCatalog.All;
```

并在文件顶部 `using` 区内已经存在 `using MyTools.Services;` 时**不要重复**添加。

#### 2.3 在 `public FrpTunnelRule DraftRule` 属性下方追加：

```csharp
public FrpPortPreset SelectedPortPreset
{
    get => _selectedPortPreset;
    set
    {
        if (ReferenceEquals(_selectedPortPreset, value))
        {
            return;
        }

        _selectedPortPreset = value;
        OnPropertyChanged();
        ApplyPortPreset(value);
    }
}

private void ApplyPortPreset(FrpPortPreset preset)
{
    if (preset == null || DraftRule == null)
    {
        return;
    }

    DraftRule.LocalPort = preset.LocalPort;
    DraftRule.RemotePort = preset.RemotePort;
    DraftRule.Description = preset.Description;
    StatusHint = "已套用预设 \"" + preset.DisplayName + "\"，请确认端口后点击「添加规则」。";
    NotifyStateProperties();
}
```

#### 2.4 在 `LoadConfigAsync()` 内部、`FrpServerAddress = config.ServerAddress ?? string.Empty;` 这一行**之后**追加：

```csharp
if (string.IsNullOrWhiteSpace(FrpServerAddress))
{
    FrpServerAddress = FrpDefaults.DefaultServerAddress;
}
```

`FrpServerPort` 那行如果当前已写 `7000` 字面值，改为 `FrpDefaults.DefaultServerPort`；如果当前写 `FrpService.IsValidPort(config.ServerPort) ? config.ServerPort : 7000`，把后面的 `7000` 改为 `FrpDefaults.DefaultServerPort`。其它不动。

自检标准：

- 配置文件 `MyTools.frpconfig.json` 不存在或 `ServerAddress` 为空时，UI 首次显示应为 `120.26.50.234`。
- 配置文件已有非空 `ServerAddress`（例如 `10.0.0.5`）时，UI 显示仍为 `10.0.0.5`，不被覆盖。

---

### 步骤 3：修改 FrpView.xaml「添加规则」卡片，在卡片顶部插入预设下拉行

定位 `c:\exe\src\MyTools\Views\FrpView.xaml` 中：

```xml
<materialDesign:Card Grid.Row="0" Grid.Column="2" Style="{StaticResource SectionCardOutlined}" Padding="14,12">
    <StackPanel>
        <TextBlock Text="添加规则" Style="{StaticResource TextSubtitle}" FontWeight="SemiBold" Margin="0,0,0,10" />
```

把这段 `<TextBlock Text="添加规则" ... Margin="0,0,0,10" />` **替换**为：

```xml
        <TextBlock Text="添加规则" Style="{StaticResource TextSubtitle}" FontWeight="SemiBold" Margin="0,0,0,8" />

        <ComboBox ItemsSource="{Binding PortPresets}"
                  SelectedItem="{Binding SelectedPortPreset, Mode=TwoWay}"
                  DisplayMemberPath="DisplayName"
                  Style="{StaticResource MaterialDesignOutlinedComboBox}"
                  materialDesign:HintAssist.Hint="选择常用预设（自动填入下方端口）"
                  materialDesign:HintAssist.IsFloating="False"
                  Height="36"
                  Margin="0,0,0,10" />
```

只插入一个 `ComboBox`，不要额外加按钮。原有「本机端口」「公网端口」「备注」三个 TextBox 和「添加规则」按钮保持不动。

如果项目里没有 `MaterialDesignOutlinedComboBox` 样式（编译报 `Resource 'MaterialDesignOutlinedComboBox' not found`），改用项目已有的样式：先在文件内全文搜索 `Style="{StaticResource ` 找到任意一处 ComboBox 在用的样式，沿用同一样式；如果整个项目没有现成 ComboBox 样式，删除 `Style` 属性并保留 `materialDesign:HintAssist.*`，使用默认样式即可，不要新建样式资源。

自检标准：

- 启动应用进入「隧道穿透」模块，「添加规则」卡片标题下方应出现一行下拉框，提示文字「选择常用预设（自动填入下方端口）」。
- 下拉里有且仅有 10 条预设（顺序：远程桌面、网页 HTTP 80、网页开发 8080、网页开发 8000、SSH、MySQL、PostgreSQL、Redis、SMB、VNC）。
- 选「远程桌面 (RDP)」后，下方三个输入框立即变为 `3389`、`33890`、`Windows 远程桌面`，状态栏文字变为「已套用预设 "远程桌面 (RDP)"，请确认端口后点击「添加规则」。」。
- 点击「添加规则」后，规则列表多出一条 `127.0.0.1:3389 → 120.26.50.234:33890 Windows 远程桌面`。

---

### 步骤 4：在 docs/开发记录.txt 顶部追加日志

打开 `c:\exe\docs\开发记录.txt`，在 `# 开发记录\n` 这一行**下方**插入：

```text
## [YYYY-MM-DD] FRP 预设端口下拉列表
- `FrpService.cs` 新增 `FrpPortPreset` 与 `FrpPortPresetCatalog`，固定 10 条常用端口预设（RDP、HTTP 80、开发 8080、开发 8000、SSH、MySQL、PostgreSQL、Redis、SMB、VNC）。
- `FrpViewModel` 暴露 `PortPresets`、`SelectedPortPreset`；选中预设后自动填入 `DraftRule.LocalPort/RemotePort/Description`，提示已套用预设。
- 新增 `FrpDefaults.DefaultServerAddress = "120.26.50.234"`；`LoadConfigAsync` 在 `ServerAddress` 为空时填入该默认值，已有配置不覆盖。
- `FrpView.xaml`「添加规则」卡片顶部插入 `ComboBox`，绑定 `PortPresets`/`SelectedPortPreset`，不新增按钮。
```

把 `YYYY-MM-DD` 替换为当天日期（年-月-日，例如 `2026-05-25`）。不要写其它日期。

---

### 步骤 5：构建与回归

依次执行（**不要并发**，每条等上一条结束）：

```powershell
Get-Process -Name MyTools -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 300
dotnet build c:\exe\src\MyTools\MyTools.csproj -c Release --no-incremental
```

期望：`已成功生成`，0 error。允许 `message NETSDK1057`。出现任何 `error` 或 `warning CS` 必须先修复再继续。

回归手测（运行 `c:\exe\src\MyTools\bin\Release\net48\MyTools.exe`）：

1. 新机首次启动：「服务器配置 → 公网 IP 或域名」框显示 `120.26.50.234`。
2. 「添加规则」卡片标题下方多出一行下拉框。
3. 下拉选「远程桌面 (RDP)」：三个 TextBox 自动变 `3389`、`33890`、`Windows 远程桌面`。
4. 改回下拉「网页 HTTP 80」：三个 TextBox 变 `80`、`8081`、`本机 80 端口网页服务`。
5. 不点添加，关闭再重开：服务器地址保持 `120.26.50.234`；规则列表保持为空。
6. 手动把服务器地址改为 `10.0.0.5`，点击「保存配置」，关闭重开：服务器地址保持 `10.0.0.5`（验证默认值不覆盖已保存配置）。

---

## 【输入说明】

执行者会接收：

- 现有源代码（`c:\exe\src\MyTools\` 下完整工程，本说明书已列出关键文件路径）。
- 用户指令：「在 frp 功能里增加端口预设下拉列表，服务器地址默认 120.26.50.234」。

执行者不需要联网下载额外文件，不需要安装额外 NuGet。

---

## 【输出要求】

提交内容必须包含且仅包含：

- `c:\exe\src\MyTools\Services\FrpService.cs`：在原文件末尾命名空间内追加 `FrpPortPreset` / `FrpPortPresetCatalog` / `FrpDefaults` 三个类型。
- `c:\exe\src\MyTools\ViewModels\FrpViewModel.cs`：新增 `_selectedPortPreset` 字段、`PortPresets` 只读属性、`SelectedPortPreset` 属性、`ApplyPortPreset` 私有方法、`LoadConfigAsync` 内默认地址兜底。
- `c:\exe\src\MyTools\Views\FrpView.xaml`：「添加规则」卡片标题下方插入一个 `ComboBox` 行。
- `c:\exe\docs\开发记录.txt`：顶部追加一条 `## [YYYY-MM-DD]` 记录。
- `dotnet build ... -c Release` 结果为 0 error。

严禁出现：

- 新增任何 `*.csproj`、`*.cs`、`*.xaml` 文件。
- 新增任何 NuGet 包。
- 修改 `FrpService.cs` 已有方法签名。
- 修改 `FrpViewModel.cs` 已有命令、已有属性的 set 行为。
- 修改 `FrpView.xaml` 中「服务器配置」、「隧道规则」、「服务器端配置说明」三张卡片。
- 在 `App.OnStartup` / `MainWindow` 构造函数 / `MainViewModel` 构造函数中追加任何调用。
- 任何同步 IO、`.Wait()`、`.Result`、`Dispatcher.Invoke` 同步等待。
- 任何明文记录 Token 的日志语句。

---

## 【边界与限制】

| 情形 | 处理 |
|---|---|
| 用户配置文件 `MyTools.frpconfig.json` 已存在且 `ServerAddress = "120.26.50.234"` | 不变；UI 显示 `120.26.50.234`。 |
| 用户配置文件已存在且 `ServerAddress = "10.0.0.5"` | 保留 `10.0.0.5`，**不**回退为默认地址。 |
| 用户配置文件不存在 | UI 首次显示 `120.26.50.234`，端口 `7000`，Token 空。 |
| 选择预设后 `DraftRule.RemotePort` 与已有规则重复 | 不在 `ApplyPortPreset` 中拦截；让用户点击「添加规则」时由现有 `CanAddDraftRule` 判定为不可用（按钮置灰），状态文字保持「已套用预设...」直到用户改端口。 |
| `MaterialDesignOutlinedComboBox` 资源不存在 | 删除 `Style` 属性，使用默认 ComboBox 样式，禁止新建样式资源。 |
| 步骤 5 编译报错 | 先回滚到本任务开始前的 git 状态，再重新按步骤 1→4 顺序逐条改，不要跳步合并。 |
| 用户要求新增预设条目 | 标注 `[待确认]` 并说明：本次说明书只交付固定 10 条；如需增删需另开任务。 |
| 用户要求把预设作用范围扩到 UDP | 标注 `[待确认]` 并说明：本说明书与现有 `BuildFrpcIni` 都假设 `Type = "tcp"`，扩 UDP 需要先改 `FrpProcessManager`。 |

---

## 【示例】

### 示例 1：正例（选预设填充 DraftRule）

操作：在「常用预设」下拉中选「SSH 远程」。  
期望：

- 「本机端口」TextBox 文本 = `22`
- 「公网端口」TextBox 文本 = `2222`
- 「备注说明」TextBox 文本 = `OpenSSH 服务`
- 状态栏文字 = `已套用预设 "SSH 远程"，请确认端口后点击「添加规则」。`
- 「添加规则」按钮可用（`CanAddDraftRule == true`，前提是 `2222` 不与现有规则重复）。

### 示例 2：反例（不要直接添加规则）

错误实现：选中预设后立刻 `FrpRules.Add(new FrpTunnelRule { ... })`。  
原因：用户可能希望微调端口（比如把 `2222` 改成 `2223`）；自动添加会绕过用户确认，且无法处理重复端口。

### 示例 3：正例（默认地址兜底）

操作：删除 `c:\exe\src\MyTools\bin\Release\net48\MyTools.frpconfig.json`，启动程序。  
期望：「公网 IP 或域名」TextBox 文本 = `120.26.50.234`。

### 示例 4：反例（不要覆盖已有地址）

错误实现：

```csharp
FrpServerAddress = FrpDefaults.DefaultServerAddress;
```

直接无条件赋值。  
原因：覆盖了用户已经保存的服务器地址，违反「已有配置不覆盖」规则。正确写法是 `if (string.IsNullOrWhiteSpace(FrpServerAddress)) FrpServerAddress = FrpDefaults.DefaultServerAddress;`。

### 示例 5：正例（XAML 插入位置）

正确：把 `<ComboBox ...>` 放在 `<TextBlock Text="添加规则" ... />` 之**下**、`<Grid>` 之**上**，仍位于 `<StackPanel>` 内。  
反例：把 `<ComboBox>` 放在 `<materialDesign:Card>` 之外或塞进 `<Grid>` 的某一行内，导致布局错位或下拉与端口 TextBox 同行。

---

## 【自检清单】

- [ ] `FrpService.cs` 末尾追加了 `FrpPortPreset`、`FrpPortPresetCatalog`、`FrpDefaults` 三个类型，无任何已有方法被改写。
- [ ] `FrpPortPresetCatalog.All` 长度 = 10，顺序与说明书一致。
- [ ] `FrpViewModel.cs` 新增 `PortPresets`（`IReadOnlyList<FrpPortPreset>`，只读属性，无 set）。
- [ ] `FrpViewModel.cs` 新增 `SelectedPortPreset`，其 set 调用 `ApplyPortPreset`。
- [ ] `ApplyPortPreset` 修改的是 `DraftRule.LocalPort/RemotePort/Description`，**不**改动 `FrpRules`。
- [ ] `LoadConfigAsync` 在 `FrpServerAddress = config.ServerAddress ?? string.Empty;` 之后判空，仅当空时填 `FrpDefaults.DefaultServerAddress`。
- [ ] `FrpView.xaml` 在「添加规则」卡片 `<StackPanel>` 内的 `TextBlock Text="添加规则"` 紧下方插入了 `ComboBox`。
- [ ] 没有新增任何 `.cs` / `.xaml` / `.csproj` 文件。
- [ ] 没有新增任何 NuGet 引用。
- [ ] `dotnet build c:\exe\src\MyTools\MyTools.csproj -c Release --no-incremental` 输出 0 error。
- [ ] 手测：新机启动地址为 `120.26.50.234`；选预设后三框自动填入；已保存非空地址不被覆盖；选预设不会自动添加规则。
- [ ] `c:\exe\docs\开发记录.txt` 顶部追加了一条 `## [YYYY-MM-DD] FRP 预设端口下拉列表` 记录。
- [ ] 日志、状态栏、提示文字均未出现 Token 明文。
