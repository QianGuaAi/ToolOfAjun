# Codex 账号自动轮换功能 - 任务执行说明书

## 【角色定义】

你是一个专门负责在 MyTools（WPF/.NET Framework 4.8）中实现 Codex 账号自动轮换功能的开发助手。

## 【任务目标】

在 MyTools 的 Codex 配置模块中，新增账号轮换功能：用户可为每个已保存的 Codex 账号档案开启/关闭"加入轮换"开关，支持手动一键轮换，以及 Phase 2/3 的自动 429 检测轮换。切换时写入 `auth.json` 和 `config.toml`；Windows 版 Codex Desktop 当前不会热加载登录态，必须提示用户完全退出并重新打开 Codex App 后生效。

## 【执行步骤】

### Phase 1（必须优先完成）

**步骤 1：在 `Services/CodexRotationService.cs` 中创建轮换核心服务**

在 `src/MyTools/Services/` 目录下新建文件 `CodexRotationService.cs`，内容如下：

1. 创建 `CodexRotationSettings` 配置类，包含 `IsEnabled`（bool，默认 false）和 `NotifyOnSwitch`（bool，默认 true）两个属性。
2. 创建 `CodexRotationResult` 结果类，包含 `Success`（bool）、`FromProfile`（string）、`ToProfile`（string）、`Message`（string）四个属性。
3. 创建静态类 `CodexRotationService`，核心方法：
   - `GetRotatingPool()`：从 `CodexProfileLibraryService` 加载当前所有账号，过滤 `EnableRotation == true` 且 `Status != "已过期"` 的记录，按 `RotationPriority` ASC + `LastAppliedAt` ASC 排序返回 `List<CodexProfileItem>`。
   - `RotateToNextAsync(CodexProfileItem current, bool notifyOnSwitch, CancellationToken ct)`：
     - 调用 `GetRotatingPool()` 获取池，排除当前账号。
     - 如果池为空，返回失败：`"没有其他已加入轮换的账号"`。
     - 选取池中第一个（已排序）。
     - 调用 `CodexProfileLibraryService.BackupCurrentCodexFolderAsync(current.DisplayName, ct)` 备份当前账号。
     - 解密新账号的 `ProtectedAuthJsonBase64` 和 `ProtectedConfigTomlBase64`（通过 `CodexConfigProfileService.UnprotectBytesFromBase64`）。
     - 调用 `CodexConfigProfileService.ApplyAsync(configBytes, authBytes, ct)` 写入新账号配置到 `~/.codex/`。
     - 更新 `CodexActiveFile`（`ActiveDisplayName` = 新账号名，`SwitchedAtUtc` = `DateTime.UtcNow`），调用 `CodexProfileLibraryService.SaveActiveAsync`。
     - 如果 `notifyOnSwitch == true`，通过托盘图标显示气泡通知：`"Codex 账号已切换\r\n{From} → {To}\r\n下次请求将使用新账号"`。
     - 返回成功结果。
     - 所有异常捕获后返回失败结果，日志记录非敏感信息。
4. `GetNextSwitchPreview(CodexProfileItem current)` 辅助方法：返回形如 `"当前：A → 切换至：B"` 的预览字符串；如果无可用目标，返回 `"无可用轮换目标"`。

**步骤 2：在 `CodexProfileLibraryService.cs` 中为序列化模型增加新字段**

找到 `CodexProfileItem` 的 JSON 序列化位置（在 `CodexProfileLibraryService.cs` 文件底部的 `NormalizeProfileItem` 方法附近）。确保 `CodexProfilesFile` 和 `CodexProfileItem` 类（已在此文件中定义）支持新增的两个属性：

1. `CodexProfileItem` 增加两个属性（注意：实际类定义在 `MainViewModel.cs` 中的嵌套类 `CodexProfileItem`，但序列化/反序列化在 `CodexProfileLibraryService.cs` 中处理）：
   - `public bool EnableRotation { get; set; }`（默认 false）
   - `public int RotationPriority { get; set; }`（默认 0）
2. 在 `NormalizeProfileItem` 方法中，添加对这两个新属性的默认值赋值（如果为 null/0）。
3. 在 `CreatePortableExportFile` → `portableFile.items.Add(...)` 处，添加 `EnableRotation = item.EnableRotation` 和 `RotationPriority = item.RotationPriority` 的映射。
4. 在 `ConvertPortableImportPackage` 中的 `CodexProfileItem` 构造处，添加同样的两个字段映射。
5. 在 `TryLoadLegacyProfilesAsync` 中的迁移代码处，同样添加这两个字段。

**步骤 3：在 `MainViewModel.cs` 的 `CodexProfileItem` 嵌套类中增加两个属性**

在 `MainViewModel.cs` 中找到 `public class CodexProfileItem : INotifyPropertyChanged`（约第 12701 行），在现有属性定义区域添加：

```csharp
private bool _enableRotation;
private int _rotationPriority;

public bool EnableRotation
{
    get => _enableRotation;
    set { if (_enableRotation == value) return; _enableRotation = value; OnPropertyChanged(); }
}

public int RotationPriority
{
    get => _rotationPriority;
    set { if (_rotationPriority == value) return; _rotationPriority = value; OnPropertyChanged(); }
}
```

同时，在 `NormalizeProfileItem` 方法中确保默认值被设置。

**步骤 4：在 `MainViewModel.cs` 中新增轮换相关命令和属性**

在 `MainViewModel` 类中：

1. 新增字段声明区域（约第 168 行附近）添加：
   ```csharp
   private string _codexNextSwitchPreview = "无可用轮换目标";
   private CodexRotationSettings _codexRotationSettings = new CodexRotationSettings();
   ```

2. 在属性声明区域添加：
   ```csharp
   public string CodexNextSwitchPreview
   {
       get => _codexNextSwitchPreview;
       private set { _codexNextSwitchPreview = value ?? "无可用轮换目标"; OnPropertyChanged(); }
   }

   public CodexRotationSettings CodexRotationSettings
   {
       get => _codexRotationSettings;
       set { _codexRotationSettings = value ?? new CodexRotationSettings(); OnPropertyChanged(); }
   }

   public bool IsCodexRotationAvailable
   {
       get
       {
           var pool = CodexProfiles?.Where(p => p != null && p.EnableRotation && p.Status != "已过期").ToList();
           return pool != null && pool.Count > 1;
       }
   }
   ```

3. 在命令初始化区域（约第 573 行附近）添加两个新命令：
   ```csharp
   _toggleCodexProfileRotationCommand = new AsyncRelayParameterCommand(ToggleCodexProfileRotationAsync, parameter => parameter is CodexProfileItem);
   ToggleCodexProfileRotationCommand = _toggleCodexProfileRotationCommand;
   _rotateToNextCodexProfileCommand = new AsyncRelayCommand(RotateToNextCodexProfileAsync, () => IsCodexRotationAvailable);
   RotateToNextCodexProfileCommand = _rotateToNextCodexProfileCommand;
   ```

4. 在 `public ICommand DeleteCodexProfileCommand { get; }` 之后添加两个新命令属性声明：
   ```csharp
   public ICommand ToggleCodexProfileRotationCommand { get; }
   public ICommand RotateToNextCodexProfileCommand { get; }
   ```

5. 在类底部添加两个新方法的实现：
   ```csharp
   private async Task ToggleCodexProfileRotationAsync(object parameter)
   {
       if (!(parameter is CodexProfileItem item)) return;
       item.EnableRotation = !item.EnableRotation;
       await SaveCodexProfilesAsync();
       OnPropertyChanged(nameof(IsCodexRotationAvailable));
       UpdateCodexNextSwitchPreview();
   }

   private async Task RotateToNextCodexProfileAsync()
   {
       var current = CodexProfiles?.FirstOrDefault(p => p != null && p.IsActive);
       if (current == null)
       {
           CodexProfilesStatusMessage = "未找到当前激活的 Codex 账号";
           return;
       }

       CodexProfilesStatusMessage = $"正在切换 Codex 账号...";
       var result = await CodexRotationService.RotateToNextAsync(
           current, CodexRotationSettings.NotifyOnSwitch, CancellationToken.None);

       if (result.Success)
       {
           current.IsActive = false;
           var next = CodexProfiles?.FirstOrDefault(p => p != null && p.DisplayName == result.ToProfile);
           if (next != null) next.IsActive = true;
           CodexProfilesStatusMessage = $"已切换：{result.FromProfile} → {result.ToProfile}";
       }
       else
       {
           CodexProfilesStatusMessage = $"轮换失败：{result.Message}";
       }
   }

   private void UpdateCodexNextSwitchPreview()
   {
       var current = CodexProfiles?.FirstOrDefault(p => p != null && p.IsActive);
       CodexNextSwitchPreview = CodexRotationService.GetNextSwitchPreview(current);
   }
   ```

6. 在 `SaveCodexProfilesAsync` 方法中，确保 `EnableRotation` 和 `RotationPriority` 被正确序列化到 `CodexProfileLibraryService.SaveAsync`。

**步骤 5：更新 `MainModulesView.xaml` 的 Codex 配置面板 UI**

在 `MainModulesView.xaml` 中找到 `DataTemplate x:Key="CodexProfilesDeferredTemplate"`（约第 2213 行开始），在账号卡片列表区域（`ItemsControl ItemsSource="{Binding CodexProfiles}"`）中，为每个账号卡片模板（`Border` 内部）增加轮换开关：

在现有按钮组（`Switch`/`Refresh`/`Rename` 等按钮所在 StackPanel）下方，添加一个新的 `StackPanel`：

```xml
<StackPanel Orientation="Horizontal" Margin="0,6,0,0">
    <CheckBox Content="加入轮换"
              IsChecked="{Binding EnableRotation, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
              VerticalAlignment="Center"
              Style="{StaticResource MaterialDesignCheckBox}"/>
    <TextBlock Text="优先级:" Margin="12,0,4,0" VerticalAlignment="Center"
               Foreground="{DynamicResource MaterialDesignBodyLight}"
               Style="{StaticResource MaterialDesignCaptionTextBlock}"/>
    <TextBox Text="{Binding RotationPriority, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
             Width="48" VerticalAlignment="Center"
             Style="{StaticResource MaterialDesignTextBox}"
             ToolTip="数字越小优先级越高"/>
</StackPanel>
```

在面板顶部导入按钮组（`ImportCodexProfileCommand` 等按钮所在区域）下方，增加轮换控制区域：

```xml
<Border Background="{DynamicResource BrushCard}" CornerRadius="4" Padding="16" Margin="0,8,0,0">
    <StackPanel>
        <TextBlock Text="账号轮换" Style="{StaticResource TextHeadline}" Margin="0,0,0,8"/>
        <StackPanel Orientation="Horizontal">
            <Button Content="立即轮换"
                    Command="{Binding RotateToNextCodexProfileCommand}"
                    IsEnabled="{Binding IsCodexRotationAvailable}"
                    Style="{StaticResource MaterialDesignOutlinedButton}"/>
            <TextBlock Text="{Binding CodexNextSwitchPreview}"
                       Margin="12,0,0,0" VerticalAlignment="Center"
                       Foreground="{DynamicResource MaterialDesignBodyLight}"/>
        </StackPanel>
        <CheckBox Content="轮换后显示托盘通知"
                  IsChecked="{Binding CodexRotationSettings.NotifyOnSwitch}"
                  Margin="0,8,0,0"
                  Style="{StaticResource MaterialDesignCheckBox}"/>
    </StackPanel>
</Border>
```

**步骤 6：更新 `docs/功能说明.md`**

在 `docs/功能说明.md` 中找到 Codex 配置相关章节，添加新功能说明。

### Phase 2（探索性，需要用户提供日志路径后实施）

**步骤 7：日志路径探索**

1. 在 `CodexRotationService` 中新增 `FindCodexLogFilesAsync()` 方法，枚举 `~/.codex/` 下的所有 `.log` 文件和子目录中的日志。
2. 如果找到日志文件，读取最新日志，搜索 "429" / "rate limit" / "quota exhausted" 等关键字。
3. 如果检测到当前 token 的 429 记录，触发 `RotateToNextAsync`。

---

## 【输入说明】

- 你将收到 Codex 配置模块的现有代码位置和相关上下文
- 涉及文件路径均以 `src/MyTools/` 为根目录
- 现有 `CodexProfileItem` 类定义在 `MainViewModel.cs` 第 12701 行附近的嵌套类中
- `CodexProfileLibraryService.cs` 中处理 JSON 序列化/反序列化
- XAML 模板在 `MainModulesView.xaml` 的 `CodexProfilesDeferredTemplate` 中

## 【输出要求】

**必须包含**：
- `Services/CodexRotationService.cs` 完整代码文件
- `MainViewModel.cs` 中所有新增的字段、属性、命令和方法的精确行号和代码
- `MainModulesView.xaml` 中新增的 XAML 片段（包含完整的上下文以便定位）
- `CodexProfileLibraryService.cs` 中所有需要修改的位置和具体改动
- 完整的编译检查（dotnet build 无错误）

**明确禁止出现**：
- 任何破坏现有功能代码的改动
- 在启动期（App.OnStartup / MainWindow 构造函数 / MainViewModel 构造函数）添加任何同步磁盘 IO
- 硬编码的文件路径（必须使用 `Path.Combine` 和 `Environment.SpecialFolder`）
- 在日志中记录完整连接字符串或密码
- 直接调用 `Task.Run(...).Result` 或 `.Wait()`

## 【边界与限制】

- 如果 `CodexProfiles` 集合为空或只有一个账号，`IsCodexRotationAvailable` 必须返回 false，"立即轮换"按钮必须禁用。
- 如果所有 `EnableRotation == true` 的账号都已过期，必须返回失败结果并提示用户。
- 轮换操作是原子性的：任何一步失败都必须回滚（如果备份成功但写入失败，需要恢复备份）。
- 如果 `CodexProfileLibraryService.SaveAsync` 抛出异常，不要吞掉，只记录非敏感的错误类型并返回失败。
- 如果遇到 `[待确认]` 的情况，标注 `[待确认]` 并说明原因，不要猜测。
- 不允许在构造函数中启动任何后台轮询任务。

## 【示例】

**正例**：`RotateToNextAsync` 正确捕获了 `BackupCurrentCodexFolderAsync` 的异常后，返回失败结果而不崩溃。
**反例**：`RotateToNextAsync` 在 `BackupCurrentCodexFolderAsync` 失败后继续执行写入，导致用户无法回滚。

**正例**：`GetRotatingPool` 使用 `Where(p => p.EnableRotation && p.Status != "已过期")` 过滤，确保不会选中过期账号。
**反例**：`GetRotatingPool` 仅按 `EnableRotation` 过滤，导致可能选中已过期的账号，用户轮换后发现新账号也无法使用。

## 【自检清单】

- [ ] `CodexRotationService.cs` 编译通过
- [ ] `MainViewModel.cs` 编译通过
- [ ] `MainModulesView.xaml` 无 XAML 解析错误
- [ ] `CodexProfileLibraryService.cs` 序列化/反序列化支持新字段
- [ ] "立即轮换"按钮在只有 0~1 个账号启用轮换时正确禁用
- [ ] 轮换后托盘通知显示正确内容
- [ ] 写入 auth.json 路径正确（`~/.codex/auth.json`），并明确提示 Codex App 需重启后生效
- [ ] 异常情况下备份可回滚
- [ ] 日志不记录敏感信息
- [ ] 不在启动路径添加任何磁盘 IO
