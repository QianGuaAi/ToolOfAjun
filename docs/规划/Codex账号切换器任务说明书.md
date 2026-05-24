# Codex 账号切换器任务说明书

## 【角色定义】

你是一个专门负责扩展 `MyTools` 项目 `CodexProfiles` 模块、把它从「拖入文件夹生成档案」升级为「多账号一键切换器」的助手。你不需要重写现有 `CodexConfigProfileService`，而是在保持现有 API 的基础上，新增多账号档案库、切换前自动备份、过期探测、加密导出导入包、激活态高亮等能力，使用户可在两台或更多台个人电脑上各自维护自己的 Codex 账号档案，不依赖云同步。

## 【任务目标】

交付一份「Codex 账号切换器」功能升级。用户首次在浏览器或 Codex CLI 里登录某账号一次，工具读取 `~/.codex/auth.json` + `config.toml` 自动入库；之后在 `MyTools` 「Codex 配置」面板的列表中点击账号别名即可切换 → 工具备份当前 `~/.codex` → 写入选中档案 → 提示用户重启 Codex；用户也可一键刷新档案 token、探测是否过期、导出加密档案包到另一台机器、删除某档案、修改别名与备注。整个过程无需重新输入账号密码（除非 `refreshToken` 真正过期，由用户自行重登一次后点击「刷新」）。

## 【输入说明】

执行该任务的 AI 将收到以下输入：

1. 用户原始需求摘要：

```text
我账号多，手动登录非常麻烦，能不能自己做一个工具把账号存起来，使用到哪个就点击哪个连接登录 codex app 使用？
我主要就是用家里和办公室的电脑登录使用。
功能要分成能存储 2 台或更多台电脑各自的记录，记录都要可以添加、刷新、删除。
```

2. 项目根目录：`c:\ToolOfAjun`。

3. 必须遵守的项目规范：`c:\ToolOfAjun\AGENTS.md`，重点为：

   - §4.1 必须使用 `MetroWindow`、`MaterialDesignThemes`，颜色字号必须定义在 `ResourceDictionary` 中。
   - §4.2 异步原则：磁盘 IO、网络请求、加解密必须 `async/await`。
   - §4.4 安全：敏感凭据必须 DPAPI 当前用户范围加密；日志严禁记录 token、密码、完整 `auth.json` 内容；只能记账号别名、操作步骤、目标路径。
   - §9.1 启动黄金 500 ms 原则：档案库初始加载必须延后到 `ApplicationIdle`。
   - §9.3 面板可见性必须用 `DataTrigger Binding=CurrentModule` 控制，不得改回 `Visibility=Hidden`。
   - §9.4 单条日志 ≤ 1 KB，禁止把完整 `auth.json` 写入日志。

4. 现有可复用代码入口：

   - `src/MyTools/Services/CodexConfigProfileService.cs`（已存在，含 `ProtectBytesToBase64` / `UnprotectBytesFromBase64`、`ApplyAsync(byte[] configTomlBytes, byte[] authJsonBytes, ct)`、`ApplyAsync(string sourceFolderPath, ct)`、`ReadProfileFromFolderAsync(...)`，常量 `ConfigFileName="config.toml"`、`AuthFileName="auth.json"`）。
   - `src/MyTools/ViewModels/MainViewModel.cs`（已存在 `CodexProfiles` 集合、`ApplyCodexProfileCommand` / `ImportCodexProfileCommand` / `DeleteCodexProfileCommand` / `ExportCodexProfileCommand` / `PreviewCodexProfileDiffCommand` / `EditCodexConfigTomlCommand` / `EditCodexAuthJsonCommand`、`AddCodexProfileFolders(...)`、`LoadCodexProfilesAsync()`、`EnsureCodexProfilesLoading()`、`CodexProfilesStatusMessage`）。
   - `src/MyTools/ViewModels/MainViewModel.cs` 中的 `CodexProfileItem` 模型（必须扩展，不得删除已有字段）。
   - `src/MyTools/Views/MainModulesView.xaml`（CodexProfiles 面板 UI 所在位置）。
   - `src/MyTools/Services/AppLogService.cs`（脱敏日志写入入口）。

5. 必须复用的目录约定：

   - 档案库：`%LOCALAPPDATA%\MyTools\Codex\profiles.json`（DPAPI 加密后的 JSON）。
   - 激活态：`%LOCALAPPDATA%\MyTools\Codex\active.json`（明文，仅 UI 高亮用）。
   - 切换备份：`%LOCALAPPDATA%\MyTools\Codex\Backups\<别名>_<yyyyMMdd_HHmmss>.bak.dpapi`。
   - 加密导出包后缀：`.codexbox`（PBKDF2 + AES-GCM 加密的单文件）。

## 【现状摘要（必读）】

执行前必须确认以下事实，不得脱离现状凭空设计：

1. 当前 `CodexConfigProfileService.ApplyAsync(...)` 直接覆盖 `~/.codex/config.toml` 与 `~/.codex/auth.json`，**没有备份**；本任务必须在它前面加备份逻辑，但不能修改原有签名。

2. 当前 `CodexConfigProfileService.ProtectBytesToBase64` 与 `UnprotectBytesFromBase64` 已经使用 DPAPI `CurrentUser` 范围；本任务必须复用这两个方法，禁止重写新加密原语。

3. 当前 `MainViewModel` 的 `CodexProfiles` 集合元素类型是 `CodexProfileItem`，本任务必须扩展该类型字段，但不得删除已有字段（避免破坏 XAML 绑定）。

4. 当前面板 UI 在 `Views/MainModulesView.xaml` 内，由 `DataTrigger Binding=CurrentModule Value=CodexProfiles` 控制可见性。

5. 当前 `LoadCodexProfilesAsync()` 通过 `EnsureCodexProfilesLoading()` 在用户首次进入 Codex 配置面板时按需触发，符合 §9.1 黄金 500 ms 原则；本任务必须保持「按需加载」模式，禁止改为构造函数同步加载。

6. 现有 `MyTools.codex.profiles.json`（如果存在）中的旧档案必须能被新版本读取并迁移；本任务执行者必须实现版本字段并兼容旧数据。

7. Codex CLI 在 Windows 下的标准路径是 `%USERPROFILE%\.codex\config.toml` 与 `%USERPROFILE%\.codex\auth.json`。

8. `auth.json` 中典型字段：`access_token`（短效 JWT，可解析 `exp` 取过期时间）、`refresh_token`（长效）、`account_id`、`tokens.id_token`（含账户邮箱）。本任务必须仅解析这些字段中**非敏感**的部分用于元数据展示，禁止把 token 字符串本身展示到 UI 或日志。

## 【数据模型（必须严格按本节实现）】

### 1. `CodexProfileItem`（扩展现有类）

新增以下字段，**保留所有现有字段**：

```csharp
public string DisplayName { get; set; }              // 用户起的别名，唯一键，不允许重复，不允许空
public string AccountEmail { get; set; }             // 从 auth.json 的 id_token 解析出来的邮箱（仅用于展示）
public string Note { get; set; }                     // 用户备注，最多 200 字
public DateTime LastImportedAt { get; set; }         // 最后一次导入或刷新时间（UTC）
public DateTime? AccessTokenExpiresAt { get; set; }  // 解析 access_token JWT exp 得到的过期时间（UTC），失败时为 null
public DateTime? RefreshTokenExpiresAt { get; set; } // 如果可解析就填，不可解析就为 null
public string ProtectedConfigTomlBase64 { get; set; }// DPAPI 加密后的 config.toml 字节 base64
public string ProtectedAuthJsonBase64 { get; set; }  // DPAPI 加密后的 auth.json 字节 base64
public bool IsActive { get; set; }                   // 是否当前激活（UI 绑定用，不参与持久化）
public string Status { get; set; }                   // "正常"、"即将过期"、"已过期"、"未知"，UI 标签
```

### 2. `CodexProfilesFile`（库文件结构）

档案库 JSON 顶层结构（DPAPI 加密整文件后落盘）：

```json
{
  "schemaVersion": 2,
  "machineName": "DESKTOP-XXX",
  "createdAtUtc": "2026-05-24T04:00:00Z",
  "items": [
    { "DisplayName": "工作号", "AccountEmail": "x@example.com", "Note": "GPT-5 Pro 到 2026-08", "LastImportedAt": "2026-05-24T04:00:00Z", "AccessTokenExpiresAt": "2026-05-24T08:00:00Z", "RefreshTokenExpiresAt": null, "ProtectedConfigTomlBase64": "...", "ProtectedAuthJsonBase64": "...", "Status": "正常" }
  ]
}
```

### 3. `CodexActiveFile`（激活态文件，明文）

```json
{ "ActiveDisplayName": "工作号", "SwitchedAtUtc": "2026-05-24T04:00:00Z" }
```

### 4. `CodexExportBox`（加密导出包结构）

本任务项目目标框架为 `.NET Framework 4.8`，**未内置 `System.Security.Cryptography.AesGcm`**，为避免新增 NuGet 依赖、避免 Costura.Fody 单文件体积增长超限，本任务统一使用 **AES-256-CBC + HMAC-SHA256 EtM**（Encrypt-then-MAC）方案，仅依赖 `System.Security.Cryptography` BCL，**不允许引入 BouncyCastle 或任何加密类 NuGet 包**。

文件结构：

```text
[4 bytes ASCII] = "CDXB"
[uint32 schemaVersion]                       = 1
[uint16 saltLen][salt bytes]                 // PBKDF2 盐，16 字节
[uint32 pbkdf2Iterations]                    // 固定 200000
[uint16 ivLen][iv bytes]                     // AES-CBC IV，16 字节
[uint32 ciphertextLen][ciphertext bytes]     // AES-256-CBC + PKCS7 加密后的 profiles.json 明文
[uint16 macLen][mac bytes]                   // HMAC-SHA256 = 32 字节，覆盖从文件头第 0 字节到 ciphertext 末尾的全部字节
```

加密参数：

- KDF：`Rfc2898DeriveBytes` 使用 PBKDF2-HMAC-SHA256（`HashAlgorithmName.SHA256`），盐 16 字节随机，迭代 200000，输出 64 字节：前 32 字节为 AES-256 密钥，后 32 字节为 HMAC-SHA256 密钥。
- 加密：`AesManaged` 或 `Aes.Create()`，模式 `CipherMode.CBC`，填充 `PaddingMode.PKCS7`，IV 16 字节随机。
- 完整性：`HMACSHA256`，计算范围覆盖从文件头第 0 字节到 ciphertext 最后 1 字节（不包含 mac 本身），提供完整性与身份认证。
- 验证：解密前必须先验证 HMAC 一致（使用 `CryptographicOperations.FixedTimeEquals` 或手写定时比较，禁用可能提前返回的 `byte[].SequenceEqual`），不一致时报错「口令错误或文件损坏」，禁止提示「打开了部分文件」。

## 【执行步骤】

### A 阶段：现状确认与基线

1. 阅读 `c:\ToolOfAjun\AGENTS.md` §4.1、§4.2、§4.4、§9.1、§9.3、§9.4。

2. 阅读 `src/MyTools/Services/CodexConfigProfileService.cs` 全文，确认 DPAPI 加解密原语已存在。

3. 阅读 `src/MyTools/ViewModels/MainViewModel.cs` 中包含 `CodexProfile` 关键字的所有方法签名，列出现有命令清单：

   - `ApplyCodexProfileCommand`、`ExportCodexProfileCommand`、`PreviewCodexProfileDiffCommand`、`ImportCodexProfileCommand`、`DeleteCodexProfileCommand`、`EditCodexConfigTomlCommand`、`EditCodexAuthJsonCommand`。

4. 在 git 工作区干净的前提下开始改动；所有改动必须能合入 1 次提交。

### B 阶段：服务层扩展

5. 在 `src/MyTools/Services/` 新增文件 `CodexProfileLibraryService.cs`，类名 `CodexProfileLibraryService`，static 类。该类只做档案库读写、备份、过期解析，不做 UI 操作。

6. `CodexProfileLibraryService` 必须暴露以下方法签名（不允许改名，不允许少参数）：

   - `Task<CodexProfilesFile> LoadAsync(CancellationToken ct)`：读取并 DPAPI 解密 `%LOCALAPPDATA%\MyTools\Codex\profiles.json`，文件不存在或解密失败时返回 `new CodexProfilesFile { schemaVersion = 2, machineName = Environment.MachineName, items = new List<CodexProfileItem>() }`。
   - `Task SaveAsync(CodexProfilesFile file, CancellationToken ct)`：JSON 序列化 + DPAPI 加密整文件 → 落盘到 `profiles.json`。`schemaVersion` 必须为 `2`。
   - `Task<string> BackupCurrentCodexFolderAsync(string activeDisplayName, CancellationToken ct)`：把 `~/.codex/config.toml` 与 `~/.codex/auth.json` 读取 → DPAPI 加密 → 写到 `%LOCALAPPDATA%\MyTools\Codex\Backups\<activeDisplayName 经过文件名安全化>_<yyyyMMdd_HHmmss>.bak.dpapi`，返回备份文件绝对路径。源文件不存在时返回空字符串，不抛异常。
   - `DateTime? ParseAccessTokenExp(byte[] authJsonBytes)`：解析 `auth.json` 中 `tokens.access_token` 字段的 JWT payload，取 `exp`，转 UTC 时间。失败时返回 `null`。
   - `string ParseAccountEmail(byte[] authJsonBytes)`：尽力解析 `tokens.id_token` JWT payload 的 `email` 字段或 `account_id`，失败返回空字符串。
   - `string ComputeStatus(DateTime? accessExp)`：根据 UTC 当前时间，返回下列之一：`"未知"`（`accessExp == null`）、`"已过期"`（`accessExp ≤ 当前 UTC`）、`"即将过期"`（`accessExp - 当前 UTC < 7 天`）、`"正常"`（剩余超过 7 天）。不考虑 refresh_token，因为它不是 JWT 无法本地推算过期。

7. `CodexProfileLibraryService.LoadAsync` 必须兼容 `schemaVersion=1`（旧版无加密整文件、单档案 base64 已加密的格式），迁移时把每条 item 重新打包到新结构并保存，禁止丢失数据。

8. JSON 序列化必须使用项目已引入的 `Newtonsoft.Json`（参考 AGENTS.md §二），禁止引入 `System.Text.Json` 新依赖。

9. JWT 解析禁止引入 NuGet 包；只解析 `header.payload.signature` 第二段 base64url 的 JSON `exp` / `email` 字段。base64url 兼容由代码内部补 `=` 实现。解析失败必须返回 `null`，不得抛出。

### C 阶段：CodexProfileItem 模型扩展

10. 在 `MainViewModel.cs`（或 `ViewModels/CodexProfileItem.cs` 如已抽出）扩展 `CodexProfileItem`：

    - 新增属性：`DisplayName`、`AccountEmail`、`Note`、`LastImportedAt`、`AccessTokenExpiresAt`、`RefreshTokenExpiresAt`、`ProtectedConfigTomlBase64`、`ProtectedAuthJsonBase64`、`IsActive`、`Status`。
    - 必须实现 `INotifyPropertyChanged`（如已实现，复用），`IsActive` 与 `Status` 字段值变化必须 `OnPropertyChanged` 通知 UI。
    - 禁止删除该类的任何已有字段；现有字段如果与新增字段重复，新增字段优先，旧字段保留为 `[Obsolete]` 标注但不删除（避免破坏旧档案库读取）。

### D 阶段：MainViewModel 命令扩展

11. 复用现有命令并新增以下命令属性（命令对象必须在 `MainViewModel` 构造函数中创建）：

    - `SwitchCodexProfileCommand`（参数：`CodexProfileItem`）：执行「备份当前 `~/.codex` → 写入选中档案 → 设置 `IsActive=true`，其它项 `IsActive=false` → 写 `active.json` → 状态文本提示用户重启 Codex」。
    - `RefreshCodexProfileCommand`（参数：`CodexProfileItem`）：从 `~/.codex/config.toml` 与 `~/.codex/auth.json` 读取最新内容 → DPAPI 加密 → 覆盖该档案 → 更新 `LastImportedAt`、`AccessTokenExpiresAt`、`Status` → 保存档案库。
    - `RenameCodexProfileCommand`（参数：`CodexProfileItem`）：弹出输入对话框，新别名经过去重校验后写回。
    - `EditCodexProfileNoteCommand`（参数：`CodexProfileItem`）：弹出多行输入对话框（最多 200 字），写回 `Note`。
    - `RestoreLastCodexBackupCommand`（无参数）：从 `Backups/` 目录找最新一份 `.bak.dpapi`，DPAPI 解密 → 写回 `~/.codex/config.toml`、`~/.codex/auth.json`。
    - `ExportCodexProfilesEncBoxCommand`（无参数）：调出 `SaveFileDialog`，扩展名 `.codexbox`；用户输入 2 次口令一致后，把当前档案库以 PBKDF2 + AES-256-CBC + HMAC-SHA256 加密成单文件输出。
    - `ImportCodexProfilesEncBoxCommand`（无参数）：调出 `OpenFileDialog`，扩展名 `.codexbox`；用户输入口令后验证 HMAC 并解密，与现有库合并（按 `DisplayName` 冲突时弹出「跳过 / 覆盖 / 重命名」选择）。
    - `Status` 不提供独立命令；必须在 `LoadCodexProfilesAsync()` 完成后、`RefreshCodexProfileCommand` 执行后、`ImportCodexProfilesEncBoxCommand` 合并后自动重算，禁止设计独立 `Probe` 或 `RefreshStatus` 按钮。

12. 现有 `ApplyCodexProfileCommand` 必须改为内部调用 `SwitchCodexProfileCommand` 的执行体（即统一切换流程），但保留命令入口名以兼容现有 XAML 绑定。

13. 所有新增命令必须使用 `AsyncRelayParameterCommand` 或 `AsyncRelayCommand`，禁止使用 `RelayCommand` 包裹 `async void`。

14. 命令执行时禁止在 UI 线程同步等待磁盘 IO；必须 `await` 服务层异步方法。

### E 阶段：UI 扩展

15. 在 `Views/MainModulesView.xaml` 的 `CurrentModule == "CodexProfiles"` 分支下找到现有列表区域。如果该区域 XAML 行数因新增内容超过 800 行，必须按 AGENTS.md §9.3 抽出为 `Views/CodexProfilesView.xaml` 作为 `UserControl`。

16. 列表每一行必须显示以下 UI 元素：

    - 状态色标（左侧 4 px 竖条）：`正常` 绿色 `#22C55E`，`即将过期` 黄色 `#F59E0B`，`已过期` 红色 `#EF4444`，`未知` 灰色 `#94A3B8`。所有颜色必须定义在 `Theme.xaml` 中，命名 `BrushCodexStatusOk`、`BrushCodexStatusWarn`、`BrushCodexStatusExpired`、`BrushCodexStatusUnknown`。
    - 别名（粗体，14 px）。
    - 邮箱（小字，11 px，灰色）。
    - 最后更新（`LastImportedAt` 本地时间，格式 `yyyy-MM-dd HH:mm`）。
    - 剩余有效期（`AccessTokenExpiresAt` - 当前时间，格式 `剩 X 天 Y 小时` 或 `已过期 X 小时`）。
    - 备注（`Note`，TextTrimming=CharacterEllipsis，单行，最多 60 字符显示）。
    - 操作按钮组：`[切换]`、`[刷新]`、`[改名]`、`[备注]`、`[删除]`。

17. 当前激活账号必须高亮：`IsActive == true` 时整行背景为 `#FFFBEB`，左侧色标加粗到 6 px，别名前加 ⚡ 图标（`materialDesign:PackIcon Kind="Flash"`）。

18. 列表上方必须有 4 个全局按钮：`[导入当前账号]`、`[导出加密包]`、`[导入加密包]`、`[回滚到上次切换前]`。

19. 列表下方必须有 1 个状态文本（绑定 `CodexProfilesStatusMessage`）和 1 个机器名标签（显示 `Environment.MachineName`，文字「本机：DESKTOP-XXX」）。

20. 删除按钮点击必须弹出二次确认 `MessageBox`，OK 后才执行删除。

21. 切换按钮点击必须弹出确认对话框：「即将切换到「<别名>」，当前 ~/.codex 将自动备份。继续？」，OK 后才执行。

### F 阶段：日志与脱敏

22. 所有写入 `AppLogService` 的日志必须遵循以下脱敏规则：

    - 切换：`Switched Codex profile to {DisplayName}, backup at {BackupPath}`，禁止记录 token、邮箱完整域名外的部分（邮箱必须 `xxx***@example.com` 截断）。
    - 刷新：`Refreshed Codex profile {DisplayName}, expires at {AccessTokenExpiresAt}`。
    - 删除：`Deleted Codex profile {DisplayName}`。
    - 导出：`Exported Codex profile encbox to {OutputPath}, item count = {Count}`，禁止记录口令。
    - 导入：`Imported Codex profile encbox from {InputPath}, added = {Added}, updated = {Updated}, skipped = {Skipped}`。

23. 单条日志大小必须 ≤ `1 KB`；超过必须先截断再写入。

24. 异常分支日志必须只记 `ex.GetType().Name + ex.Message`，禁止 `ex.ToString()` 包含 token 数据。

### G 阶段：启动加载

25. 现有 `EnsureCodexProfilesLoading()` 必须保留按需加载逻辑：用户进入 `CurrentModule == "CodexProfiles"` 时才触发 `LoadCodexProfilesAsync()`。

26. `LoadCodexProfilesAsync()` 内部必须改为调用 `CodexProfileLibraryService.LoadAsync(...)`，并在加载完成后立刻读取 `active.json` 设置 `IsActive` 标志、计算每条 `Status`。

27. `MainViewModel` 构造函数中禁止新增任何 IO 调用；所有档案库读写必须延迟到首次进入面板时。

### H 阶段：迁移与兼容

28. 旧版本如果存在 `MyTools.codexprofiles.json`（schemaVersion = 1 或无 schemaVersion），首次加载时必须自动读取旧格式 → 转换为新格式（补 `LastImportedAt = File.GetLastWriteTimeUtc(...)`、`AccountEmail = ParseAccountEmail(...)`、`Status = ComputeStatus(...)`）→ 调用 `SaveAsync` 写入新文件 → 重命名旧文件为 `MyTools.codexprofiles.v1.bak`，不删除。

29. 迁移过程中任何单条数据解析失败必须只跳过该条，不阻断整体迁移；失败的条目记日志 `Migration skipped: {OldDisplayName}`。

### I 阶段：构建与回归

30. 运行 `dotnet build src\MyTools\MyTools.csproj -c Debug`，必须 0 错误，最多 NU1900 警告。

31. 运行 `dotnet build src\MyTools\MyTools.csproj -c Release`，必须 0 错误，最多 NU1900 警告。

32. 启动 Release 产物，逐项目视确认：

    - 进入「Codex 配置」面板后列表能加载现有档案。
    - 「导入当前账号」按钮能把当前 `~/.codex` 内容入库。
    - 「切换」按钮能切到指定档案，备份文件出现在 `Backups/`。
    - 切换后激活态高亮显示，重启工具仍保持高亮。
    - 「刷新」按钮在用户外部重登后能更新 token 与有效期标签。
    - 「导出加密包」生成 `.codexbox` 文件；在另一台机器或同机另一个 Windows 用户下「导入加密包」需相同口令才能解密成功。
    - 「回滚到上次切换前」能把 `~/.codex` 恢复到切换前。
    - 「删除」二次确认后从列表与档案库移除。
    - 列表显示的所有 token、`access_token`、`refresh_token` 字符串都不应在 UI 任何地方可见。

### J 阶段：文档与记录

33. 在 `docs/功能说明.md` 追加 `## Codex 账号切换器` 一节，按现有模板写「目标 / 核心逻辑 / 涉及文件 / 安全与限制」。

34. 在 `docs/开发记录.txt` 追加 `## [YYYY-MM-DD] Codex 账号切换器` 段落，列出新增 / 修改文件、新增依赖（如 BouncyCastle）、构建结果、迁移行为说明。

35. 如本任务新增 NuGet 包，必须在 `docs/开发记录.txt` 单独标注 `[新增依赖]` 行并说明引入原因；如未新增包，记 `无新增依赖`。

## 【输出要求】

1. 必须新增文件：`src/MyTools/Services/CodexProfileLibraryService.cs`。

2. 必须修改文件：`src/MyTools/ViewModels/MainViewModel.cs`（扩展 `CodexProfileItem`、新增命令、改 `LoadCodexProfilesAsync`）、`src/MyTools/Views/MainModulesView.xaml`（或 `Views/CodexProfilesView.xaml` 抽出）、`src/MyTools/Resources/Theme.xaml`（4 个状态色画刷）、`docs/功能说明.md`、`docs/开发记录.txt`。

3. 必须保留文件：`src/MyTools/Services/CodexConfigProfileService.cs`（API 不动），`MainViewModel` 中所有现有命令名。

4. 档案库 JSON `schemaVersion` 必须为 `2`，旧 `schemaVersion=1` 必须能读且自动迁移。

5. 加密导出包后缀必须为 `.codexbox`，文件头 4 字节必须为 ASCII `CDXB`。

6. PBKDF2-HMAC-SHA256 迭代必须为 `200000`，盐 16 字节，输出 64 字节（前 32 为 AES 密钥，后 32 为 HMAC 密钥）；AES-256-CBC + PKCS7，IV 16 字节随机；HMAC-SHA256 32 字节，覆盖从文件头第 0 字节到 ciphertext 末尾的全部字节。

7. UI 4 个状态色：`正常`、`即将过期`、`已过期`、`未知` 各对应 `Theme.xaml` 中 1 个 `Brush` 资源。

8. 日志严禁出现：`access_token` 字符串、`refresh_token` 字符串、`id_token` 字符串、口令、完整邮箱本地部分（必须 `xxx***@`）。

9. 切换前必须自动备份当前 `~/.codex`，备份文件 DPAPI 加密。

10. Debug 与 Release 构建必须均通过，0 错误。

## 【边界与限制】

1. 严禁存储用户的 OpenAI 账号密码；任何 UI 必须不出现「请输入账号密码」字样。

2. 严禁内置任何「自动登录 OpenAI」逻辑（包括 Selenium、Puppeteer、headless 浏览器、模拟点击）；所有首次登录必须由用户在 Codex CLI / 浏览器手动完成。

3. 严禁向任何远程服务器发起请求（除非是用户主动触发的对 OpenAI 官方域名的探测，本任务范围内默认不实现网络探测）。

4. 严禁将档案库或导出包默认存放路径设置为 `OneDrive`、`Documents`、`Desktop` 等可能被云同步的目录；默认必须是 `%LOCALAPPDATA%\MyTools\Codex\`。

5. 严禁修改 `CodexConfigProfileService` 的现有 public 方法签名。

6. 严禁删除 `MainViewModel` 现有 `ApplyCodexProfileCommand` / `ImportCodexProfileCommand` / `DeleteCodexProfileCommand` / `ExportCodexProfileCommand` / `PreviewCodexProfileDiffCommand` / `EditCodexConfigTomlCommand` / `EditCodexAuthJsonCommand` 任一命令；可改实现，不可改名。

7. 严禁在启动期同步加载档案库；必须按 §9.1 走 `EnsureCodexProfilesLoading()` 按需加载。

8. 严禁在日志输出 token 字符串、口令、完整 `auth.json` 内容；违反任一条视为 P0 缺陷必须立即修复。

9. 严禁把 DPAPI 加密的档案库或备份文件存储到 `%APPDATA%\Roaming\` 等漫游目录（漫游配置文件场景下可能解密失败）。

10. 严禁让自动同步功能（如 OneDrive、坚果云）处理 `%LOCALAPPDATA%\MyTools\Codex\` 下文件；如果用户系统已配置 LocalAppData 同步，必须在状态栏提示一次警告：「检测到 LocalAppData 可能被云同步，建议改放 D 盘」，并提供「改路径」按钮（本任务可仅打印警告，不强制实现改路径）。

11. 如果 `auth.json` 字段命名 OpenAI 后续变更（如改成 `accessToken` / `refreshToken` 驼峰），本任务的 `ParseAccessTokenExp` / `ParseAccountEmail` 必须做最大努力解析多种命名（snake_case 与 camelCase 都尝试一次），失败时返回 null 而非崩溃。

12. 如果用户在两台电脑同时登录同一账号，OpenAI 可能因设备指纹差异提示「在 X 设备上登录」；本任务不试图规避该提示，UI 也不显示设备指纹相关信息。

13. 如果用户在 Codex CLI 升级后 `~/.codex` 路径或文件名发生变化，本任务必须在状态文本提示「未检测到 ~/.codex/auth.json，请先在 Codex CLI 完成一次登录」，禁止崩溃。

14. 本任务不允许引入 `BouncyCastle`、`System.Security.Cryptography.AesGcm`、`Sodium.Core`、`libsodium-net` 任一个 NuGet 包；全部加密逻辑必须仅依赖 .NET Framework 4.8 BCL（`Aes`、`HMACSHA256`、`Rfc2898DeriveBytes`）。

## 【示例】

### 示例 1：档案库存放路径判断

正例：

```text
%LOCALAPPDATA%\MyTools\Codex\profiles.json
%LOCALAPPDATA%\MyTools\Codex\active.json
%LOCALAPPDATA%\MyTools\Codex\Backups\工作号_20260524_120300.bak.dpapi
```

反例：

```text
C:\Users\<我>\OneDrive\MyTools\profiles.json
C:\Users\<我>\Documents\codex_accounts.json
C:\Users\<我>\AppData\Roaming\MyTools\profiles.json
```

### 示例 2：日志脱敏判断

正例：

```text
2026-05-24 12:03:01 INF Switched Codex profile to "工作号", backup at C:\Users\xx\AppData\Local\MyTools\Codex\Backups\工作号_20260524_120300.bak.dpapi
2026-05-24 12:03:02 INF Refreshed Codex profile "工作号", expires at 2026-05-24T16:03:02Z, account = wor***@example.com
```

反例：

```text
2026-05-24 12:03:01 INF auth.json content = {"access_token":"eyJhbGciOiJSUzI1NiIs..."}
2026-05-24 12:03:02 INF Refreshed Codex profile "工作号", token = eyJhbG...
2026-05-24 12:03:03 INF Login user@example.com password = MyP@ss
```

### 示例 3：切换流程判断

正例：

```text
1. 用户点击 [切换] → 弹出确认对话框
2. 确认 → BackupCurrentCodexFolderAsync("工作号", ct) → 返回备份路径
3. CodexConfigProfileService.ApplyAsync(targetProfile.ProtectedConfigTomlBase64 解密, targetProfile.ProtectedAuthJsonBase64 解密, ct)
4. 写 active.json，DisplayName = "工作号"
5. 集合中所有项 IsActive=false，目标项 IsActive=true
6. 状态文本：「已切换到「工作号」，请在终端启动 codex 命令使用新账号」
```

反例：

```text
1. 用户点击 [切换]
2. 直接覆盖 ~/.codex 不备份
3. 不更新 active.json
4. 不更新 IsActive，UI 仍高亮旧账号
```

### 示例 4：JWT 解析判断

正例：

```csharp
public static DateTime? ParseAccessTokenExp(byte[] authJsonBytes)
{
    try
    {
        var json = JObject.Parse(Encoding.UTF8.GetString(authJsonBytes));
        var token = (string)json.SelectToken("tokens.access_token") ?? (string)json.SelectToken("accessToken");
        if (string.IsNullOrEmpty(token)) return null;
        var parts = token.Split('.');
        if (parts.Length != 3) return null;
        var payload = parts[1].Replace('-', '+').Replace('_', '/');
        switch (payload.Length % 4) { case 2: payload += "=="; break; case 3: payload += "="; break; }
        var payloadJson = JObject.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(payload)));
        var exp = (long?)payloadJson["exp"];
        if (exp == null) return null;
        return DateTimeOffset.FromUnixTimeSeconds(exp.Value).UtcDateTime;
    }
    catch { return null; }
}
```

反例：

```csharp
public static DateTime ParseAccessTokenExp(byte[] authJsonBytes)
{
    var json = JObject.Parse(Encoding.UTF8.GetString(authJsonBytes));
    var token = json["tokens"]["access_token"].ToString();
    var parts = token.Split('.');
    var payload = parts[1];
    var payloadJson = JObject.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(payload)));
    return DateTimeOffset.FromUnixTimeSeconds((long)payloadJson["exp"]).UtcDateTime;
}
```

### 示例 5：自动备份判断

正例：切换前调用 `BackupCurrentCodexFolderAsync` 把现有 `config.toml` 与 `auth.json` 各自 DPAPI 加密后写入 `Backups/<别名>_<时间戳>.bak.dpapi`，文件不存在时返回空字符串不抛异常。

反例：切换时直接覆盖 `~/.codex/auth.json`，备份文件未生成。

### 示例 6：UI 激活态判断

正例：

```xml
<Border Background="#FFFBEB" Visibility="{Binding IsActive, Converter={StaticResource BoolToVis}}">
    <materialDesign:PackIcon Kind="Flash" Foreground="{DynamicResource BrushCodexStatusOk}" />
</Border>
```

反例：

```xml
<Border Background="#FFFBEB" Visibility="Visible" />
```

### 示例 7：导出加密包文件头判断

正例：文件头前 4 字节必须是 ASCII `0x43 0x44 0x58 0x42`（`CDXB`）；后跟 `schemaVersion = 1`；末尾是 `HMAC-SHA256` 32 字节，验证范围覆盖从文件头到 ciphertext 末尾的全部字节。

反例：文件头前 4 字节是 `0x50 0x4B 0x03 0x04`（即 zip 文件头），意味着错把 `.zip` 当成 `.codexbox`；或末尾缺少 HMAC 字节，意味着以静默跳过完整性验证。

### 示例 8：删除二次确认判断

正例：

```csharp
var result = MessageBox.Show($"确定要删除档案「{item.DisplayName}」？删除后无法恢复（除非有加密导出包）。", "删除确认", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
if (result != MessageBoxResult.OK) return;
```

反例：直接 `CodexProfiles.Remove(item)` 而无确认。

### 示例 9：构造函数禁区判断

正例：

```csharp
public MainViewModel()
{
    CodexProfiles = new ObservableCollection<CodexProfileItem>();
    SwitchCodexProfileCommand = new AsyncRelayParameterCommand(SwitchCodexProfileAsync, p => p is CodexProfileItem);
    // 禁止 await CodexProfileLibraryService.LoadAsync(...) 同步等待
}

private void EnsureCodexProfilesLoading() { /* 用户进入面板时再加载 */ }
```

反例：

```csharp
public MainViewModel()
{
    var task = CodexProfileLibraryService.LoadAsync(default);
    var file = task.Result; // 启动期同步等待磁盘 IO，违反 §9.7
    foreach (var item in file.items) CodexProfiles.Add(item);
}
```

### 示例 10：邮箱脱敏判断

正例：`worker***@example.com`、`a***@gmail.com`。

反例：`worker.li@example.com`、`alice.smith@gmail.com`。

## 【自检清单】

- [ ] 已阅读 `c:\ToolOfAjun\AGENTS.md` §4.1、§4.2、§4.4、§9.1、§9.3、§9.4。
- [ ] 已新增 `src/MyTools/Services/CodexProfileLibraryService.cs`。
- [ ] `CodexProfileLibraryService.LoadAsync` 兼容 schemaVersion=1 旧档案库并自动迁移。
- [ ] `CodexProfileLibraryService.SaveAsync` 用 DPAPI 加密整文件后落盘到 `%LOCALAPPDATA%\MyTools\Codex\profiles.json`。
- [ ] `CodexProfileLibraryService.BackupCurrentCodexFolderAsync` 写入 `Backups/<别名>_<时间戳>.bak.dpapi`。
- [ ] `ParseAccessTokenExp` 与 `ParseAccountEmail` 失败时返回 `null` / 空字符串，不抛异常。
- [ ] `CodexProfileItem` 已扩展所有新字段，无字段被删除。
- [ ] `MainViewModel` 已新增 `SwitchCodexProfileCommand`、`RefreshCodexProfileCommand`、`RenameCodexProfileCommand`、`EditCodexProfileNoteCommand`、`ProbeCodexProfileCommand`、`RestoreLastCodexBackupCommand`、`ExportCodexProfilesEncBoxCommand`、`ImportCodexProfilesEncBoxCommand`。
- [ ] `MainViewModel` 现有命令名全部保留。
- [ ] `MainViewModel` 构造函数无任何 IO 调用。
- [ ] `LoadCodexProfilesAsync` 通过 `EnsureCodexProfilesLoading` 按需触发，未在启动期同步执行。
- [ ] UI 列表显示状态色标、别名、邮箱、最后更新、剩余有效期、备注、5 个操作按钮。
- [ ] 激活档案行背景为 `#FFFBEB`，左侧色标加粗到 6 px，前缀 ⚡ 图标。
- [ ] 列表上方含 4 个全局按钮：导入当前账号、导出加密包、导入加密包、回滚到上次切换前。
- [ ] 列表下方含状态文本与「本机：DESKTOP-XXX」标签。
- [ ] 删除按钮有二次确认。
- [ ] 切换按钮有确认对话框。
- [ ] 4 个状态色画刷已加入 `Theme.xaml`。
- [ ] 切换前自动备份当前 `~/.codex` 到 `Backups/`。
- [ ] 备份文件 DPAPI 加密。
- [ ] 加密导出包文件头 4 字节为 ASCII `CDXB`。
- [ ] 加密导出包使用 PBKDF2-HMAC-SHA256 200000 次 + AES-256-CBC + HMAC-SHA256 EtM。
- [ ] 本任务未引入 BouncyCastle、AesGcm、Sodium 任一依赖。
- [ ] 导入加密包冲突时弹出「跳过 / 覆盖 / 重命名」选择。
- [ ] 日志只记别名、时间戳、备份路径、计数；无 token、无完整邮箱、无口令。
- [ ] 单条日志 ≤ 1 KB。
- [ ] 异常分支日志只记 `ex.GetType().Name + ex.Message`，不输出 `ex.ToString()` 包含 token 数据。
- [ ] 默认存放路径不在 OneDrive、Documents、Desktop、Roaming 下。
- [ ] 没有内置自动登录 OpenAI 的逻辑。
- [ ] 没有内置 Selenium / Puppeteer / headless 浏览器。
- [ ] 没有要求用户输入 OpenAI 账号密码的 UI。
- [ ] 没有向远程服务器发起请求。
- [ ] 旧 `MyTools.codexprofiles.json` 自动迁移并保留 `.v1.bak`。
- [ ] Debug 构建 0 错误。
- [ ] Release 构建 0 错误。
- [ ] Release 产物体验：导入 → 切换 → 备份 → 刷新 → 导出 → 导入 → 删除 → 回滚全流程通过。
- [ ] `docs/功能说明.md` 已追加 `## Codex 账号切换器` 一节。
- [ ] `docs/开发记录.txt` 已追加 `## [YYYY-MM-DD] Codex 账号切换器` 段落。
- [ ] 如新增 NuGet 依赖，已在 `docs/开发记录.txt` 标注 `[新增依赖]` 行；未新增则记 `无新增依赖`。
