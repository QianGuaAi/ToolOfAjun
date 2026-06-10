# Codex 本地中转连接失败修复任务说明书

## 【角色定义】

你是一个专门负责修复 MyTools「Codex 本地中转」连接故障的助手。你的工作对象是 `C:\ToolOfAjun` 仓库中的 .NET Framework 4.8 / WPF 项目，你需要按本说明书逐步定位并修复故障，不做说明书以外的任何改动。

## 【任务目标】

修复以下故障：用户依次点击「启动本地中转」（relay 状态灯变绿）→「重启 Codex 使用中转」后，Codex App 重启成功，但提问时 Codex 显示重试 5 次连接后报错、无法得到回复；修复后用户在 Codex App 中提问必须能正常收到流式回复。

## 【执行步骤】

### 第一阶段：读取代码（只读，不修改）

1. 读取 `AGENTS.md` 全文，记住第四章开发规范和第九章性能红线。
2. 读取 `src/MyTools/Services/CodexLocalRelayService.cs` 全文（约 2100 行）。重点记录以下 8 个位置：
   - `EnsureBackgroundRelayProcessAsync`：拉起 `MyTools.exe --codex-local-relay` 隐藏后台进程，80 次 × 500 ms 健康检查。
   - `EnsureStartedInCurrentProcessAsync`：`TcpListener` 监听 `127.0.0.1:<port>`（默认端口 48176）。
   - `HandleClientAsync`：逐请求转发；`/health` 返回固定文本；鉴权失败返回 401；转发异常返回 502。
   - `IsAuthorized`：比对请求 `Authorization: Bearer <token>` 与本地 token（DPAPI 解密）。
   - `BuildRelayConfig`：生成固定到本地中转的 `config.toml`，包含 `[model_providers.mytools_local_relay.auth]`，其 `command = "powershell.exe"` 通过脚本 `ReadCodexLocalRelayToken.ps1` 读取 DPAPI 加密的本地 token，`timeout_ms = 5000`。
   - `WriteUpstreamResponseAsync`：把上游响应以 `HTTP/1.1` + `Connection: close` 写回客户端。
   - `TryBuildV1FallbackUpstreamUri`：上游非 2xx 且基址不含 `/v1` 时自动补 `/v1` 重试一次。
   - `IsLocalRelayHealthyAsync`：健康检查实现。
3. 读取 `src/MyTools/ViewModels/MainViewModel.cs` 中第 5030–5135 行附近的两个方法：「启动本地中转」对应 `CodexLocalRelayService.StartFromCurrentCodexAndProbeAsync` 调用处，「重启 Codex 使用中转」对应 `RestartCodexWithLocalRelayAsync`。
4. 读取 `src/MyTools/Services/CodexRelayTestService.cs` 全文，记录探测请求的构造方式（`stream=true` 最小请求、`/responses` 与 `/chat/completions` 路径、`/v1` 回退）。
5. 读取 `src/MyTools/Services/CodexDesktopService.cs` 全文，记录 Codex App 的关闭、进程扫描与重启逻辑。
6. 读取 `docs/程序逻辑.md` 中包含「本地中转」的全部章节（约第 227–241 行与第 457–464 行）。

### 第二阶段：收集运行证据（只读，不修改代码）

7. 用 PowerShell 执行 `Get-Content "$env:LOCALAPPDATA\MyTools\Codex\LocalRelay\local-relay.json"`，记录 `Enabled`、`Port`、`LocalBaseUrl`、`UpstreamBaseUrl`、`WireApi` 5 个字段的值。该文件不存在时，记录「设置文件不存在」并直接跳到步骤 14 分支 A。
8. 用 PowerShell 执行 `Test-NetConnection 127.0.0.1 -Port <步骤7记录的Port>`，记录 `TcpTestSucceeded` 是 True 还是 False。
9. 端口可连通时，执行 `Invoke-WebRequest "http://127.0.0.1:<Port>/health" -UseBasicParsing`，确认响应正文是否包含 `MyTools Codex local relay is running.`。
10. 读取程序输出目录（`src\MyTools\bin\Release\` 下，或用户实际运行 exe 的同级目录）中 `MyTools.log` 的最后 300 行，统计以下 3 类日志的出现次数并摘录各 2 条原文：
    - `Codex local relay upstream returned HTTP`（含状态码）；
    - `Codex local relay request failed`（含异常类型）；
    - 含 `Unauthorized` 或 `401` 的行。
11. 读取 `%USERPROFILE%\.codex\config.toml` 全文，逐项核对 4 点：根级 `model_provider` 是否等于 `"mytools_local_relay"`；`[model_providers.mytools_local_relay]` 的 `base_url` 是否等于步骤 7 的 `LocalBaseUrl`；`wire_api` 是否等于步骤 7 的 `WireApi`；是否存在 `[model_providers.mytools_local_relay.auth]` 小节。
12. 手动执行 token 读取脚本并记录结果：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\MyTools\Codex\LocalRelay\ReadCodexLocalRelayToken.ps1" "$env:LOCALAPPDATA\MyTools\Codex\LocalRelay\local-token.dpapi"
```

   成功标准：3 秒内退出且 stdout 输出非空字符串（不要把输出内容写入任何文档或日志）。失败标准：报错、输出为空或超过 5 秒未退出。
13. 查找 Codex App 自身日志（依次检查 `%USERPROFILE%\.codex\log\`、`%USERPROFILE%\.codex\sessions\`、`%USERPROFILE%\.codex\logs_2.sqlite`），摘录与「重试 5 次 / retry / stream disconnected / connection」相关的最后 5 条错误原文。

### 第三阶段：按证据分支定位根因

14. 按以下顺序逐条判断，命中第一条即确定根因分支，不再继续往下匹配：
    - **分支 A（relay 进程未存活）**：步骤 8 为 False，或步骤 9 不含预期文本。根因是后台 relay 进程未启动或中途退出。修复位置：`CodexLocalRelayService.EnsureBackgroundRelayProcessAsync` 与 `--codex-local-relay` 进程模式的生命周期（检查 `App.xaml.cs` 中 relay 进程模式的进入与退出条件）。必须实现：Codex App 发起请求时 relay 一定在监听；若 relay 进程随主程序退出而消失，需让「重启 Codex 使用中转」流程在重启 Codex 前再次执行健康检查并重新拉起。
    - **分支 B（鉴权失败）**：步骤 10 中出现 401/Unauthorized 日志，或步骤 12 失败。根因是 Codex App 没有把正确的本地 token 带给 relay（`auth.command` 执行失败、超时、输出含 BOM/换行，或用户安装的 Codex 版本不支持 `[model_providers.*.auth]` 的 `command` 形式）。修复方向按优先级：(1) 修正脚本输出（去 BOM、去换行、缩短执行时间到 3 秒内）；(2) 若确认当前 Codex 版本不支持 `auth.command`，改为在 `BuildRelayConfig` 中直接写入 `env_key` 或 `api_key` 字段传递本地 token——本地 token 仅用于 127.0.0.1 回环鉴权，写入 `config.toml` 前必须在 `docs/开发记录.txt` 中说明该 token 不是上游 key、泄露面仅限本机当前用户。
    - **分支 C（请求未到达 relay）**：步骤 10 三类日志在用户复现问题的时间段内一条都没有，且步骤 11 有任一项核对不通过。根因是 `config.toml` 固定不正确或被 Codex 启动时改写。修复位置：`BuildRelayConfig`、`IsCodexConfigPinnedToLocalRelay`、`EnsureCurrentCodexConfigPinnedAsync`，让 4 项核对全部通过且重启 Codex 后仍保持。
    - **分支 D（上游失败）**：步骤 10 中 `upstream returned HTTP` 日志的状态码 ≥ 400 且反复出现。根因是上游 NewAPI 地址、key 或 `/v1` 路径问题。修复位置：`BuildUpstreamUri`、`TryBuildV1FallbackUpstreamUri`、`SelectEffectiveUpstreamBaseUrl`；必须保证探测成功时保存的有效上游基址与真实转发使用的基址完全一致。
    - **分支 E（转发实现缺陷）**：relay 存活、鉴权通过、上游返回 2xx，但 Codex 端仍报重试 5 次。重点排查 `WriteUpstreamResponseAsync` 与 `HandleClientAsync` 对 SSE 流式响应的处理：每个 TCP 连接只处理 1 个请求并 `Connection: close`，需验证 Codex 客户端在收到 `Connection: close` 后是否把后续请求当作连接失败；以及响应头中 `Content-Length`/`Transfer-Encoding` 被剥除后正文边界是否仍然正确。修复时保持「监听仅限 127.0.0.1」「逐请求读取最新设置」两个既有行为不变。
15. 在 `docs/规划/` 下新建 `Codex本地中转连接失败修复诊断记录_<YYYY-MM-DD>.md`，写入：步骤 7–13 的全部证据（token 明文除外）、命中的分支编号、判定依据原文。

### 第四阶段：实施修复

16. 只修改步骤 14 命中分支列出的文件；单次修复涉及的 `.cs` 文件不超过 3 个。
17. 修复代码必须遵守：所有 IO 用 `async/await`；不新增 NuGet 包；不在日志中写入 token、请求体、响应正文；上游 key 与本地 token 持久化仍走 DPAPI 当前用户范围。
18. 执行构建：`dotnet build src\MyTools\MyTools.csproj -c Release`，要求 0 个 error。出现 error 时修复后重新构建，最多重试 3 次；3 次后仍失败则停止并按【边界与限制】上报。

### 第五阶段：验证

19. 运行新构建的 `MyTools.exe`，按用户原始操作路径完整复现一遍：打开 Codex 配置面板 → 点击「启动本地中转」并等待状态灯变绿 → 点击「重启 Codex 使用中转」并确认 → 等待 Codex App 重启完成。
20. 在 Codex App 中提问 1 个问题（例如「1+1等于几」），验收标准：60 秒内收到流式回复，且不出现重试 5 次后报错。
21. 重复步骤 20 共 3 次，3 次全部成功才算通过。任何 1 次失败则回到步骤 10 重新收集证据。
22. 验证 `MyTools.log` 新增内容中不含 token、请求体、响应正文。

### 第六阶段：文档同步

23. 在 `docs/开发记录.txt` 追加一节，格式 `## [YYYY-MM-DD] 修复 Codex 本地中转连接失败`，列出根因分支、修改文件、验证结果。
24. 更新 `docs/程序逻辑.md` 中「本地中转」相关描述（约第 233–237 行）：根因分支改变了哪条逻辑，就在原句处更新哪条；未改变的句子不动。
25. 若修改了 `App.OnStartup`、`MainWindow` 构造函数或 `MainViewModel` 构造函数中的任何代码，在 `docs/规划/` 下补一份启动基线记录（文件名 `启动基线_Codex本地中转修复_<YYYY-MM-DD>.md`）；未修改则不写。

## 【输入说明】

你将收到：
- 本仓库 `C:\ToolOfAjun` 的完整读写权限（含 `src/`、`docs/`）。
- 运行机上的实时文件：`%LOCALAPPDATA%\MyTools\Codex\LocalRelay\` 下的 `local-relay.json`、`local-token.dpapi`、`ReadCodexLocalRelayToken.ps1`；`%USERPROFILE%\.codex\` 下的 `config.toml`、`auth.json` 与 Codex 日志。
- 用户故障描述原文：「点击『启动本地中转』后，点击『重启codex使用中转』，当codex app启动后，总是不能正常使用，问问题总是显示试了5次链接然后就出错不能使用了。」
- 没有错误截图、没有 Codex 报错原文，错误细节必须自行从步骤 13 的日志中获取。

## 【输出要求】

必须交付：
1. 诊断记录文件 `docs/规划/Codex本地中转连接失败修复诊断记录_<YYYY-MM-DD>.md`（含证据原文与分支判定）。
2. 修改后的源码文件（仅限命中分支列出的文件）。
3. 构建结果：`dotnet build` 0 error 的输出摘录。
4. 验证结果：3 次提问全部成功的记录（每次记录提问时间与是否收到回复）。
5. `docs/开发记录.txt` 与 `docs/程序逻辑.md` 的同步更新。

明确禁止出现：
- 任何文档、日志、代码注释中出现 token 明文、上游 key 明文、完整连接字符串。
- 新增 NuGet 依赖。
- 删除或重命名 `%LOCALAPPDATA%\MyTools\Codex\Backups\` 下的任何备份。
- 改动与本故障无关的功能模块。

## 【边界与限制】

- 不能修改 Codex App 的安装文件或其内部代码，只能修改 MyTools 侧代码与生成的 `config.toml` 内容。
- 不能关闭 relay 的 127.0.0.1 回环鉴权来「修通」连接；如果鉴权机制本身是根因，按分支 B 给出的两个修复方向处理。
- 不能把监听地址从 `IPAddress.Loopback` 改为 `0.0.0.0`。
- 覆盖 `~/.codex/config.toml` 之前必须确认现有代码的备份调用（`BackupCurrentCodexFolderAsync`）仍被执行。
- 如果步骤 8–13 全部正常、无法复现故障，则执行：在诊断记录中写明「未复现」+ 全部证据，再检查分支 E，仍无结论则标注 [待确认] 并说明原因。
- 如果确认根因是用户安装的 Codex 版本不支持 `auth.command` 且 `env_key`/`api_key` 直写方案也不被该版本支持，则标注 [待确认]，写明已验证的 Codex 版本号与不支持的配置项，停止修改代码。
- 如果需要的运行时文件（如 `local-relay.json`）在执行环境中不存在且无法运行程序生成，则标注 [待确认] 并说明缺失文件路径。

## 【示例】

**判断点 1：relay 进程是否存活（步骤 8–9）**
- 正例：`Test-NetConnection` 返回 `TcpTestSucceeded : True`，且 `/health` 响应正文为 `MyTools Codex local relay is running.` → 判定 relay 存活，排除分支 A。
- 反例：`TcpTestSucceeded : True` 但 `/health` 返回 404 或其他文本 → 不能判定存活，端口可能被其他程序占用，仍属分支 A。

**判断点 2：是否鉴权失败（步骤 10、12）**
- 正例：`MyTools.log` 中出现 3 条以上 `Unauthorized` 响应记录，且手动执行脚本输出为空 → 判定分支 B。
- 反例：日志只有 1 条 401 且时间在用户点击「停止使用本地中转」之后 → 不是本故障的根因，继续匹配后续分支。

**判断点 3：请求是否到达 relay（分支 C）**
- 正例：用户 14:00 复现故障，`MyTools.log` 在 13:55–14:05 之间没有任何 relay 相关日志，且 `config.toml` 的 `base_url` 端口与 `local-relay.json` 的 `Port` 不一致 → 判定分支 C。
- 反例：日志没有 relay 记录，但 `config.toml` 四项核对全部通过 → 不能判定分支 C，应怀疑 relay 未存活，回查分支 A。

**判断点 4：修复方式选择（分支 B 方向 2）**
- 正例：在 Codex 官方文档或其 `--help` 输出中确认当前版本不支持 `auth.command`，于是改用 `env_key` 直写并在开发记录中说明 token 泄露面 → 允许。
- 反例：没有验证 Codex 版本支持情况，直接删除 `[model_providers.mytools_local_relay.auth]` 小节并把上游真实 key 明文写入 `config.toml` → 禁止，上游 key 不得离开 DPAPI 保护。

**判断点 5：验证是否通过（步骤 20–21）**
- 正例：3 次提问分别在 12 秒、9 秒、15 秒内收到完整回复 → 验证通过。
- 反例：前 2 次成功、第 3 次重试报错 → 验证不通过，回到步骤 10。

## 【自检清单】

- [ ] 已读取 `AGENTS.md`、`CodexLocalRelayService.cs`、`MainViewModel.cs` 相关方法、`CodexRelayTestService.cs`、`CodexDesktopService.cs`、`docs/程序逻辑.md` 相关章节。
- [ ] 已记录 `local-relay.json` 的 5 个字段值。
- [ ] 已完成端口连通、`/health`、`MyTools.log`、`config.toml` 四项核对、token 脚本、Codex 日志共 6 项证据收集。
- [ ] 已按 A→B→C→D→E 顺序匹配并只命中一个根因分支。
- [ ] 诊断记录文件已创建且不含任何 token 明文。
- [ ] 代码修改只涉及命中分支列出的文件，且不超过 3 个 `.cs` 文件。
- [ ] `dotnet build src\MyTools\MyTools.csproj -c Release` 结果为 0 error。
- [ ] 完整复现用户操作路径：启动本地中转 → 重启 Codex 使用中转 → 提问。
- [ ] 3 次提问全部在 60 秒内收到流式回复。
- [ ] `MyTools.log` 新增内容不含 token、请求体、响应正文。
- [ ] `docs/开发记录.txt` 已按 `## [YYYY-MM-DD] 标题` 格式追加。
- [ ] `docs/程序逻辑.md` 中被改变的逻辑句已同步更新。
- [ ] 未修改启动路径代码，或已补启动基线记录。
- [ ] 所有无法判断的点均已标注 [待确认] 并说明原因。
