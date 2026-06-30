# C++ 原生重写总体规划

日期：2026-06-30  
状态：规划稿 v1  
目标平台：Windows 10 / Windows 11  
推荐路线：C++20 + Win32 + Direct2D + DirectWrite + 少量 C 底层库

## 1. 背景与目标

现有 MyTools 是基于 .NET Framework 4.8 / WPF 的 Windows 桌面工具集。近期已经完成多轮体积裁剪，包括 FFmpeg 外置、SQL 导出工具移除、Codex 轮换与 Ollama 导入移除、硬件实时传感器移除等。下一阶段目标是用 C/C++ 原生技术重写整个程序，进一步降低体积、启动成本和运行时依赖，同时保留当前仍有价值的个人工具能力。

本规划只定义重写方向、边界和阶段，不直接替换现有 WPF 程序。正式进入实现前，需要以当前代码和活文档为准重新盘点保留功能，并为每个迁移阶段建立独立验收门禁。

## 2. 总体结论

推荐采用“新建原生主程序，分阶段迁移模块”的方式，而不是在现有 WPF 项目中逐步混入 C++。新程序使用 C++ 负责 UI、业务编排、生命周期和资源管理；底层可复用能力使用 C 或 C ABI 封装，保持边界稳定、易测试、易替换。

最终目标不是复刻旧技术栈，而是做一个新的轻量原生工具：

- 主程序：C++20。
- UI：Win32 窗口 + Direct2D 绘制 + DirectWrite 文字。
- 系统调用：Windows API / COM / WMI 按需延迟调用。
- 加密：DPAPI 与 Windows CNG/BCrypt。
- 配置：JSON 文件 + schema version + 原子写入。
- 依赖：优先源码静态链接，禁止引入 Qt、Electron、.NET Runtime、JVM 等大型运行时。
- 分发：优先单 exe；无法单 exe 时，安装目录文件必须可解释、数量可控。

## 3. 新边界

### 3.1 平台边界

- 只支持 Windows 10 与 Windows 11。
- 建议第一阶段基线定为 Windows 10 22H2 x64 与 Windows 11 23H2/24H2 x64。
- 如需要兼容更早 Windows 10 版本，必须在使用 Windows Graphics Capture、现代任务栏 API、暗色标题栏 API 等能力前单独确认最低 build。
- 建议默认只发布 x64。x86 仅在用户明确需要时再做第二目标。
- 正式进入 Native 实现前，需要同步更新 `AGENTS.md` 中的旧 Win7/.NET Framework/WPF 边界；在更新前，本规划只作为新原生程序的目标方案，不改变现有 WPF 项目的维护规则。

### 3.2 功能边界

以下已裁剪功能不迁回原生版本：

- SQL 导出工具、SQL 查询、SQL 连接历史。
- Codex 立即轮换、全选轮换、反选轮换、轮换后通知、轮换池优先级。
- 导入本地 Ollama 模型并自动生成本地模型档案。
- 硬件实时传感器轮询。
- 内置 `ffmpeg.exe`。

以下能力优先保留并重写：

- 主窗口、托盘、菜单式导航、状态栏。
- FRP 隧道穿透。
- Codex 档案管理、切换、测试中转、复制、改名、备注、差异、config/auth、加密包导入导出。
- Codex 本地中转服务。
- 截图、区域截图、窗口截图、录屏、录音、热键。
- 多媒体浏览、图片/PDF 转换、音视频预览，FFmpeg 继续作为外置依赖。
- 排班管理、Excel 导入导出、自动排休、冲突提示。
- 微信清理、备份、恢复。
- 系统优化、垃圾清理、启动项管理、程序卸载、网络信息、系统设置、系统备份。
- 应用日志、配置、DPAPI、崩溃记录、安装器/卸载器。

## 4. 技术选型

### 4.1 UI

首选：Win32 + Direct2D + DirectWrite。

原因：

- 体积小，不依赖大型 UI 运行时。
- 对 Win10/Win11 原生支持稳定。
- 可完全控制启动路径、绘制成本、DPI、主题和控件行为。
- 适合个人工具集的长期维护和极简分发目标。

不采用：

- Qt：开发效率高，但依赖体积较大。
- Electron：体积和运行时开销不符合目标。
- WebView2 作为主界面：引入 HTML/JS UI 栈，不符合全部 C/C++ 重写目标。
- WinUI 3：现代但运行时和打包复杂度高，不适合作为本项目第一版重写路线。

可保留为后续备选：

- WebView2 仅用于必要网页预览，不作为主 UI。
- Windows App SDK 仅在明确接受额外运行时和安装复杂度后评估。

### 4.2 构建

- 构建系统：CMake。
- 编译器：MSVC v143 或更新。
- 标准：C++20。
- 依赖管理：vcpkg manifest 或源码 vendor 二选一；第一阶段优先源码 vendor，减少外部构建变量。
- Release CRT：优先 `/MT` 静态链接，降低运行时依赖。
- Debug：保留 `/MDd` 或常规调试配置，便于工具链调试。

建议编译参数：

```text
/std:c++20
/O2
/W4
/permissive-
/utf-8
/EHsc
/MP
```

建议质量门禁：

- Debug 开启 `/RTC1`。
- 定期运行 `/analyze`。
- 可选引入 clang-tidy，但不作为第一阶段阻断门禁。
- 对 C 层库启用边界测试和内存泄漏检查。

### 4.3 第三方库原则

优先级：

1. Windows 原生 API。
2. 单文件或少文件、可静态链接、许可证清晰的 C/C++ 库。
3. 已长期稳定、可被测试覆盖的小型库。

建议候选：

| 能力 | 候选 |
| --- | --- |
| JSON | yyjson、jsmn、nlohmann/json 三选一；若强调体积优先 yyjson |
| ZIP | miniz |
| SQLite | sqlite3 amalgamation，仅在结构化存储阶段启用 |
| Hash/HMAC/AES/PBKDF2 | Windows CNG/BCrypt 优先 |
| HTTP | WinHTTP |
| TOML | 优先写受控 patcher；确需完整解析再评估 toml++ |
| Excel xlsx | 自写最小 OpenXML 或 miniz + XML writer |
| PDF 渲染 | 优先系统能力；复杂需求另行评估 |

## 5. 目录规划

建议新建独立原生工程目录，避免污染现有 WPF 项目：

```text
src/MyTools.Native/
├─ CMakeLists.txt
├─ app/
│  ├─ main.cpp
│  ├─ app_context.*
│  ├─ single_instance.*
│  ├─ crash_handler.*
│  └─ tray_host.*
├─ ui/
│  ├─ window.*
│  ├─ renderer_d2d.*
│  ├─ layout.*
│  ├─ controls/
│  ├─ theme/
│  └─ dpi.*
├─ modules/
│  ├─ home/
│  ├─ codex_profiles/
│  ├─ codex_relay/
│  ├─ frp_tunnel/
│  ├─ screenshot/
│  ├─ multimedia/
│  ├─ schedule/
│  ├─ wechat_tools/
│  └─ system_tools/
├─ services/
│  ├─ config_store.*
│  ├─ secret_store_dpapi.*
│  ├─ logger.*
│  ├─ process_runner.*
│  ├─ file_system.*
│  ├─ task_runner.*
│  ├─ hotkey_service.*
│  └─ network_service.*
├─ ccore/
│  ├─ include/
│  ├─ fs_scan/
│  ├─ json/
│  ├─ zip/
│  ├─ crypto/
│  └─ tests/
├─ resources/
│  ├─ app.rc
│  ├─ icons/
│  └─ images/
├─ installer/
└─ tests/
```

现有 `src/MyTools/` 在迁移完成前作为参考实现保留。原生版本不直接依赖 WPF 项目产物。

## 6. 架构分层

### 6.1 app 层

职责：

- 程序入口。
- 单实例锁。
- 主窗口创建。
- 托盘生命周期。
- 全局异常与崩溃记录。
- 应用路径、版本、数据目录管理。
- 启动期性能埋点。

约束：

- 启动期同步工作必须控制在极小范围内。
- 不扫描磁盘，不枚举注册表，不启动外部进程，不读大型配置。
- 主窗口首帧之后再延迟加载模块数据。

### 6.2 ui 层

职责：

- Win32 消息循环。
- Direct2D 绘制。
- DirectWrite 字体与文本布局。
- DPI 缩放。
- 控件基础库。
- 主题、颜色、间距、图标。

第一阶段需要实现的基础控件：

- Button。
- IconButton。
- TextBox。
- PasswordBox。
- ComboBox。
- CheckBox。
- Toggle。
- Slider。
- Tab。
- ListView。
- VirtualListView。
- Table/Grid。
- Dialog。
- Toast/Status message。

UI 规则：

- 支持 Per-Monitor V2 DPI。
- 所有尺寸使用 dp 逻辑单位，通过 DPI 转物理像素。
- 大列表必须虚拟化。
- 模块切换不销毁长期状态，除非该模块明确可重建。
- 文本裁剪、换行和 tooltip 必须统一处理。

### 6.3 services 层

职责：

- 配置读写。
- DPAPI 加解密。
- 日志脱敏。
- 文件扫描。
- 进程启动/停止。
- 网络探测。
- 热键。
- 任务队列。

服务层不能持有 UI 控件指针，只能通过事件、回调或消息队列向 UI 报告状态。

### 6.4 modules 层

每个模块必须包含：

- module controller：模块入口与生命周期。
- module state：当前状态。
- module view：界面渲染与输入处理。
- module service：业务服务封装。
- tests：可复用的非 UI 测试。

模块之间不能直接互相调用 UI；共享能力必须下沉到 services。

### 6.5 ccore 层

C 层只做底层能力，不做业务编排。

建议 C ABI 规则：

```c
typedef struct mt_result {
    int code;
    const char* message;
} mt_result;

typedef struct mt_buffer {
    unsigned char* data;
    size_t size;
} mt_buffer;

void mt_buffer_free(mt_buffer* buffer);
```

规则：

- C 层不抛异常。
- C 层返回错误码和消息。
- 跨边界内存必须提供释放函数。
- 所有字符串边界明确 UTF-8 或 UTF-16。
- Windows 文件路径在服务层统一转 UTF-16。

## 7. 数据与配置

### 7.1 配置格式

建议每个模块一个配置文件：

```text
MyTools.settings.json
Codex/profiles.json
Codex/active.json
Frp/frp-settings.json
Schedules/{YYYY-MM}/{version}.json
Screenshot/screenshot-settings.json
```

所有配置必须包含：

```json
{
  "schema_version": 1
}
```

迁移策略：

- 原生版本首次启动时只读取必要全局配置。
- 模块首次进入时再迁移该模块配置。
- 迁移前写 `.bak` 备份。
- 迁移失败不破坏旧数据，提示用户查看日志。

### 7.2 密钥与敏感信息

必须继续使用当前用户范围 DPAPI。

要求：

- Native 使用 `CryptProtectData` / `CryptUnprotectData`。
- 兼容 .NET `ProtectedData.Protect` 生成的数据。
- 日志严禁写 Token、密码、完整连接串、完整 auth.json。
- 内存中的明文密钥使用后尽快清零。
- `.codexbox` 使用 PBKDF2 + AES-256 + HMAC 的现有格式，Native 使用 Windows CNG 实现。

### 7.3 日志

日志文件建议：

```text
MyTools.log
MyTools.startup.log
MyTools.crash.log
```

规则：

- 单条日志不超过 1 KB。
- 敏感字段统一脱敏。
- 文件滚动，避免无限增长。
- 启动失败必须记录到 startup log。
- 崩溃时可选生成 minidump，但默认不包含敏感内存块。

## 8. UI 设计系统规划

### 8.1 主界面结构

建议继续使用桌面工具风格：

```text
标题栏
菜单栏
模块内容区
底部状态栏
托盘入口
```

不要做营销式首页。启动后直接进入可用工具界面。

### 8.2 视觉方向

- 安静、清晰、工具感强。
- 避免大面积单一色相。
- 操作按钮优先使用图标 + tooltip。
- 卡片半径不超过 8px。
- 不使用装饰性渐变球、纯视觉大背景。
- 信息密度高但留足扫描空间。

### 8.3 自绘控件优先级

第一批：

- 菜单栏。
- 按钮。
- 输入框。
- 下拉框。
- 列表。
- 虚拟列表。
- 对话框。
- 状态条。

第二批：

- 表格。
- 多选网格。
- 图片预览。
- 进度条。
- 音量/进度滑块。
- 日期/月历控件。

第三批：

- 排班复杂表格。
- 截图选区蒙层。
- 多媒体沉浸预览。
- 图表/统计视图。

## 9. 模块迁移顺序

### 阶段 0：盘点与基线

目标：确认当前要迁移的功能清单、体积、启动时间和配置文件。

交付物：

- 原生重写功能清单。
- 旧版当前体积与启动基线。
- 保留/删除功能表。
- Native 工程初始 CMake。

验收：

- 不改现有 WPF 行为。
- `scripts/codex-eval.ps1 -Quick` 通过。

### 阶段 1：Native Shell

目标：做出可启动的原生空壳。

范围：

- WinMain。
- 单实例。
- 主窗口。
- 菜单栏。
- 状态栏。
- 托盘。
- DPI。
- 日志。
- 崩溃记录。
- 配置目录。
- DPAPI smoke test。

验收：

- 冷启动首帧有可测量日志。
- 主窗口可显示/隐藏到托盘。
- 退出释放资源。
- Release 产物体积建立基线。

### 阶段 2：基础服务

目标：为后续模块提供稳定底座。

范围：

- ConfigStore。
- SecretStore。
- TaskRunner。
- ProcessRunner。
- FileScanner。
- NetworkService。
- HotkeyService。
- Zip/JSON/Crypto C 层封装。

验收：

- 单元测试覆盖 DPAPI、JSON、原子写入、进程启动停止。
- 所有磁盘扫描支持取消。
- UI 线程不阻塞。

### 阶段 3：Codex 档案管理

目标：优先迁移最适合验证架构的本地文件管理模块。

范围：

- 档案库读取。
- 档案列表。
- 切换档案。
- 测试中转。
- 复制、改名、备注、差异。
- config/auth 导出。
- `.codexbox` 导入导出。
- 本地中转启动/停止。

不迁回：

- 轮换池。
- 立即轮换。
- Ollama 模型导入。

验收：

- 可读取旧 DPAPI 档案。
- 可切换到 `~/.codex`。
- 失败时保留备份。
- 日志不泄露 token。

### 阶段 4：FRP 隧道穿透

目标：迁移进程控制和端口规则能力。

范围：

- 服务器配置。
- Token DPAPI 保存。
- 端口预设。
- 规则增删。
- frpc 配置生成。
- frpc 启动/停止。
- 输出日志截取。

验收：

- Token 不明文落盘。
- 远程端口重复校验。
- 停止流程先温和关闭，超时再 kill。
- 不在启动期访问网络。

### 阶段 5：截图、录屏、录音

目标：迁移高频工具模块。

范围：

- 全屏截图。
- 区域截图。
- 窗口截图。
- 截图编辑器第一版。
- 全局热键。
- 区域录屏。
- 系统声音录制。
- 外置 FFmpeg 查找与转码。

建议技术：

- 截图：Win32/GDI 或 DXGI Desktop Duplication。
- 录屏：DXGI Desktop Duplication 优先；Windows Graphics Capture 作为后续增强。
- 音频：WASAPI loopback。
- 转码：外置 `ffmpeg.exe`。

验收：

- 热键启动后可用。
- FFmpeg 缺失时提示，不阻塞截图。
- 剪贴板图片兼容微信、Office、画图。
- 录音无声检测保留。

### 阶段 6：系统工具

目标：迁移 Windows API 密集模块。

范围：

- 当前网络。
- 启动项管理。
- 程序卸载。
- 垃圾清理。
- 系统设置。
- 系统备份。
- 系统信息摘要。

约束：

- 不恢复硬件实时传感器。
- WMI 与注册表枚举必须按需触发。
- 删除/清理必须保留安全边界和可审计日志。

验收：

- 清理候选默认保守。
- junction/symlink 跳过。
- 提权操作参数可审计。
- 不在启动期扫描系统。

### 阶段 7：多媒体

目标：迁移文件浏览、预览和转换。

范围：

- 文件列表。
- 图片预览。
- 音视频播放。
- PDF/图片转换。
- 文本/Markdown/Office 文件信息预览。
- 视频缩略图，依赖外置 FFmpeg。

验收：

- 大目录列表虚拟化。
- 缩略图后台生成。
- FFmpeg 缺失时使用占位图。
- 不依赖 Office COM。

### 阶段 8：排班管理

目标：迁移交互最复杂的数据模块。

范围：

- 月度版本。
- 动态表格。
- 多选填充。
- 自动排休。
- 冲突检测。
- Excel 导入导出。

建议：

- 先迁移模型和算法测试。
- 再迁移表格控件。
- 最后迁移复杂交互。

验收：

- 与旧版核心排班样例输出一致。
- 大人员表格不卡顿。
- Excel 导入导出无需 Office。

### 阶段 9：微信工具

目标：迁移微信清理、备份、恢复。

范围：

- 账号定位。
- 候选扫描。
- 清理。
- 备份。
- 恢复。
- 最近备份列表。

验收：

- 不误判 QQ/Tencent 数字目录为微信账号。
- 删除前可预览候选。
- 备份/恢复支持取消和日志。
- 长扫描不阻塞 UI。

### 阶段 10：安装器与切换

目标：原生版本具备独立分发能力。

范围：

- Native 安装器。
- 卸载器。
- 旧版配置迁移。
- 开始菜单/桌面快捷方式。
- 安装后首启检查。

验收：

- 安装包 payload 可验证。
- 卸载保留用户数据选项。
- 旧版与 Native 版本不会互相破坏配置。
- 最终 installer pipeline 纳入自动验证。

## 10. 验证体系

建议新增 Native 专用验证入口：

```text
scripts/native-eval.ps1
```

建议参数：

```text
-Quick
-Build
-Unit
-Installer
-Smoke
```

第一阶段门禁：

- CMake configure。
- CMake build Release。
- 单元测试。
- `git diff --check`。
- 产物大小统计。
- 启动日志存在且首帧耗时可读。

后续阶段增加：

- 配置迁移测试。
- DPAPI 兼容测试。
- installer payload 检查。
- 外置 FFmpeg 检查。
- 大目录扫描取消测试。
- 排班样例导入导出测试。

## 11. 性能基线

每个阶段都要记录：

- Release exe 大小。
- 安装包大小。
- 冷启动首帧耗时。
- 首屏稳定后私有工作集。
- 模块首次进入耗时。
- 大列表渲染耗时。

建议目标：

| 指标 | 阶段 1 目标 | 完整版目标 |
| --- | ---: | ---: |
| 主 exe 体积 | 小于 3 MB | 小于 15 MB |
| 安装包体积 | 小于 5 MB | 小于 25 MB，不含外置 FFmpeg |
| 冷启动首帧 | 小于 300 ms | 小于 500 ms |
| 首屏私有工作集 | 小于 40 MB | 小于 120 MB |
| 模块切换响应 | 小于 100 ms | 小于 200 ms |

这些目标是初始建议。正式实现时以真实功能复杂度和测试结果修正。

## 12. 安全要求

- 所有敏感配置必须 DPAPI 加密。
- 日志脱敏必须集中实现，禁止模块自行拼完整敏感文本。
- 外部进程启动参数必须避免泄露 token。
- 导入包必须校验格式、版本、HMAC 或签名。
- 文件清理必须限制在明确候选路径内。
- 提权操作必须最小化，并保留用户确认。
- 网络请求必须由用户动作触发，启动期不自动探测外部服务。

## 13. 兼容与迁移

### 13.1 旧数据兼容

必须兼容：

- `MyTools.settings.json`。
- Codex `profiles.json` 和 `active.json`。
- `.codexbox` 加密包。
- FRP 配置。
- 排班 JSON。
- 截图设置。
- 微信备份记录。

不需要兼容：

- SQL 历史配置。
- 已移除轮换设置的 UI 行为。
- Ollama 自动导入生成记录。
- 硬件传感器缓存。

### 13.2 迁移策略

- 原生程序首次启动不主动全量迁移。
- 用户进入模块时按需迁移该模块数据。
- 迁移前写备份。
- 迁移后写 `schema_version`。
- 迁移失败保留旧数据并给出可读错误。

## 14. 风险与对策

| 风险 | 表现 | 对策 |
| --- | --- | --- |
| 自绘 UI 工期膨胀 | 控件细节占用大量时间 | 第一阶段只做基础控件，不追求完整复刻 |
| 排班表复杂度高 | 表格、选择、导入导出难迁 | 排班放到后期，先迁算法测试 |
| 配置迁移破坏旧数据 | 用户原数据丢失 | 所有迁移先备份，失败不写回 |
| DPAPI 兼容问题 | Native 读不了旧密文 | 阶段 2 独立做 DPAPI 兼容测试 |
| FFmpeg 外置体验下降 | 转码功能失败 | 统一查找路径和提示，功能降级不崩溃 |
| Win10 build 差异 | 捕获/暗色 API 不可用 | 建立最低 build，运行时动态检测 |
| 安装包回退 | 新旧 exe 混用 | installer payload 和版本写入门禁 |

## 15. 第一阶段详细任务

第一阶段只做 Native Shell，不迁业务模块。

任务列表：

1. 新建 `src/MyTools.Native/`。
2. 新建 CMake 工程。
3. 添加 `app.rc` 和图标资源。
4. 实现 `WinMain`。
5. 实现单实例 Mutex。
6. 实现主窗口创建与消息循环。
7. 实现 Direct2D renderer 初始化。
8. 实现 DirectWrite 文本绘制。
9. 实现 Per-Monitor V2 DPI 感知。
10. 实现菜单栏、内容占位区和状态栏。
11. 实现托盘图标、显示窗口、退出程序。
12. 实现 `Logger`。
13. 实现 `AppContext`。
14. 实现启动日志。
15. 实现崩溃日志。
16. 实现 `SecretStore` DPAPI smoke test。
17. 新建 `scripts/native-eval.ps1`。
18. 记录启动基线。

第一阶段不做：

- 不迁移 Codex。
- 不迁移 FRP。
- 不迁移截图。
- 不写安装器。
- 不删除 WPF 代码。

第一阶段验收标准：

- `scripts/native-eval.ps1 -Quick` 通过。
- Release 产物可双击启动。
- 首屏有菜单栏、内容区、状态栏。
- 托盘隐藏/恢复/退出可用。
- 关闭窗口行为符合设计。
- 日志文件在程序同级目录生成；如后续改为专用本地数据目录，必须先在项目规则和迁移方案中明确。
- DPAPI smoke test 通过。
- 无额外 DLL 散落，除非规划明确说明。

## 16. 需要用户确认的决策

正式实施前建议确认：

1. Windows 10 最低版本是否定为 22H2。
2. 是否只发布 x64。
3. 是否接受第一版 UI 不完全复刻 WPF 视觉，只保证专业、清晰、可用。
4. 是否允许新建 `src/MyTools.Native/` 与现有 WPF 程序并行一段时间。
5. 是否继续保留当前 WPF 安装器，直到 Native 安装器完成。
6. 是否把完整迁移目标拆成多个可安装预览版。
7. 是否将“SQL 导出、Codex 轮换、Ollama 导入、硬件传感器”永久列入 Native 不迁移清单。

## 17. 完成定义

Native 重写完成不是“代码能编译”，而是满足以下条件：

- 所有保留模块均已迁移或明确取消。
- 旧数据迁移路径可验证。
- 安装器可独立安装、卸载、升级。
- 启动、体积和内存基线优于 WPF 版本，或有明确不可避免原因。
- 所有敏感数据继续 DPAPI 加密。
- 日志脱敏通过检查。
- 外置 FFmpeg 缺失时降级可控。
- 自动验证覆盖构建、核心配置、安装器 payload、关键模块 smoke test。
- 活文档同步到新的 Native 架构。

## 18. 建议下一步

下一步不要直接迁业务模块，先执行阶段 0 与阶段 1：

1. 用当前 WPF 版本生成“保留功能清单”和“Native 不迁移清单”。
2. 建立 Native CMake 空壳。
3. 跑通 Win32 + Direct2D + DirectWrite 主窗口。
4. 建立日志、DPAPI、托盘、DPI 和启动基线。
5. 再决定 Codex 档案管理是否作为第一个业务模块迁移。
