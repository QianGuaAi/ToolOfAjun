# 启动基线说明 - Codex 导入 Ollama 模型

日期：2026-06-11

## 变更范围

- `MainViewModel` 构造函数新增 `ImportCodexOllamaProfilesCommand` 命令对象创建与公开属性绑定。
- `MainModulesView.xaml` 的 Codex 配置面板新增“导入 Ollama 模型”按钮。
- 新增 `CodexOllamaProfileService`，但只在用户点击“导入 Ollama 模型”后读取 `http://127.0.0.1:11434/api/tags`。
- `CodexRelayTestService` 与 `CodexLocalRelayService` 调整 loopback provider 无上游 key 处理，不改变启动期流程。

## 启动影响判断

- 构造期新增内容仅为一个 `AsyncRelayCommand` 对象和属性引用，属于启动期允许的命令对象创建。
- Ollama 模型枚举、HTTP 请求、档案加密保存均只在用户点击按钮后执行，不在 `App.OnStartup`、`MainWindow` 构造函数或 `MainViewModel` 构造函数中执行。
- Codex 面板仍由 `CodexProfilesDeferredTemplate` 按 `CurrentModule=CodexProfiles` 延迟挂载，不增加 Home 首屏模块渲染负担。

## 未实测原因

本次为窄范围 UI/命令与按需网络探测变更，未启动桌面程序采集首帧计时；以 Release 构建作为基础验证。若后续同时调整 `App.OnStartup`、`MainWindow` 构造函数或 `MainViewModel` 启动加载调度，应重新采集首帧、稳态私有工作集和单 exe 体积。
