# 启动基线 - Native Shell - 2026-06-30

## 范围

本记录对应 `src/MyTools.Native/` 阶段 1/2 Native Shell 与基础服务，并记录阶段 3 Codex Profiles 承载点起步。当前阶段包含 Win32 主窗口、Direct2D/DirectWrite 渲染、单实例、托盘、日志、崩溃记录、DPAPI smoke test、配置、二进制读取、二进制原子写入、原子文件写入、可取消扫描、任务队列、进程、网络探测、热键封装、模块导航、Codex Profiles 本地文件存在状态页、外层档案库元数据摘要读取、切换前备份 service、差异摘要 service、元数据编辑 service、当前目录导入 service、普通文件导出 service、`.codexbox` 加密包 service、最近备份恢复 service 和显式切换 service 骨架；这些 service 不在启动期或页面导航时自动执行，尚不包含任何完整业务模块迁移。

## 当前验证状态

- `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\native-eval.ps1 -Quick`：通过，包含阶段 2 基础服务静态守卫。
- `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\native-eval.ps1 -Build`：未通过，原因是当前机器未安装 `cmake`，也未找到 MSVC/Visual Studio Build Tools。
- `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\codex-eval.ps1 -Quick`：通过，包含阶段 1/2 `native shell scaffold` 静态守卫；WPF Release build 0 warning / 0 error。
- `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\native-eval.ps1 -Unit`：通过，当前仍为阶段 1/2/3 scaffold 源码守卫占位，尚未建立可执行 Native 单元测试二进制。

## 未测指标

以下指标暂未实测：

- Native Release exe 大小。
- 冷启动首帧耗时。
- 首屏稳定后私有工作集。
- 托盘隐藏/恢复交互耗时。
- Direct2D 首次渲染耗时。

## 未测原因

当前环境中未找到 CMake、MSVC `cl.exe`、`msbuild.exe` 或 Visual Studio `VsDevCmd.bat`，因此无法完成 Native Release 构建和真实启动测量。工具链补齐后，优先运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\native-eval.ps1 -Build
```

随后记录 `artifacts\native\build` 下 Release 产物大小，并对 `MyToolsNative.startup.log` 中的首窗口创建耗时做读回。

## 当前阶段目标

| 指标 | 目标 | 当前状态 |
| --- | ---: | --- |
| Native exe 大小 | 小于 3 MB | 未测 |
| 冷启动首帧 | 小于 300 ms | 未测 |
| 首屏私有工作集 | 小于 40 MB | 未测 |
| 静态源码守卫 | 通过 | 已通过 |
| WPF Quick 回归 | 通过 | 已通过 |
