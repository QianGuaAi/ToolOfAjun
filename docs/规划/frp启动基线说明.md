# frp 启动基线说明

## [2026-05-22] 新增 frp 隧道穿透模块

### 变更范围
- `MainViewModel` 构造函数仅新增 `ShowFrpCommand` 命令对象创建，属于启动期允许的轻量同步操作。
- `MainViewModel` 新增 `Frp` 懒加载属性，不在构造函数中创建 `FrpViewModel`。
- frp 配置读取通过 `ScheduleStartupBackgroundLoads()` 在 `DispatcherPriority.ApplicationIdle` 中调用 `SafeFireAndForget(Frp.LoadConfigAsync())`。
- `FrpViewModel` 构造函数只创建空集合、命令对象和默认草稿规则，不读取磁盘、不解压 `frpc.exe`、不启动进程。

### 基线说明
- 本次未引入启动期同步 WMI、注册表查询、大文件 IO 或进程启动。
- `NativeBinaries\frpc.exe` 作为嵌入资源会增加单 exe 体积约 12.9 MB，属于功能内置客户端的必要增长。
- 当前未进行可重复的首帧耗时和稳态私有工作集实测；原因是本次验证环境以 Release 构建产物检查为主，没有接入 GUI 首帧计时采样脚本。
- 未测指标在后续引入启动性能采样脚本后补充；如首帧、私有工作集或单 exe 体积除 frpc 嵌入外出现超过 5% 的额外回退，需要单独记录原因。
