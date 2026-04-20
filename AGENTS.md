# MyTools - Windows 个人实用工具集

本项目是一个基于 .NET Framework 4.8 的原生 Windows 桌面工具，旨在提供高效、美观、且在 Win7/10/11 环境下均可直接运行的个人常用功能（如系统增强、轻量数据处理等）。

## 一、项目原则
1. **高兼容性**：必须支持 Windows 7 SP1 以上所有系统，严禁引入 .NET Core 或高版本运行时依赖。
2. **极简分发**：最终产物必须是单一的 `.exe` 文件（利用 Costura.Fody 打包），严禁产生零散 DLL。
3. **视觉 Premium**：界面必须现代、流畅，符合 Material Design 审美。
4. **本地优先**：数据存储在本地 SQLite，日志记录在同级目录，不依赖云端。

## 二、核心技术栈
- **框架**: .NET Framework 4.8 / WPF (MVVM 模式)
- **UI 组件**: 
  - **MaterialDesignInXaml**: 全局风格。
  - **MahApps.Metro**: 窗体容器。
  - **FluentWPF**: 动态毛玻璃效果。
- **数据层**: **SQLite** + **Dapper** (极简 ORM)。
- **工具**: **Newtonsoft.Json** (序列化), **Serilog** (异步日志)。

## 三、开发规范
### 3.1 UI/UX 开发
- 必须支持 **Per-Monitor V2 高 DPI** 自适应。
- 使用 `MetroWindow` 作为主窗体，并开启 `MaterialDesign` 主题集成。
- 所有颜色、字体大小必须定义在 `ResourceDictionary` 中，方便全局调整。

### 3.2 编码规范
- **异步原则**: 所有磁盘 IO、网络请求必须使用 `async/await`，禁止卡顿 UI。
- **异常处理**: 全局拦截未处理异常，并弹窗提示用户，同时记录到 Serilog。
- **依赖管理**: 尽量减少 NuGet 包依赖，优先选择轻量级、无二次依赖的库。

## 四、工作流程
1. **需求定义**: 每次新增功能先在 `docs/功能说明.md` 中简述逻辑。
2. **高效开发**: 优先编写 ViewModel 和 Model，最后打磨 UI。
3. **分发打包**: 编译 Release 版本时，确保 Costura.Fody 正确合并了所有资源。
4. **日志维护**: 每次更新后，在 `docs/开发记录.txt` 中简要说明修改点。
