请将 WireGuard 的二进制文件放入此目录，以实现绿色集成模式：

1. 从 WireGuard 官方提取或编译以下文件：
   - wireguard.exe
   - wg.exe
   - wintun.dll
   - wireguard.dll (针对 WireGuardNT)

2. 编译时，这些文件会自动复制到输出目录的 `NativeBinaries` 文件夹中。
3. 程序运行时会优先检测该目录下的二进制文件，从而不依赖系统安装。
