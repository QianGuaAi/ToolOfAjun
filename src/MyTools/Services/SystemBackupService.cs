using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    /// <summary>
    /// 程序数据 / 依赖文件的导出与导入。
    /// 备份结构（一份独立文件夹）：
    ///   <root>/
    ///     manifest.json
    ///     App/                     程序目录下的用户数据 + NativeBinaries（含 ffmpeg）
    ///     LocalAppData/MyTools/    %LOCALAPPDATA%\MyTools（Schedules、holidays.json 等）
    ///     Codex/                   %USERPROFILE%\.codex\{config.toml,auth.json}
    /// </summary>
    public static class SystemBackupService
    {
        private const string ManifestFileName = "manifest.json";
        private const string BackupFolderPrefix = "MyToolsBackup_";

        // 程序目录下需要打包的相对路径（文件 / 文件夹）
        private static readonly string[] AppDirRelativeTargets = new[]
        {
            "MyTools.settings.json",
            "MyTools.sqlhistory.json",
            "Configs",          // WireGuard 配置文件夹
            "NativeBinaries",   // 含 ffmpeg.exe / 锁定脚本等
        };

        public static string AppDir => AppDomain.CurrentDomain.BaseDirectory;
        public static string LocalAppDataDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyTools");
        public static string CodexDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex");

        // ====================================================================
        // Export
        // ====================================================================

        public class ExportResult
        {
            public string BackupRoot { get; set; }
            public int FilesCopied { get; set; }
            public long TotalBytes { get; set; }
            public List<string> Skipped { get; set; } = new List<string>();
        }

        /// <summary>导出全部数据到 <paramref name="targetParentFolder"/> 下一个时间戳子文件夹。</summary>
        public static async Task<ExportResult> ExportAsync(string targetParentFolder)
        {
            if (string.IsNullOrWhiteSpace(targetParentFolder))
                throw new ArgumentException("目标文件夹路径不能为空。", nameof(targetParentFolder));

            Directory.CreateDirectory(targetParentFolder);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var backupRoot = Path.Combine(targetParentFolder, BackupFolderPrefix + stamp);
            Directory.CreateDirectory(backupRoot);

            var result = new ExportResult { BackupRoot = backupRoot };

            await Task.Run(() =>
            {
                // 1) App dir 子集
                var appOut = Path.Combine(backupRoot, "App");
                foreach (var rel in AppDirRelativeTargets)
                {
                    var src = Path.Combine(AppDir, rel);
                    var dst = Path.Combine(appOut, rel);
                    if (File.Exists(src))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(dst));
                        File.Copy(src, dst, true);
                        result.FilesCopied++;
                        result.TotalBytes += new FileInfo(dst).Length;
                    }
                    else if (Directory.Exists(src))
                    {
                        var (n, b) = CopyDirectory(src, dst);
                        result.FilesCopied += n;
                        result.TotalBytes += b;
                    }
                    else
                    {
                        result.Skipped.Add(rel);
                    }
                }

                // 2) LocalAppData\MyTools
                if (Directory.Exists(LocalAppDataDir))
                {
                    var dst = Path.Combine(backupRoot, "LocalAppData", "MyTools");
                    var (n, b) = CopyDirectory(LocalAppDataDir, dst);
                    result.FilesCopied += n;
                    result.TotalBytes += b;
                }
                else
                {
                    result.Skipped.Add("%LOCALAPPDATA%\\MyTools");
                }

                // 3) %USERPROFILE%\.codex\{config.toml,auth.json}
                var codexOut = Path.Combine(backupRoot, "Codex");
                foreach (var name in new[] { "config.toml", "auth.json" })
                {
                    var src = Path.Combine(CodexDir, name);
                    if (File.Exists(src))
                    {
                        Directory.CreateDirectory(codexOut);
                        var dst = Path.Combine(codexOut, name);
                        File.Copy(src, dst, true);
                        result.FilesCopied++;
                        result.TotalBytes += new FileInfo(dst).Length;
                    }
                    else
                    {
                        result.Skipped.Add(".codex\\" + name);
                    }
                }

                // 4) Manifest
                var manifest = new BackupManifest
                {
                    CreatedAt = DateTime.Now,
                    MachineName = Environment.MachineName,
                    UserName = Environment.UserName,
                    AppDirSource = AppDir,
                    LocalAppDataSource = LocalAppDataDir,
                    CodexSource = CodexDir,
                    AppDirTargets = AppDirRelativeTargets.ToList(),
                    Note = "MyTools 系统设置导出文件。MyTools.settings.json 与 MyTools.sqlhistory.json " +
                           "采用 Windows DPAPI（当前用户）加密，仅在同一 Windows 账户上可解密。"
                };
                File.WriteAllText(Path.Combine(backupRoot, ManifestFileName),
                    JsonConvert.SerializeObject(manifest, Formatting.Indented),
                    System.Text.Encoding.UTF8);
                result.FilesCopied++;
            }).ConfigureAwait(false);

            AppLogService.Information(
                "System data exported. Root={Root}, Files={Files}, Bytes={Bytes}",
                result.BackupRoot, result.FilesCopied, result.TotalBytes);
            return result;
        }

        // ====================================================================
        // Import
        // ====================================================================

        public class ImportResult
        {
            public int FilesCopied { get; set; }
            public long TotalBytes { get; set; }
            public List<string> Sections { get; set; } = new List<string>();
            public List<string> Skipped { get; set; } = new List<string>();
            public bool HadDpapiData { get; set; }
            public string SourceUserName { get; set; }
        }

        /// <summary>
        /// 校验 <paramref name="sourceFolder"/> 是否是有效备份。
        /// 接受两种结构：直接是备份根（含 manifest.json），或上一级父文件夹（自动选最新的子备份）。
        /// </summary>
        public static string ResolveBackupRoot(string sourceFolder)
        {
            if (string.IsNullOrWhiteSpace(sourceFolder)) return null;
            if (File.Exists(Path.Combine(sourceFolder, ManifestFileName))) return sourceFolder;

            // 尝试在子目录里找最新的备份
            try
            {
                var dirs = Directory.GetDirectories(sourceFolder, BackupFolderPrefix + "*");
                if (dirs.Length == 0) return null;
                Array.Sort(dirs, StringComparer.OrdinalIgnoreCase);
                var latest = dirs[dirs.Length - 1];
                if (File.Exists(Path.Combine(latest, ManifestFileName))) return latest;
            }
            catch { }
            return null;
        }

        public static async Task<ImportResult> ImportAsync(string backupRoot)
        {
            if (string.IsNullOrWhiteSpace(backupRoot)) throw new ArgumentException(nameof(backupRoot));
            if (!Directory.Exists(backupRoot)) throw new DirectoryNotFoundException(backupRoot);

            var manifestPath = Path.Combine(backupRoot, ManifestFileName);
            if (!File.Exists(manifestPath))
                throw new InvalidOperationException("所选文件夹不是有效的备份目录（缺少 manifest.json）。");

            var manifest = JsonConvert.DeserializeObject<BackupManifest>(File.ReadAllText(manifestPath))
                           ?? new BackupManifest();
            var result = new ImportResult { SourceUserName = manifest.UserName };

            await Task.Run(() =>
            {
                // 1) App dir 子集
                var appIn = Path.Combine(backupRoot, "App");
                if (Directory.Exists(appIn))
                {
                    var (n, b) = CopyDirectory(appIn, AppDir, overwrite: true);
                    result.FilesCopied += n;
                    result.TotalBytes += b;
                    result.Sections.Add($"程序目录：{n} 个文件");

                    // 标记是否含 DPAPI 数据
                    if (File.Exists(Path.Combine(appIn, "MyTools.settings.json"))
                        || File.Exists(Path.Combine(appIn, "MyTools.sqlhistory.json")))
                    {
                        result.HadDpapiData = true;
                    }
                }
                else
                {
                    result.Skipped.Add("App 子目录");
                }

                // 2) LocalAppData
                var ladIn = Path.Combine(backupRoot, "LocalAppData", "MyTools");
                if (Directory.Exists(ladIn))
                {
                    Directory.CreateDirectory(LocalAppDataDir);
                    var (n, b) = CopyDirectory(ladIn, LocalAppDataDir, overwrite: true);
                    result.FilesCopied += n;
                    result.TotalBytes += b;
                    result.Sections.Add($"LocalAppData：{n} 个文件");
                }
                else
                {
                    result.Skipped.Add("LocalAppData\\MyTools");
                }

                // 3) Codex
                var codexIn = Path.Combine(backupRoot, "Codex");
                if (Directory.Exists(codexIn))
                {
                    Directory.CreateDirectory(CodexDir);
                    int copiedCodex = 0;
                    long bytesCodex = 0;
                    foreach (var f in Directory.GetFiles(codexIn))
                    {
                        var dst = Path.Combine(CodexDir, Path.GetFileName(f));
                        File.Copy(f, dst, true);
                        copiedCodex++;
                        bytesCodex += new FileInfo(dst).Length;
                    }
                    result.FilesCopied += copiedCodex;
                    result.TotalBytes += bytesCodex;
                    if (copiedCodex > 0)
                        result.Sections.Add($"Codex：{copiedCodex} 个文件");
                }
                else
                {
                    result.Skipped.Add(".codex");
                }
            }).ConfigureAwait(false);

            AppLogService.Information(
                "System data imported. Root={Root}, Files={Files}, Bytes={Bytes}, HadDpapi={Dpapi}, FromUser={User}",
                backupRoot, result.FilesCopied, result.TotalBytes, result.HadDpapiData, result.SourceUserName ?? "");
            return result;
        }

        // ====================================================================
        // Helpers
        // ====================================================================

        private static (int files, long bytes) CopyDirectory(string source, string target, bool overwrite = true)
        {
            int files = 0;
            long bytes = 0;
            if (!Directory.Exists(source)) return (0, 0);

            Directory.CreateDirectory(target);
            foreach (var dir in Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories))
            {
                var relative = GetRelativePath(source, dir);
                Directory.CreateDirectory(Path.Combine(target, relative));
            }
            foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
            {
                var relative = GetRelativePath(source, file);
                var dst = Path.Combine(target, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(dst));
                try
                {
                    File.Copy(file, dst, overwrite);
                    files++;
                    bytes += new FileInfo(dst).Length;
                }
                catch (Exception ex)
                {
                    AppLogService.Warning("Copy file failed: {Src} -> {Dst}: {Msg}", file, dst, ex.Message);
                }
            }
            return (files, bytes);
        }

        private static string GetRelativePath(string baseDir, string fullPath)
        {
            var b = Path.GetFullPath(baseDir).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            var f = Path.GetFullPath(fullPath);
            return f.StartsWith(b, StringComparison.OrdinalIgnoreCase)
                ? f.Substring(b.Length)
                : Path.GetFileName(fullPath);
        }

        public static string FormatBytes(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024L * 1024) return (bytes / 1024d).ToString("0.#") + " KB";
            if (bytes < 1024L * 1024 * 1024) return (bytes / 1024d / 1024d).ToString("0.##") + " MB";
            return (bytes / 1024d / 1024d / 1024d).ToString("0.##") + " GB";
        }
    }

    public class BackupManifest
    {
        public DateTime CreatedAt { get; set; }
        public string MachineName { get; set; }
        public string UserName { get; set; }
        public string AppDirSource { get; set; }
        public string LocalAppDataSource { get; set; }
        public string CodexSource { get; set; }
        public List<string> AppDirTargets { get; set; } = new List<string>();
        public string Note { get; set; }
    }
}
