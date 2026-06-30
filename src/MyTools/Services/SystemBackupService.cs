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
    ///   root/
    ///     manifest.json
    ///     App/                     程序目录下的用户数据 + NativeBinaries（外置依赖占位）
    ///     LocalAppData/MyTools/    %LOCALAPPDATA%\MyTools（Schedules、holidays.json 等）
    ///     Codex/                   %USERPROFILE%\.codex\{config.toml,auth.json}
    /// </summary>
    public static class SystemBackupService
    {
        private const string ManifestFileName = "manifest.json";
        private const string BackupFolderPrefix = "MyToolsBackup_";

        public const BackupSection DefaultSections =
            BackupSection.Settings
            | BackupSection.LocalAppData
            | BackupSection.Codex
            | BackupSection.NativeBinaries;

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
            public List<string> Sections { get; set; } = new List<string>();
            public List<string> Skipped { get; set; } = new List<string>();
        }

        public class BackupPlan
        {
            public BackupSection SelectedSections { get; set; }
            public int Files { get; set; }
            public long TotalBytes { get; set; }
            public List<BackupPlanItem> Items { get; set; } = new List<BackupPlanItem>();
            public List<string> Skipped { get; set; } = new List<string>();
        }

        public class BackupPlanItem
        {
            public BackupSection Section { get; set; }
            public string SectionName { get; set; }
            public string SourcePath { get; set; }
            public string BackupRelativePath { get; set; }
            public bool Exists { get; set; }
            public bool IsDirectory { get; set; }
            public int Files { get; set; }
            public long TotalBytes { get; set; }
        }

        public static Task<BackupPlan> BuildExportPlanAsync(BackupSection sections)
        {
            return Task.Run(() =>
            {
                var selectedSections = NormalizeSections(sections);
                var plan = new BackupPlan { SelectedSections = selectedSections };
                foreach (var target in BuildTargets(selectedSections, null, forImport: false))
                {
                    var item = BuildPlanItem(target);
                    plan.Items.Add(item);
                    if (item.Exists)
                    {
                        plan.Files += item.Files;
                        plan.TotalBytes += item.TotalBytes;
                    }
                    else
                    {
                        plan.Skipped.Add(target.SectionName + "：" + target.SourcePath);
                    }
                }

                return plan;
            });
        }

        public static Task<ExportResult> ExportAsync(string targetParentFolder)
        {
            return ExportAsync(targetParentFolder, DefaultSections);
        }

        /// <summary>导出选中数据到 <paramref name="targetParentFolder"/> 下一个时间戳子文件夹。</summary>
        public static async Task<ExportResult> ExportAsync(string targetParentFolder, BackupSection sections)
        {
            if (string.IsNullOrWhiteSpace(targetParentFolder))
            {
                throw new ArgumentException("目标文件夹路径不能为空。", nameof(targetParentFolder));
            }

            var selectedSections = NormalizeSections(sections);
            Directory.CreateDirectory(targetParentFolder);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var backupRoot = Path.Combine(targetParentFolder, BackupFolderPrefix + stamp);
            Directory.CreateDirectory(backupRoot);

            var result = new ExportResult { BackupRoot = backupRoot };

            await Task.Run(() =>
            {
                var sectionStats = new Dictionary<BackupSection, SectionCopyStats>();
                foreach (var target in BuildTargets(selectedSections, backupRoot, forImport: false))
                {
                    if (!File.Exists(target.SourcePath) && !Directory.Exists(target.SourcePath))
                    {
                        result.Skipped.Add(target.SectionName + "：" + target.SourcePath);
                        continue;
                    }

                    var copied = CopySourceToTarget(target.SourcePath, target.TargetPath);
                    AddSectionStats(sectionStats, target.Section, target.SectionName, copied.files, copied.bytes);
                    result.FilesCopied += copied.files;
                    result.TotalBytes += copied.bytes;
                }

                foreach (var stat in sectionStats.Values.OrderBy(item => item.SortOrder))
                {
                    result.Sections.Add($"{stat.SectionName}：{stat.Files} 个文件，{FormatBytes(stat.Bytes)}");
                }

                var manifest = new BackupManifest
                {
                    CreatedAt = DateTime.Now,
                    MachineName = Environment.MachineName,
                    UserName = Environment.UserName,
                    AppDirSource = AppDir,
                    LocalAppDataSource = LocalAppDataDir,
                    CodexSource = CodexDir,
                    Sections = GetSectionNames(selectedSections).ToList(),
                    Files = BuildManifestFileItems(backupRoot).ToList(),
                    Note = "MyTools 系统设置导出文件。MyTools.settings.json " +
                           "采用 Windows DPAPI（当前用户）加密，仅在同一 Windows 账户上可解密。"
                };
                File.WriteAllText(
                    Path.Combine(backupRoot, ManifestFileName),
                    JsonConvert.SerializeObject(manifest, Formatting.Indented),
                    System.Text.Encoding.UTF8);
                result.FilesCopied++;
            }).ConfigureAwait(false);

            AppLogService.Information(
                "System data exported. Root={Root}, Files={Files}, Bytes={Bytes}, Sections={Sections}",
                result.BackupRoot,
                result.FilesCopied,
                result.TotalBytes,
                string.Join(",", GetSectionNames(selectedSections)));
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

        public class ImportPreview
        {
            public string BackupRoot { get; set; }
            public DateTime CreatedAt { get; set; }
            public string MachineName { get; set; }
            public string SourceUserName { get; set; }
            public bool HadDpapiData { get; set; }
            public int IncomingFiles { get; set; }
            public long IncomingBytes { get; set; }
            public int ExistingTargetFiles { get; set; }
            public int NewTargetFiles { get; set; }
            public List<ImportPreviewItem> Items { get; set; } = new List<ImportPreviewItem>();
            public List<string> Skipped { get; set; } = new List<string>();
        }

        public class ImportPreviewItem
        {
            public BackupSection Section { get; set; }
            public string SectionName { get; set; }
            public string BackupPath { get; set; }
            public string TargetPath { get; set; }
            public bool Exists { get; set; }
            public bool IsDirectory { get; set; }
            public int Files { get; set; }
            public long TotalBytes { get; set; }
            public int ExistingTargetFiles { get; set; }
        }

        public class BackupVerifyResult
        {
            public string BackupRoot { get; set; }
            public int ExpectedFiles { get; set; }
            public int VerifiedFiles { get; set; }
            public long TotalBytes { get; set; }
            public List<string> Problems { get; set; } = new List<string>();
            public bool IsValid => Problems.Count == 0;
        }

        /// <summary>
        /// 校验 <paramref name="sourceFolder"/> 是否是有效备份。
        /// 接受两种结构：直接是备份根（含 manifest.json），或上一级父文件夹（自动选最新的子备份）。
        /// </summary>
        public static string ResolveBackupRoot(string sourceFolder)
        {
            if (string.IsNullOrWhiteSpace(sourceFolder))
            {
                return null;
            }

            if (File.Exists(Path.Combine(sourceFolder, ManifestFileName)))
            {
                return sourceFolder;
            }

            try
            {
                var dirs = Directory.GetDirectories(sourceFolder, BackupFolderPrefix + "*");
                if (dirs.Length == 0)
                {
                    return null;
                }

                Array.Sort(dirs, StringComparer.OrdinalIgnoreCase);
                var latest = dirs[dirs.Length - 1];
                return File.Exists(Path.Combine(latest, ManifestFileName)) ? latest : null;
            }
            catch
            {
                return null;
            }
        }

        public static Task<ImportPreview> BuildImportPreviewAsync(string backupRoot, BackupSection sections)
        {
            return Task.Run(() =>
            {
                var manifest = ReadManifest(backupRoot);
                var selectedSections = NormalizeSections(sections);
                var preview = new ImportPreview
                {
                    BackupRoot = backupRoot,
                    CreatedAt = manifest.CreatedAt,
                    MachineName = manifest.MachineName,
                    SourceUserName = manifest.UserName
                };

                foreach (var target in BuildTargets(selectedSections, backupRoot, forImport: true))
                {
                    var item = BuildImportPreviewItem(target);
                    preview.Items.Add(item);
                    if (!item.Exists)
                    {
                        preview.Skipped.Add(target.SectionName + "：" + target.SourcePath);
                        continue;
                    }

                    preview.IncomingFiles += item.Files;
                    preview.IncomingBytes += item.TotalBytes;
                    preview.ExistingTargetFiles += item.ExistingTargetFiles;
                }

                preview.NewTargetFiles = Math.Max(0, preview.IncomingFiles - preview.ExistingTargetFiles);
                preview.HadDpapiData = preview.Items.Any(item =>
                    item.Exists
                    && item.Section == BackupSection.Settings
                    && string.Equals(Path.GetFileName(item.BackupPath), "MyTools.settings.json", StringComparison.OrdinalIgnoreCase));
                return preview;
            });
        }

        public static Task<ImportResult> ImportAsync(string backupRoot)
        {
            return ImportAsync(backupRoot, DefaultSections);
        }

        public static Task<BackupVerifyResult> VerifyBackupAsync(string backupRoot)
        {
            return Task.Run(() =>
            {
                var manifest = ReadManifest(backupRoot);
                var result = new BackupVerifyResult { BackupRoot = backupRoot };
                var files = manifest.Files ?? new List<BackupManifestFileItem>();
                if (files.Count == 0)
                {
                    files = BuildManifestFileItems(backupRoot).ToList();
                }

                result.ExpectedFiles = files.Count;
                foreach (var item in files)
                {
                    var relativePath = item.RelativePath ?? string.Empty;
                    if (ContainsUnsafeRelativePath(relativePath))
                    {
                        result.Problems.Add("清单包含非法路径：" + relativePath);
                        continue;
                    }

                    var fullPath = Path.Combine(backupRoot, relativePath);
                    if (!File.Exists(fullPath))
                    {
                        result.Problems.Add("缺失文件：" + relativePath);
                        continue;
                    }

                    var length = new FileInfo(fullPath).Length;
                    if (item.Size >= 0 && length != item.Size)
                    {
                        result.Problems.Add($"大小不一致：{relativePath}，清单 {FormatBytes(item.Size)}，实际 {FormatBytes(length)}");
                        continue;
                    }

                    result.VerifiedFiles++;
                    result.TotalBytes += length;
                }

                return result;
            });
        }

        public static async Task<ImportResult> ImportAsync(string backupRoot, BackupSection sections)
        {
            var manifest = ReadManifest(backupRoot);
            var selectedSections = NormalizeSections(sections);
            var result = new ImportResult { SourceUserName = manifest.UserName };

            await Task.Run(() =>
            {
                var sectionStats = new Dictionary<BackupSection, SectionCopyStats>();
                foreach (var target in BuildTargets(selectedSections, backupRoot, forImport: true))
                {
                    if (!File.Exists(target.SourcePath) && !Directory.Exists(target.SourcePath))
                    {
                        result.Skipped.Add(target.SectionName + "：" + target.SourcePath);
                        continue;
                    }

                    if (target.Section == BackupSection.Settings
                        && string.Equals(Path.GetFileName(target.SourcePath), "MyTools.settings.json", StringComparison.OrdinalIgnoreCase))
                    {
                        result.HadDpapiData = true;
                    }

                    var copied = CopySourceToTarget(target.SourcePath, target.TargetPath);
                    AddSectionStats(sectionStats, target.Section, target.SectionName, copied.files, copied.bytes);
                    result.FilesCopied += copied.files;
                    result.TotalBytes += copied.bytes;
                }

                foreach (var stat in sectionStats.Values.OrderBy(item => item.SortOrder))
                {
                    result.Sections.Add($"{stat.SectionName}：{stat.Files} 个文件，{FormatBytes(stat.Bytes)}");
                }
            }).ConfigureAwait(false);

            AppLogService.Information(
                "System data imported. Root={Root}, Files={Files}, Bytes={Bytes}, HadDpapi={Dpapi}, FromUser={User}, Sections={Sections}",
                backupRoot,
                result.FilesCopied,
                result.TotalBytes,
                result.HadDpapiData,
                result.SourceUserName ?? string.Empty,
                string.Join(",", GetSectionNames(selectedSections)));
            return result;
        }

        // ====================================================================
        // Helpers
        // ====================================================================

        private static BackupSection NormalizeSections(BackupSection sections)
        {
            var normalized = sections & DefaultSections;
            if (normalized == 0)
            {
                throw new InvalidOperationException("请至少选择一个备份或恢复类别。");
            }

            return normalized;
        }

        private static IEnumerable<BackupTarget> BuildTargets(BackupSection sections, string backupRoot, bool forImport)
        {
            if ((sections & BackupSection.Settings) == BackupSection.Settings)
            {
                yield return CreateTarget(
                    BackupSection.Settings,
                    "设置与历史",
                    Path.Combine(AppDir, "MyTools.settings.json"),
                    Path.Combine("App", "MyTools.settings.json"),
                    backupRoot,
                    forImport);
            }

            if ((sections & BackupSection.NativeBinaries) == BackupSection.NativeBinaries)
            {
                yield return CreateTarget(
                    BackupSection.NativeBinaries,
                    "外置依赖",
                    Path.Combine(AppDir, "NativeBinaries"),
                    Path.Combine("App", "NativeBinaries"),
                    backupRoot,
                    forImport);
            }

            if ((sections & BackupSection.LocalAppData) == BackupSection.LocalAppData)
            {
                yield return CreateTarget(
                    BackupSection.LocalAppData,
                    "排班与本地数据",
                    LocalAppDataDir,
                    Path.Combine("LocalAppData", "MyTools"),
                    backupRoot,
                    forImport);
            }

            if ((sections & BackupSection.Codex) == BackupSection.Codex)
            {
                yield return CreateTarget(
                    BackupSection.Codex,
                    "Codex 当前配置",
                    Path.Combine(CodexDir, "config.toml"),
                    Path.Combine("Codex", "config.toml"),
                    backupRoot,
                    forImport);
                yield return CreateTarget(
                    BackupSection.Codex,
                    "Codex 当前配置",
                    Path.Combine(CodexDir, "auth.json"),
                    Path.Combine("Codex", "auth.json"),
                    backupRoot,
                    forImport);
            }
        }

        private static BackupTarget CreateTarget(
            BackupSection section,
            string sectionName,
            string livePath,
            string backupRelativePath,
            string backupRoot,
            bool forImport)
        {
            var backupPath = string.IsNullOrWhiteSpace(backupRoot)
                ? backupRelativePath
                : Path.Combine(backupRoot, backupRelativePath);
            return new BackupTarget
            {
                Section = section,
                SectionName = sectionName,
                SourcePath = forImport ? backupPath : livePath,
                TargetPath = forImport ? livePath : backupPath,
                BackupRelativePath = backupRelativePath
            };
        }

        private static BackupPlanItem BuildPlanItem(BackupTarget target)
        {
            var item = new BackupPlanItem
            {
                Section = target.Section,
                SectionName = target.SectionName,
                SourcePath = target.SourcePath,
                BackupRelativePath = target.BackupRelativePath
            };

            if (File.Exists(target.SourcePath))
            {
                var info = new FileInfo(target.SourcePath);
                item.Exists = true;
                item.Files = 1;
                item.TotalBytes = info.Length;
                return item;
            }

            if (Directory.Exists(target.SourcePath))
            {
                var counted = CountDirectory(target.SourcePath);
                item.Exists = true;
                item.IsDirectory = true;
                item.Files = counted.files;
                item.TotalBytes = counted.bytes;
            }

            return item;
        }

        private static ImportPreviewItem BuildImportPreviewItem(BackupTarget target)
        {
            var item = new ImportPreviewItem
            {
                Section = target.Section,
                SectionName = target.SectionName,
                BackupPath = target.SourcePath,
                TargetPath = target.TargetPath
            };

            if (File.Exists(target.SourcePath))
            {
                item.Exists = true;
                item.Files = 1;
                item.TotalBytes = new FileInfo(target.SourcePath).Length;
                item.ExistingTargetFiles = File.Exists(target.TargetPath) ? 1 : 0;
                return item;
            }

            if (Directory.Exists(target.SourcePath))
            {
                item.Exists = true;
                item.IsDirectory = true;
                var counted = CountDirectory(target.SourcePath);
                item.Files = counted.files;
                item.TotalBytes = counted.bytes;
                item.ExistingTargetFiles = CountExistingTargetFiles(target.SourcePath, target.TargetPath);
            }

            return item;
        }

        private static (int files, long bytes) CopySourceToTarget(string source, string target)
        {
            if (File.Exists(source))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                File.Copy(source, target, true);
                return (1, new FileInfo(target).Length);
            }

            return CopyDirectory(source, target);
        }

        private static (int files, long bytes) CopyDirectory(string source, string target)
        {
            var files = 0;
            long bytes = 0;
            if (!Directory.Exists(source))
            {
                return (0, 0);
            }

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
                    File.Copy(file, dst, true);
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

        private static (int files, long bytes) CountDirectory(string source)
        {
            var files = 0;
            long bytes = 0;
            if (!Directory.Exists(source))
            {
                return (0, 0);
            }

            foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
            {
                try
                {
                    files++;
                    bytes += new FileInfo(file).Length;
                }
                catch (Exception ex)
                {
                    AppLogService.Warning("Count file failed: {Path}: {Msg}", file, ex.Message);
                }
            }

            return (files, bytes);
        }

        private static int CountExistingTargetFiles(string sourceDirectory, string targetDirectory)
        {
            if (!Directory.Exists(sourceDirectory) || string.IsNullOrWhiteSpace(targetDirectory))
            {
                return 0;
            }

            var count = 0;
            foreach (var file in Directory.EnumerateFiles(sourceDirectory, "*", SearchOption.AllDirectories))
            {
                var relative = GetRelativePath(sourceDirectory, file);
                if (File.Exists(Path.Combine(targetDirectory, relative)))
                {
                    count++;
                }
            }

            return count;
        }

        private static BackupManifest ReadManifest(string backupRoot)
        {
            if (string.IsNullOrWhiteSpace(backupRoot))
            {
                throw new ArgumentException(nameof(backupRoot));
            }

            if (!Directory.Exists(backupRoot))
            {
                throw new DirectoryNotFoundException(backupRoot);
            }

            var manifestPath = Path.Combine(backupRoot, ManifestFileName);
            if (!File.Exists(manifestPath))
            {
                throw new InvalidOperationException("所选文件夹不是有效的备份目录（缺少 manifest.json）。");
            }

            return JsonConvert.DeserializeObject<BackupManifest>(File.ReadAllText(manifestPath))
                   ?? new BackupManifest();
        }

        private static IEnumerable<BackupManifestFileItem> BuildManifestFileItems(string backupRoot)
        {
            if (string.IsNullOrWhiteSpace(backupRoot) || !Directory.Exists(backupRoot))
            {
                yield break;
            }

            foreach (var file in Directory.EnumerateFiles(backupRoot, "*", SearchOption.AllDirectories))
            {
                var relative = GetRelativePath(backupRoot, file);
                if (string.Equals(relative, ManifestFileName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                FileInfo info;
                try
                {
                    info = new FileInfo(file);
                }
                catch
                {
                    continue;
                }

                yield return new BackupManifestFileItem
                {
                    RelativePath = relative,
                    Size = info.Length
                };
            }
        }

        private static bool ContainsUnsafeRelativePath(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath) || Path.IsPathRooted(relativePath))
            {
                return true;
            }

            var normalized = relativePath.Replace('/', '\\');
            return normalized.StartsWith("..\\", StringComparison.Ordinal)
                   || normalized.IndexOf("\\..\\", StringComparison.Ordinal) >= 0;
        }

        private static void AddSectionStats(
            IDictionary<BackupSection, SectionCopyStats> stats,
            BackupSection section,
            string sectionName,
            int files,
            long bytes)
        {
            SectionCopyStats stat;
            if (!stats.TryGetValue(section, out stat))
            {
                stat = new SectionCopyStats
                {
                    Section = section,
                    SectionName = sectionName,
                    SortOrder = GetSectionSortOrder(section)
                };
                stats[section] = stat;
            }

            stat.Files += files;
            stat.Bytes += bytes;
        }

        private static IEnumerable<string> GetSectionNames(BackupSection sections)
        {
            foreach (BackupSection section in new[]
            {
                BackupSection.Settings,
                BackupSection.LocalAppData,
                BackupSection.Codex,
                BackupSection.NativeBinaries
            })
            {
                if ((sections & section) == section)
                {
                    yield return GetSectionName(section);
                }
            }
        }

        private static string GetSectionName(BackupSection section)
        {
            switch (section)
            {
                case BackupSection.Settings:
                    return "设置与历史";
                case BackupSection.LocalAppData:
                    return "排班与本地数据";
                case BackupSection.Codex:
                    return "Codex 当前配置";
                case BackupSection.NativeBinaries:
                    return "外置依赖";
                default:
                    return section.ToString();
            }
        }

        private static int GetSectionSortOrder(BackupSection section)
        {
            switch (section)
            {
                case BackupSection.Settings:
                    return 0;
                case BackupSection.LocalAppData:
                    return 1;
                case BackupSection.Codex:
                    return 2;
                case BackupSection.NativeBinaries:
                    return 3;
                default:
                    return 99;
            }
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
            if (bytes < 1024)
            {
                return bytes + " B";
            }

            if (bytes < 1024L * 1024)
            {
                return (bytes / 1024d).ToString("0.#") + " KB";
            }

            if (bytes < 1024L * 1024 * 1024)
            {
                return (bytes / 1024d / 1024d).ToString("0.##") + " MB";
            }

            return (bytes / 1024d / 1024d / 1024d).ToString("0.##") + " GB";
        }

        private class BackupTarget
        {
            public BackupSection Section { get; set; }
            public string SectionName { get; set; }
            public string SourcePath { get; set; }
            public string TargetPath { get; set; }
            public string BackupRelativePath { get; set; }
        }

        private class SectionCopyStats
        {
            public BackupSection Section { get; set; }
            public string SectionName { get; set; }
            public int SortOrder { get; set; }
            public int Files { get; set; }
            public long Bytes { get; set; }
        }
    }

    [Flags]
    public enum BackupSection
    {
        None = 0,
        Settings = 1,
        LocalAppData = 2,
        Codex = 4,
        NativeBinaries = 8
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
        public List<string> Sections { get; set; } = new List<string>();
        public List<BackupManifestFileItem> Files { get; set; } = new List<BackupManifestFileItem>();
        public string Note { get; set; }
    }

    public class BackupManifestFileItem
    {
        public string RelativePath { get; set; }
        public long Size { get; set; }
    }
}
