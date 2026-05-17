using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using VbFileIO = Microsoft.VisualBasic.FileIO;

namespace MyTools.Services
{
    public sealed class JunkCleanupService
    {
        private static readonly string WindowsRoot = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        private static readonly string LocalAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        private static readonly string[] ForbiddenRoots =
        {
            Path.Combine(WindowsRoot, "System32"),
            Path.Combine(WindowsRoot, "WinSxS"),
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
            Path.Combine(Environment.GetEnvironmentVariable("ProgramFiles") ?? @"C:\Program Files", string.Empty),
            Path.Combine(Environment.GetEnvironmentVariable("ProgramFiles(x86)") ?? @"C:\Program Files (x86)", string.Empty)
        };

        public async Task<List<JunkCandidate>> ScanAsync(IProgress<string> progress, CancellationToken ct)
        {
            var results = new List<JunkCandidate>();
            var step = 0;
            var now = DateTime.Now;
            var userTempThreshold = now.AddHours(-24);
            var systemTempThreshold = now.AddDays(-7);
            var prefetchThreshold = now.AddDays(-30);

            var userTempRoots = new[]
            {
                Path.GetTempPath(),
                Path.Combine(LocalAppData, "Temp")
            }
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(Directory.Exists)
            .ToList();

            progress?.Report($"[{++step}/9] 扫描用户临时目录...");
            foreach (var root in userTempRoots)
            {
                ScanAgedFiles(results, root, JunkCategory.UserTemp, userTempThreshold, "24 小时前临时文件", ct);
            }

            progress?.Report($"[{++step}/9] 扫描系统临时目录...");
            ScanAgedFiles(results, Path.Combine(WindowsRoot, "Temp"), JunkCategory.SystemTemp, systemTempThreshold, "7 天前系统临时文件", ct);

            progress?.Report($"[{++step}/9] 读取回收站容量...");
            var recycleInfo = NativeMethods.QueryRecycleBin();
            if (recycleInfo.bytes > 0)
            {
                results.Add(new JunkCandidate
                {
                    Path = "RecycleBin",
                    Category = JunkCategory.RecycleBin,
                    Bytes = recycleInfo.bytes,
                    Reason = "回收站可释放空间",
                    IsSelected = true
                });
            }

            progress?.Report($"[{++step}/9] 扫描浏览器缓存...");
            var browserCachePaths = new[]
            {
                Path.Combine(LocalAppData, "Microsoft", "Edge", "User Data", "Default", "Cache", "Cache_Data"),
                Path.Combine(LocalAppData, "Google", "Chrome", "User Data", "Default", "Cache", "Cache_Data")
            };

            var edgeRunning = Process.GetProcessesByName("msedge").Length > 0;
            var chromeRunning = Process.GetProcessesByName("chrome").Length > 0;
            var browserRunning = edgeRunning || chromeRunning;
            foreach (var cachePath in browserCachePaths.Where(Directory.Exists))
            {
                var bytes = SafePathHelper.CalculateDirectorySize(cachePath);
                if (bytes <= 0)
                {
                    continue;
                }

                results.Add(new JunkCandidate
                {
                    Path = cachePath,
                    Category = JunkCategory.BrowserCache,
                    Bytes = bytes,
                    Reason = browserRunning ? "浏览器正在运行，清理时将跳过" : "浏览器缓存",
                    IsSelected = false
                });
            }

            progress?.Report($"[{++step}/9] 扫描缩略图缓存...");
            var explorerCacheRoot = Path.Combine(LocalAppData, "Microsoft", "Windows", "Explorer");
            if (Directory.Exists(explorerCacheRoot))
            {
                foreach (var file in SafePathHelper.EnumerateFilesSafe(explorerCacheRoot, "thumbcache_*.db"))
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        var info = new FileInfo(file);
                        if (info.Length <= 0) continue;
                        results.Add(new JunkCandidate
                        {
                            Path = info.FullName,
                            Category = JunkCategory.ThumbnailCache,
                            Bytes = info.Length,
                            Reason = "缩略图缓存文件",
                            IsSelected = true
                        });
                    }
                    catch
                    {
                    }
                }
            }

            progress?.Report($"[{++step}/9] 扫描 Prefetch...");
            var prefetchPath = Path.Combine(WindowsRoot, "Prefetch");
            if (Directory.Exists(prefetchPath))
            {
                foreach (var file in SafePathHelper.EnumerateFilesSafe(prefetchPath, "*.pf"))
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        var info = new FileInfo(file);
                        if (info.LastWriteTime >= prefetchThreshold) continue;
                        results.Add(new JunkCandidate
                        {
                            Path = info.FullName,
                            Category = JunkCategory.PrefetchAged,
                            Bytes = info.Length,
                            Reason = "30 天前 Prefetch 缓存",
                            IsSelected = true
                        });
                    }
                    catch
                    {
                    }
                }
            }

            progress?.Report($"[{++step}/9] 扫描 Windows 更新缓存...");
            var wuCacheRoot = Path.Combine(WindowsRoot, "SoftwareDistribution", "Download");
            ScanAllFiles(results, wuCacheRoot, JunkCategory.WindowsUpdateCache, "Windows Update 缓存", true, ct);

            progress?.Report($"[{++step}/9] 扫描 Delivery Optimization 缓存...");
            var doCacheRoot = Path.Combine(WindowsRoot, "ServiceProfiles", "NetworkService", "AppData", "Local", "Microsoft", "Windows", "DeliveryOptimization", "Cache");
            ScanAllFiles(results, doCacheRoot, JunkCategory.DeliveryOptimization, "Delivery Optimization 缓存", true, ct);

            progress?.Report($"[{++step}/9] 扫描 Installer 空目录...");
            var packageCachePath = Path.Combine(LocalAppData, "Package Cache");
            if (Directory.Exists(packageCachePath))
            {
                foreach (var directory in SafePathHelper.EnumerateDirectoriesSafe(packageCachePath))
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        if (Directory.EnumerateFileSystemEntries(directory).Any())
                        {
                            continue;
                        }

                        results.Add(new JunkCandidate
                        {
                            Path = directory,
                            Category = JunkCategory.InstallerLeftover,
                            Bytes = 0,
                            Reason = "空目录（仅删除空目录）",
                            IsSelected = true
                        });
                    }
                    catch
                    {
                    }
                }
            }

            await Task.Yield();
            return results;
        }

        public async Task<JunkCleanupExecutionResult> CleanupAsync(
            IEnumerable<JunkCandidate> selectedCandidates,
            IProgress<string> progress,
            CancellationToken ct)
        {
            var candidates = (selectedCandidates ?? Enumerable.Empty<JunkCandidate>())
                .Where(x => x != null && x.IsSelected)
                .ToList();

            var result = new JunkCleanupExecutionResult();
            if (candidates.Count == 0)
            {
                return result;
            }

            var safeRoots = BuildSafeRoots();
            var adminGroups = candidates
                .Where(x => RequiresAdmin(x.Category))
                .GroupBy(x => x.Category)
                .ToDictionary(g => g.Key, g => g.ToList());

            var normalCandidates = candidates.Where(x => !RequiresAdmin(x.Category)).ToList();

            progress?.Report("正在清理普通缓存项...");
            foreach (var candidate in normalCandidates)
            {
                ct.ThrowIfCancellationRequested();
                var step = new OptimizationStep
                {
                    Name = $"清理 {candidate.Category}",
                    Status = "OK",
                    Detail = candidate.Path,
                    BytesFreed = 0
                };

                try
                {
                    if (candidate.Category == JunkCategory.RecycleBin)
                    {
                        AppLogService.Information("Junk cleanup plan: empty recycle bin");
                        var hr = NativeMethods.EmptyRecycleBin();
                        if (hr != 0 && hr != NativeMethods.S_FALSE)
                        {
                            step.Status = "Failed";
                            step.Detail = $"回收站清理失败（HRESULT=0x{hr:X8}）。";
                        }
                        else
                        {
                            step.BytesFreed = candidate.Bytes;
                        }
                    }
                    else if (candidate.Category == JunkCategory.InstallerLeftover)
                    {
                        if (!SafePathHelper.IsPathInsideAny(candidate.Path, safeRoots))
                        {
                            throw new InvalidOperationException("路径不在白名单内。");
                        }

                        if (Directory.Exists(candidate.Path) && !Directory.EnumerateFileSystemEntries(candidate.Path).Any())
                        {
                            AppLogService.Information("Junk cleanup plan: delete empty installer folder {Path}", candidate.Path);
                            Directory.Delete(candidate.Path, false);
                            step.BytesFreed = 0;
                        }
                        else
                        {
                            step.Status = "Skipped";
                            step.Detail = "目录非空，已跳过。";
                        }
                    }
                    else
                    {
                        if (!SafePathHelper.IsPathInsideAny(candidate.Path, safeRoots))
                        {
                            throw new InvalidOperationException("路径不在白名单内。");
                        }

                        if (candidate.Category == JunkCategory.BrowserCache
                            && (Process.GetProcessesByName("msedge").Length > 0 || Process.GetProcessesByName("chrome").Length > 0))
                        {
                            step.Status = "Skipped";
                            step.Detail = "浏览器正在运行，已跳过。";
                        }
                        else
                        {
                            step.BytesFreed = await DeleteCandidatePathAsync(candidate.Path).ConfigureAwait(false);
                        }
                    }
                }
                catch (IOException)
                {
                    step.Status = "Skipped";
                    step.Detail = "被进程占用";
                }
                catch (Exception ex)
                {
                    AppLogService.Error(ex, "Junk cleanup failed for {Path}", candidate.Path);
                    step.Status = "Failed";
                    step.Detail = ex.Message;
                }

                result.Steps.Add(step);
            }

            foreach (var pair in adminGroups)
            {
                ct.ThrowIfCancellationRequested();
                progress?.Report($"正在清理 {pair.Key}（管理员权限）...");
                var bytes = pair.Value.Sum(x => x.Bytes);
                var count = pair.Value.Count;
                var step = new OptimizationStep
                {
                    Name = $"清理 {pair.Key}",
                    Status = "OK",
                    Detail = $"计划删除 {count} 项。",
                    BytesFreed = 0
                };

                try
                {
                    foreach (var c in pair.Value)
                    {
                        if (!SafePathHelper.IsPathInsideAny(c.Path, safeRoots))
                        {
                            throw new InvalidOperationException("路径不在白名单内。");
                        }
                    }

                    AppLogService.Information("Junk cleanup plan: admin cleanup category {Category}, count {Count}, bytes {Bytes}",
                        pair.Key.ToString(), count, bytes);
                    await RunAdminCleanupScriptAsync(pair.Key, ct).ConfigureAwait(false);
                    AppLogService.Information("Junk cleanup done: admin cleanup category {Category}, deleted {Count} items, freed {Bytes}B",
                        pair.Key.ToString(), count, bytes);
                    step.BytesFreed = bytes;
                }
                catch (OperationCanceledException)
                {
                    step.Status = "Failed";
                    step.Detail = "用户取消了 UAC 授权";
                }
                catch (Exception ex)
                {
                    AppLogService.Error(ex, "Admin junk cleanup failed for {Category}", pair.Key.ToString());
                    step.Status = "Failed";
                    step.Detail = ex.Message;
                }

                result.Steps.Add(step);
            }

            result.DeletedCount = result.Steps.Count(x => string.Equals(x.Status, "OK", StringComparison.OrdinalIgnoreCase));
            result.FreedBytes = result.Steps.Sum(x => x.BytesFreed);
            return result;
        }

        private static bool RequiresAdmin(JunkCategory category)
        {
            return category == JunkCategory.SystemTemp
                || category == JunkCategory.WindowsUpdateCache
                || category == JunkCategory.DeliveryOptimization;
        }

        private static async Task RunAdminCleanupScriptAsync(JunkCategory category, CancellationToken ct)
        {
            string script;
            switch (category)
            {
                case JunkCategory.SystemTemp:
                    script = @"
$target = Join-Path $env:windir 'Temp'
$threshold = (Get-Date).AddDays(-7)
Get-ChildItem -LiteralPath $target -Force -Recurse -ErrorAction SilentlyContinue |
Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt $threshold } |
ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch { } }
Get-ChildItem -LiteralPath $target -Directory -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending |
ForEach-Object {
    try {
        if ((Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}";
                    break;
                case JunkCategory.WindowsUpdateCache:
                    script = @"
$target = Join-Path $env:windir 'SoftwareDistribution\Download'
Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
if (Test-Path $target) {
    Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue |
    ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { } }
}
Start-Service -Name wuauserv -ErrorAction SilentlyContinue";
                    break;
                case JunkCategory.DeliveryOptimization:
                    script = @"
$target = Join-Path $env:windir 'ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache'
Stop-Service -Name dosvc -Force -ErrorAction SilentlyContinue
if (Test-Path $target) {
    Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue |
    ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { } }
}
Start-Service -Name dosvc -ErrorAction SilentlyContinue";
                    break;
                default:
                    return;
            }

            await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
        }

        private static async Task<long> DeleteCandidatePathAsync(string fullPath)
        {
            if (File.Exists(fullPath))
            {
                var info = new FileInfo(fullPath);
                var bytes = info.Length;
                AppLogService.Information("Junk cleanup plan: recycle file {Path}", fullPath);
                info.Attributes = FileAttributes.Normal;
                // 发送到回收站而非永久删除，避免误删后无法恢复
                VbFileIO.FileSystem.DeleteFile(fullPath, VbFileIO.UIOption.OnlyErrorDialogs, VbFileIO.RecycleOption.SendToRecycleBin, VbFileIO.UICancelOption.ThrowException);
                await Task.Yield();
                AppLogService.Information("Junk cleanup done: recycled file {Path}, freed {Bytes}B", fullPath, bytes);
                return bytes;
            }

            if (Directory.Exists(fullPath))
            {
                if (Directory.EnumerateFileSystemEntries(fullPath).Any())
                {
                    return 0;
                }

                AppLogService.Information("Junk cleanup plan: delete empty directory {Path}", fullPath);
                Directory.Delete(fullPath, false);
                await Task.Yield();
                AppLogService.Information("Junk cleanup done: deleted empty directory {Path}", fullPath);
                return 0;
            }

            return 0;
        }

        private static List<string> BuildSafeRoots()
        {
            var roots = new List<string>
            {
                Path.GetTempPath(),
                Path.Combine(LocalAppData, "Temp"),
                Path.Combine(WindowsRoot, "Temp"),
                Path.Combine(WindowsRoot, "Prefetch"),
                Path.Combine(WindowsRoot, "SoftwareDistribution", "Download"),
                Path.Combine(WindowsRoot, "ServiceProfiles", "NetworkService", "AppData", "Local", "Microsoft", "Windows", "DeliveryOptimization", "Cache"),
                Path.Combine(LocalAppData, "Microsoft", "Windows", "Explorer"),
                Path.Combine(LocalAppData, "Microsoft", "Edge", "User Data", "Default", "Cache", "Cache_Data"),
                Path.Combine(LocalAppData, "Google", "Chrome", "User Data", "Default", "Cache", "Cache_Data"),
                Path.Combine(LocalAppData, "Package Cache")
            };

            return roots
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(path => !ForbiddenRoots.Any(forbidden => SafePathHelper.IsPathInside(path, forbidden)))
                .ToList();
        }

        private static void ScanAgedFiles(
            ICollection<JunkCandidate> target,
            string root,
            JunkCategory category,
            DateTime threshold,
            string reason,
            CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return;
            }

            foreach (var file in SafePathHelper.EnumerateFilesSafe(root))
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var info = new FileInfo(file);
                    if (info.LastWriteTime >= threshold)
                    {
                        continue;
                    }

                    target.Add(new JunkCandidate
                    {
                        Path = info.FullName,
                        Category = category,
                        Bytes = info.Length,
                        Reason = reason,
                        IsSelected = category != JunkCategory.BrowserCache
                    });
                }
                catch
                {
                }
            }
        }

        private static void ScanAllFiles(
            ICollection<JunkCandidate> target,
            string root,
            JunkCategory category,
            string reason,
            bool selected,
            CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return;
            }

            foreach (var file in SafePathHelper.EnumerateFilesSafe(root))
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var info = new FileInfo(file);
                    target.Add(new JunkCandidate
                    {
                        Path = info.FullName,
                        Category = category,
                        Bytes = info.Length,
                        Reason = reason,
                        IsSelected = selected
                    });
                }
                catch
                {
                }
            }
        }

        private static class NativeMethods
        {
            private const uint SHERB_NOCONFIRMATION = 0x00000001;
            private const uint SHERB_NOPROGRESSUI = 0x00000002;
            private const uint SHERB_NOSOUND = 0x00000004;
            public const int S_FALSE = 1;

            [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
            private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct SHQUERYRBINFO
            {
                public int cbSize;
                public long i64Size;
                public long i64NumItems;
            }

            [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
            private static extern int SHQueryRecycleBin(string pszRootPath, ref SHQUERYRBINFO pSHQueryRBInfo);

            public static (long bytes, long items) QueryRecycleBin()
            {
                var info = new SHQUERYRBINFO { cbSize = Marshal.SizeOf(typeof(SHQUERYRBINFO)) };
                var hr = SHQueryRecycleBin(null, ref info);
                if (hr != 0)
                {
                    return (0, 0);
                }

                return (info.i64Size, info.i64NumItems);
            }

            public static int EmptyRecycleBin()
            {
                return SHEmptyRecycleBin(IntPtr.Zero, null,
                    SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);
            }
        }
    }

    public sealed class JunkCleanupExecutionResult
    {
        public List<OptimizationStep> Steps { get; set; } = new List<OptimizationStep>();
        public int DeletedCount { get; set; }
        public long FreedBytes { get; set; }
    }

    public sealed class JunkCandidate : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string Path { get; set; }
        public JunkCategory Category { get; set; }
        public long Bytes { get; set; }
        public string Reason { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public string BytesDisplay => FileSizeFormatter.Format(Bytes);
        public string CategoryDisplay => GetCategoryDisplay(Category);
        public string RiskDisplay => GetRiskDisplay(Category);
        public string AdviceDisplay => GetAdviceDisplay(Category, Reason);

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        private static string GetCategoryDisplay(JunkCategory category)
        {
            switch (category)
            {
                case JunkCategory.UserTemp: return "用户临时";
                case JunkCategory.SystemTemp: return "系统临时";
                case JunkCategory.RecycleBin: return "回收站";
                case JunkCategory.BrowserCache: return "浏览器缓存";
                case JunkCategory.ThumbnailCache: return "缩略图";
                case JunkCategory.PrefetchAged: return "Prefetch";
                case JunkCategory.WindowsUpdateCache: return "更新缓存";
                case JunkCategory.DeliveryOptimization: return "传递优化";
                case JunkCategory.InstallerLeftover: return "安装残留";
                default: return category.ToString();
            }
        }

        private static string GetRiskDisplay(JunkCategory category)
        {
            switch (category)
            {
                case JunkCategory.RecycleBin:
                case JunkCategory.BrowserCache:
                case JunkCategory.SystemTemp:
                case JunkCategory.WindowsUpdateCache:
                case JunkCategory.DeliveryOptimization:
                    return "中";
                default:
                    return "低";
            }
        }

        private static string GetAdviceDisplay(JunkCategory category, string reason)
        {
            switch (category)
            {
                case JunkCategory.RecycleBin:
                    return "确认不再需要后清理";
                case JunkCategory.BrowserCache:
                    return "关闭浏览器后清理";
                case JunkCategory.SystemTemp:
                case JunkCategory.WindowsUpdateCache:
                case JunkCategory.DeliveryOptimization:
                    return "需管理员权限，按提示授权";
                case JunkCategory.ThumbnailCache:
                    return "会自动重建，可清理";
                case JunkCategory.PrefetchAged:
                    return "仅清理 30 天前缓存";
                case JunkCategory.InstallerLeftover:
                    return "仅删除空目录";
                default:
                    return string.IsNullOrWhiteSpace(reason) ? "可清理" : reason;
            }
        }
    }

    public enum JunkCategory
    {
        UserTemp,
        SystemTemp,
        RecycleBin,
        BrowserCache,
        ThumbnailCache,
        PrefetchAged,
        WindowsUpdateCache,
        DeliveryOptimization,
        InstallerLeftover
    }

    internal static class SafePathHelper
    {
        public static bool IsPathInsideAny(string fullPath, IEnumerable<string> roots)
        {
            var path = Normalize(fullPath);
            foreach (var root in roots ?? Enumerable.Empty<string>())
            {
                if (IsPathInside(path, root))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsPathInside(string fullPath, string root)
        {
            var path = Normalize(fullPath);
            var baseRoot = Normalize(root);
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(baseRoot))
            {
                return false;
            }

            if (path.Equals(baseRoot, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return path.StartsWith(baseRoot + "\\", StringComparison.OrdinalIgnoreCase);
        }

        public static IEnumerable<string> EnumerateFilesSafe(string root, string pattern = "*")
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                yield break;
            }

            var pending = new Stack<string>();
            pending.Push(root);
            while (pending.Count > 0)
            {
                var current = pending.Pop();

                string[] directories;
                try { directories = Directory.GetDirectories(current); }
                catch { directories = Array.Empty<string>(); }

                foreach (var dir in directories)
                {
                    pending.Push(dir);
                }

                string[] files;
                try { files = Directory.GetFiles(current, pattern, SearchOption.TopDirectoryOnly); }
                catch { files = Array.Empty<string>(); }

                foreach (var file in files)
                {
                    yield return file;
                }
            }
        }

        public static IEnumerable<string> EnumerateDirectoriesSafe(string root)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                yield break;
            }

            var pending = new Stack<string>();
            pending.Push(root);
            while (pending.Count > 0)
            {
                var current = pending.Pop();
                string[] directories;
                try { directories = Directory.GetDirectories(current); }
                catch { directories = Array.Empty<string>(); }

                foreach (var dir in directories)
                {
                    yield return dir;
                    pending.Push(dir);
                }
            }
        }

        public static long CalculateDirectorySize(string root)
        {
            long total = 0;
            foreach (var file in EnumerateFilesSafe(root))
            {
                try { total += new FileInfo(file).Length; }
                catch { }
            }

            return total;
        }

        private static string Normalize(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(path).TrimEnd('\\');
            }
            catch
            {
                return path.Trim().TrimEnd('\\');
            }
        }
    }
}
