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
        private const int JunkScanStepCount = 11;

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
            var safeRoots = BuildSafeRoots();
            var step = 0;
            var now = DateTime.Now;
            var userTempThreshold = now.AddHours(-24);
            var systemTempThreshold = now.AddDays(-7);
            var prefetchThreshold = now.AddDays(-30);
            var errorReportThreshold = now.AddDays(-14);
            var crashDumpThreshold = now.AddDays(-30);

            var userTempRoots = new[]
            {
                Path.GetTempPath(),
                Path.Combine(LocalAppData, "Temp")
            }
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(Directory.Exists)
            .Where(root => SafePathHelper.IsPathInsideAny(root, safeRoots))
            .ToList();

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描用户临时目录...");
            foreach (var root in userTempRoots)
            {
                ScanAgedFiles(
                    results,
                    safeRoots,
                    root,
                    JunkCategory.UserTemp,
                    userTempThreshold,
                    "24 小时前临时文件",
                    "仅限当前用户 Temp；跳过 24 小时内文件和 junction/symlink。",
                    true,
                    ct);
            }

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描系统临时目录...");
            ScanAgedFiles(
                results,
                safeRoots,
                Path.Combine(WindowsRoot, "Temp"),
                JunkCategory.SystemTemp,
                systemTempThreshold,
                "7 天前系统临时文件",
                "需 UAC；仅限 Windows\\Temp；跳过 7 天内文件和 junction/symlink。",
                true,
                ct);

            progress?.Report($"[{++step}/{JunkScanStepCount}] 读取回收站容量...");
            var recycleInfo = NativeMethods.QueryRecycleBin();
            if (recycleInfo.bytes > 0)
            {
                results.Add(new JunkCandidate
                {
                    Path = "RecycleBin",
                    Category = JunkCategory.RecycleBin,
                    Bytes = recycleInfo.bytes,
                    Reason = "回收站可释放空间",
                    SafetyBoundary = "通过 Shell 回收站 API 读取和清空；不枚举用户目录。",
                    Evidence = $"回收站项目约 {recycleInfo.items} 项；估算 {FileSizeFormatter.Format(recycleInfo.bytes)}。",
                    IsSelected = false
                });
            }

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描浏览器缓存...");
            var edgeRunning = Process.GetProcessesByName("msedge").Length > 0;
            var chromeRunning = Process.GetProcessesByName("chrome").Length > 0;
            var browserRunning = edgeRunning || chromeRunning;
            foreach (var cachePath in BuildBrowserCacheRoots(safeRoots))
            {
                ct.ThrowIfCancellationRequested();
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
                    SafetyBoundary = "仅限 Edge/Chrome 用户数据下的缓存目录；浏览器运行时跳过；跳过 junction/symlink。",
                    Evidence = $"扫描目录：{cachePath}；估算 {FileSizeFormatter.Format(bytes)}。",
                    IsSelected = false
                });
            }

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描缩略图缓存...");
            ScanAllFiles(
                results,
                safeRoots,
                Path.Combine(LocalAppData, "Microsoft", "Windows", "Explorer"),
                "thumbcache_*.db",
                JunkCategory.ThumbnailCache,
                "缩略图缓存文件",
                "仅限 Explorer 缩略图缓存 thumbcache_*.db；跳过 junction/symlink。",
                true,
                ct);

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描 Prefetch...");
            ScanAgedFiles(
                results,
                safeRoots,
                Path.Combine(WindowsRoot, "Prefetch"),
                JunkCategory.PrefetchAged,
                prefetchThreshold,
                "30 天前 Prefetch 缓存",
                "仅限 Windows\\Prefetch；跳过 30 天内文件和 junction/symlink。",
                true,
                ct,
                "*.pf");

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描 Windows 更新缓存...");
            ScanAllFiles(
                results,
                safeRoots,
                Path.Combine(WindowsRoot, "SoftwareDistribution", "Download"),
                "*",
                JunkCategory.WindowsUpdateCache,
                "Windows Update 缓存",
                "需 UAC；仅限 SoftwareDistribution\\Download；脚本内二次校验根目录。",
                true,
                ct);

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描 Delivery Optimization 缓存...");
            ScanAllFiles(
                results,
                safeRoots,
                Path.Combine(WindowsRoot, "ServiceProfiles", "NetworkService", "AppData", "Local", "Microsoft", "Windows", "DeliveryOptimization", "Cache"),
                "*",
                JunkCategory.DeliveryOptimization,
                "Delivery Optimization 缓存",
                "需 UAC；仅限 DeliveryOptimization\\Cache；脚本内二次校验根目录。",
                true,
                ct);

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描 Windows 错误报告...");
            foreach (var root in BuildWindowsErrorReportRoots(safeRoots))
            {
                ScanAgedFiles(
                    results,
                    safeRoots,
                    root,
                    JunkCategory.WindowsErrorReports,
                    errorReportThreshold,
                    "14 天前 Windows 错误报告",
                    "仅限当前用户 WER ReportArchive/ReportQueue；默认不勾选；跳过 junction/symlink。",
                    false,
                    ct);
            }

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描崩溃转储...");
            foreach (var root in BuildCrashDumpRoots(safeRoots))
            {
                ScanAgedFiles(
                    results,
                    safeRoots,
                    root,
                    JunkCategory.CrashDumps,
                    crashDumpThreshold,
                    "30 天前崩溃转储",
                    "仅限当前用户 CrashDumps；默认不勾选；跳过 junction/symlink。",
                    false,
                    ct,
                    "*.dmp");
            }

            progress?.Report($"[{++step}/{JunkScanStepCount}] 扫描 Installer 空目录...");
            var packageCachePath = Path.Combine(LocalAppData, "Package Cache");
            if (Directory.Exists(packageCachePath) && SafePathHelper.IsPathInsideAny(packageCachePath, safeRoots))
            {
                foreach (var directory in SafePathHelper.EnumerateDirectoriesSafe(packageCachePath))
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        if (!SafePathHelper.IsPathInsideAny(directory, safeRoots)
                            || SafePathHelper.IsReparsePoint(directory)
                            || Directory.EnumerateFileSystemEntries(directory).Any())
                        {
                            continue;
                        }

                        results.Add(new JunkCandidate
                        {
                            Path = directory,
                            Category = JunkCategory.InstallerLeftover,
                            Bytes = 0,
                            Reason = "空目录（仅删除空目录）",
                            SafetyBoundary = "仅限 LocalAppData\\Package Cache 下空目录；非空目录和 junction/symlink 跳过。",
                            Evidence = $"空目录：{directory}。",
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
                            step.BytesFreed = await DeleteCandidatePathAsync(candidate.Path, safeRoots).ConfigureAwait(false);
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
                    await RunAdminCleanupScriptAsync(pair.Key, pair.Value, ct).ConfigureAwait(false);
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

        private static async Task RunAdminCleanupScriptAsync(JunkCategory category, IReadOnlyCollection<JunkCandidate> candidates, CancellationToken ct)
        {
            var targetArray = BuildPowerShellStringArray((candidates ?? Array.Empty<JunkCandidate>())
                .Where(x => x != null)
                .Select(x => x.Path));
            if (string.Equals(targetArray, "@()", StringComparison.Ordinal))
            {
                return;
            }

            string script;
            switch (category)
            {
                case JunkCategory.SystemTemp:
                    script = BuildAdminCleanupScriptPrelude() + @"
$target = Join-Path $env:windir 'Temp'
$threshold = (Get-Date).AddDays(-7)
$targets = " + targetArray + @"
Invoke-SafeFileCleanup -Root $target -Targets $targets -Threshold $threshold";
                    break;
                case JunkCategory.WindowsUpdateCache:
                    script = BuildAdminCleanupScriptPrelude() + @"
$target = Join-Path $env:windir 'SoftwareDistribution\Download'
$targets = " + targetArray + @"
$service = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
$wasRunning = ($null -ne $service -and $service.Status -eq 'Running')
if ($null -ne $service) {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
}
Invoke-SafeFileCleanup -Root $target -Targets $targets -Threshold $null
if ($wasRunning) {
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}";
                    break;
                case JunkCategory.DeliveryOptimization:
                    script = BuildAdminCleanupScriptPrelude() + @"
$target = Join-Path $env:windir 'ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache'
$targets = " + targetArray + @"
$service = Get-Service -Name dosvc -ErrorAction SilentlyContinue
$wasRunning = ($null -ne $service -and $service.Status -eq 'Running')
if ($null -ne $service) {
    Stop-Service -Name dosvc -Force -ErrorAction SilentlyContinue
}
Invoke-SafeFileCleanup -Root $target -Targets $targets -Threshold $null
if ($wasRunning) {
    Start-Service -Name dosvc -ErrorAction SilentlyContinue
}";
                    break;
                default:
                    return;
            }

            await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
        }

        private static string BuildPowerShellStringArray(IEnumerable<string> values)
        {
            var items = (values ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(ToPowerShellSingleQuotedString)
                .ToList();
            return "@(" + string.Join(", ", items) + ")";
        }

        private static string ToPowerShellSingleQuotedString(string value)
        {
            return "'" + (value ?? string.Empty).Replace("'", "''") + "'";
        }

        private static string BuildAdminCleanupScriptPrelude()
        {
            return @"
function Get-FullPath($path) {
    try {
        return [System.IO.Path]::GetFullPath($path).TrimEnd('\')
    } catch {
        return ''
    }
}

function Test-PathInsideRoot($path, $root) {
    $fullPath = Get-FullPath $path
    $fullRoot = Get-FullPath $root
    if ($fullPath.Length -eq 0 -or $fullRoot.Length -eq 0) { return $false }
    if ([string]::Equals($fullPath, $fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $fullPath.StartsWith($fullRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-ReparsePoint($item) {
    try {
        return (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint)
    } catch {
        return $true
    }
}

function Invoke-SafeFileCleanup($Root, $Targets, $Threshold) {
    $safeRoot = Get-FullPath $Root
    if ($safeRoot.Length -eq 0 -or -not (Test-Path -LiteralPath $safeRoot)) { return }
    $rootItem = Get-Item -LiteralPath $safeRoot -Force -ErrorAction SilentlyContinue
    if ($null -eq $rootItem -or (Test-ReparsePoint $rootItem)) { return }

    foreach ($targetPath in $Targets) {
        try {
            $fullTarget = Get-FullPath $targetPath
            if ($fullTarget.Length -eq 0 -or -not (Test-PathInsideRoot $fullTarget $safeRoot)) { continue }
            $item = Get-Item -LiteralPath $fullTarget -Force -ErrorAction SilentlyContinue
            if ($null -eq $item -or $item.PSIsContainer -or (Test-ReparsePoint $item)) { continue }
            if ($null -ne $Threshold -and $item.LastWriteTime -ge $Threshold) { continue }
            Remove-Item -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue
        } catch { }
    }

    Get-ChildItem -LiteralPath $safeRoot -Force -Recurse -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending |
    ForEach-Object {
        try {
            if ($_.PSIsContainer -and -not (Test-ReparsePoint $_) -and (Test-PathInsideRoot $_.FullName $safeRoot)) {
                if ((Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                }
            }
        } catch { }
    }
}
";
        }

        private static async Task<long> DeleteCandidatePathAsync(string fullPath, IEnumerable<string> safeRoots)
        {
            if (File.Exists(fullPath))
            {
                var info = new FileInfo(fullPath);
                if (SafePathHelper.IsReparsePoint(info.FullName) || !SafePathHelper.IsPathInsideAny(info.FullName, safeRoots))
                {
                    return 0;
                }

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
                if (SafePathHelper.IsReparsePoint(fullPath) || !SafePathHelper.IsPathInsideAny(fullPath, safeRoots))
                {
                    return 0;
                }

                var bytes = 0L;
                foreach (var file in SafePathHelper.EnumerateFilesSafe(fullPath))
                {
                    if (!SafePathHelper.IsPathInsideAny(file, safeRoots) || SafePathHelper.IsReparsePoint(file))
                    {
                        continue;
                    }

                    var info = new FileInfo(file);
                    var length = info.Length;
                    AppLogService.Information("Junk cleanup plan: recycle file {Path}", info.FullName);
                    info.Attributes = FileAttributes.Normal;
                    VbFileIO.FileSystem.DeleteFile(info.FullName, VbFileIO.UIOption.OnlyErrorDialogs, VbFileIO.RecycleOption.SendToRecycleBin, VbFileIO.UICancelOption.ThrowException);
                    bytes += length;
                }

                DeleteEmptyDirectoriesSafe(fullPath, safeRoots);
                await Task.Yield();
                return bytes;
            }

            return 0;
        }

        private static void DeleteEmptyDirectoriesSafe(string root, IEnumerable<string> safeRoots)
        {
            var directories = SafePathHelper.EnumerateDirectoriesSafe(root)
                .OrderByDescending(x => x.Length)
                .ToList();

            foreach (var directory in directories)
            {
                try
                {
                    if (SafePathHelper.IsReparsePoint(directory) || !SafePathHelper.IsPathInsideAny(directory, safeRoots))
                    {
                        continue;
                    }

                    if (!Directory.EnumerateFileSystemEntries(directory).Any())
                    {
                        Directory.Delete(directory, false);
                    }
                }
                catch
                {
                }
            }
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
                Path.Combine(LocalAppData, "Package Cache"),
                Path.Combine(LocalAppData, "Microsoft", "Windows", "WER", "ReportArchive"),
                Path.Combine(LocalAppData, "Microsoft", "Windows", "WER", "ReportQueue"),
                Path.Combine(LocalAppData, "CrashDumps")
            };

            roots.AddRange(BuildBrowserCacheRootCandidates());

            return roots
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(path => !ForbiddenRoots.Any(forbidden => SafePathHelper.IsPathInside(path, forbidden)))
                .ToList();
        }

        private static IEnumerable<string> BuildBrowserCacheRoots(IEnumerable<string> safeRoots)
        {
            return BuildBrowserCacheRootCandidates()
                .Where(Directory.Exists)
                .Where(path => SafePathHelper.IsPathInsideAny(path, safeRoots))
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> BuildBrowserCacheRootCandidates()
        {
            foreach (var browserRoot in new[]
            {
                Path.Combine(LocalAppData, "Microsoft", "Edge", "User Data"),
                Path.Combine(LocalAppData, "Google", "Chrome", "User Data")
            })
            {
                if (string.IsNullOrWhiteSpace(browserRoot) || !Directory.Exists(browserRoot))
                {
                    continue;
                }

                foreach (var profileRoot in EnumerateBrowserProfileRoots(browserRoot))
                {
                    yield return Path.Combine(profileRoot, "Cache", "Cache_Data");
                    yield return Path.Combine(profileRoot, "Code Cache");
                    yield return Path.Combine(profileRoot, "GPUCache");
                    yield return Path.Combine(profileRoot, "Service Worker", "CacheStorage");
                }
            }
        }

        private static IEnumerable<string> EnumerateBrowserProfileRoots(string browserRoot)
        {
            var defaultProfile = Path.Combine(browserRoot, "Default");
            if (Directory.Exists(defaultProfile) && !SafePathHelper.IsReparsePoint(defaultProfile))
            {
                yield return defaultProfile;
            }

            string[] profileRoots;
            try
            {
                profileRoots = Directory.GetDirectories(browserRoot, "Profile *", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                profileRoots = Array.Empty<string>();
            }

            foreach (var profileRoot in profileRoots)
            {
                if (IsLikelyBrowserProfileRoot(profileRoot) && !SafePathHelper.IsReparsePoint(profileRoot))
                {
                    yield return profileRoot;
                }
            }
        }

        private static bool IsLikelyBrowserProfileRoot(string path)
        {
            var name = string.IsNullOrWhiteSpace(path) ? string.Empty : Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            return string.Equals(name, "Default", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("Profile ", StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> BuildWindowsErrorReportRoots(IEnumerable<string> safeRoots)
        {
            return new[]
            {
                Path.Combine(LocalAppData, "Microsoft", "Windows", "WER", "ReportArchive"),
                Path.Combine(LocalAppData, "Microsoft", "Windows", "WER", "ReportQueue")
            }
            .Where(Directory.Exists)
            .Where(path => SafePathHelper.IsPathInsideAny(path, safeRoots));
        }

        private static IEnumerable<string> BuildCrashDumpRoots(IEnumerable<string> safeRoots)
        {
            return new[]
            {
                Path.Combine(LocalAppData, "CrashDumps")
            }
            .Where(Directory.Exists)
            .Where(path => SafePathHelper.IsPathInsideAny(path, safeRoots));
        }

        private static void ScanAgedFiles(
            ICollection<JunkCandidate> target,
            IEnumerable<string> safeRoots,
            string root,
            JunkCategory category,
            DateTime threshold,
            string reason,
            string safetyBoundary,
            bool selected,
            CancellationToken ct,
            string pattern = "*")
        {
            if (string.IsNullOrWhiteSpace(root)
                || !Directory.Exists(root)
                || !SafePathHelper.IsPathInsideAny(root, safeRoots)
                || SafePathHelper.IsReparsePoint(root))
            {
                return;
            }

            foreach (var file in SafePathHelper.EnumerateFilesSafe(root, pattern))
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var info = new FileInfo(file);
                    if (SafePathHelper.IsReparsePoint(info.FullName)
                        || !SafePathHelper.IsPathInsideAny(info.FullName, safeRoots)
                        || info.LastWriteTime >= threshold)
                    {
                        continue;
                    }

                    target.Add(new JunkCandidate
                    {
                        Path = info.FullName,
                        Category = category,
                        Bytes = info.Length,
                        Reason = reason,
                        SafetyBoundary = safetyBoundary,
                        Evidence = BuildFileEvidence(root, info, threshold),
                        IsSelected = selected
                    });
                }
                catch
                {
                }
            }
        }

        private static void ScanAllFiles(
            ICollection<JunkCandidate> target,
            IEnumerable<string> safeRoots,
            string root,
            string pattern,
            JunkCategory category,
            string reason,
            string safetyBoundary,
            bool selected,
            CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(root)
                || !Directory.Exists(root)
                || !SafePathHelper.IsPathInsideAny(root, safeRoots)
                || SafePathHelper.IsReparsePoint(root))
            {
                return;
            }

            foreach (var file in SafePathHelper.EnumerateFilesSafe(root, pattern))
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var info = new FileInfo(file);
                    if (SafePathHelper.IsReparsePoint(info.FullName) || !SafePathHelper.IsPathInsideAny(info.FullName, safeRoots))
                    {
                        continue;
                    }

                    target.Add(new JunkCandidate
                    {
                        Path = info.FullName,
                        Category = category,
                        Bytes = info.Length,
                        Reason = reason,
                        SafetyBoundary = safetyBoundary,
                        Evidence = BuildFileEvidence(root, info, null),
                        IsSelected = selected
                    });
                }
                catch
                {
                }
            }
        }

        private static string BuildFileEvidence(string root, FileInfo info, DateTime? threshold)
        {
            var thresholdText = threshold.HasValue
                ? $"阈值：早于 {threshold.Value:yyyy-MM-dd HH:mm}；"
                : string.Empty;
            return $"{thresholdText}扫描根：{root}；文件时间：{info.LastWriteTime:yyyy-MM-dd HH:mm}；大小：{FileSizeFormatter.Format(info.Length)}。";
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
        public string SafetyBoundary { get; set; }
        public string Evidence { get; set; }

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
                case JunkCategory.WindowsErrorReports: return "错误报告";
                case JunkCategory.CrashDumps: return "崩溃转储";
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
                case JunkCategory.WindowsErrorReports:
                    return "中";
                case JunkCategory.CrashDumps:
                    return "高";
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
                case JunkCategory.WindowsErrorReports:
                    return "确认不需要排障后清理";
                case JunkCategory.CrashDumps:
                    return "高风险，排障结束后再清理";
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
        InstallerLeftover,
        WindowsErrorReports,
        CrashDumps
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
            var safeRoot = Normalize(root);
            if (string.IsNullOrWhiteSpace(safeRoot) || !Directory.Exists(safeRoot) || IsReparsePoint(safeRoot))
            {
                yield break;
            }

            var pending = new Stack<string>();
            pending.Push(safeRoot);
            while (pending.Count > 0)
            {
                var current = pending.Pop();

                string[] directories;
                try { directories = Directory.GetDirectories(current); }
                catch { directories = Array.Empty<string>(); }

                foreach (var dir in directories)
                {
                    if (IsPathInside(dir, safeRoot) && !IsReparsePoint(dir))
                    {
                        pending.Push(dir);
                    }
                }

                string[] files;
                try { files = Directory.GetFiles(current, pattern, SearchOption.TopDirectoryOnly); }
                catch { files = Array.Empty<string>(); }

                foreach (var file in files)
                {
                    if (IsPathInside(file, safeRoot) && !IsReparsePoint(file))
                    {
                        yield return file;
                    }
                }
            }
        }

        public static IEnumerable<string> EnumerateDirectoriesSafe(string root)
        {
            var safeRoot = Normalize(root);
            if (string.IsNullOrWhiteSpace(safeRoot) || !Directory.Exists(safeRoot) || IsReparsePoint(safeRoot))
            {
                yield break;
            }

            var pending = new Stack<string>();
            pending.Push(safeRoot);
            while (pending.Count > 0)
            {
                var current = pending.Pop();
                string[] directories;
                try { directories = Directory.GetDirectories(current); }
                catch { directories = Array.Empty<string>(); }

                foreach (var dir in directories)
                {
                    if (!IsPathInside(dir, safeRoot) || IsReparsePoint(dir))
                    {
                        continue;
                    }

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

        public static bool IsReparsePoint(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return true;
            }

            try
            {
                if (File.Exists(path))
                {
                    return (new FileInfo(path).Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint;
                }

                if (Directory.Exists(path))
                {
                    return (new DirectoryInfo(path).Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint;
                }
            }
            catch
            {
                return true;
            }

            return false;
        }
    }
}
