using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public sealed class SystemOptimizationService
    {
        private const int MaxProcessWaitMs = 120000;
        private const string StepKeyUserTemp = "UserTemp";
        private const string StepKeyWindowsTemp = "WindowsTemp";
        private const string StepKeyThumbnailCache = "ThumbnailCache";
        private const string StepKeyDnsCache = "DnsCache";
        private const string StepKeyFontCache = "FontCache";
        private const string StepKeyWindowsUpdateCache = "WindowsUpdateCache";
        private const string StepKeyDeliveryOptimizationCache = "DeliveryOptimizationCache";
        private const string StepKeyWindowsErrorReports = "WindowsErrorReports";
        private const string StepKeyCrashDumps = "CrashDumps";
        private const string StepKeySystemDrive = "SystemDrive";
        private const string StepKeyComponentStoreCleanup = "ComponentStoreCleanup";
        private const string StepKeyRecycleBin = "RecycleBin";
        private const string StepKeyEventLogs = "EventLogs";
        private const string StepKeyWorkingSet = "WorkingSet";

        private const string RiskLow = "低";
        private const string RiskMedium = "中";
        private const string RiskHigh = "高";

        public bool AllowExplorerRestartForThumbnailCleanup { get; set; }

        public async Task<OptimizationReportItem> RunAsync(IProgress<string> progress, CancellationToken ct)
        {
            return await RunAsync(progress, ct, null).ConfigureAwait(false);
        }

        public async Task<OptimizationReportItem> RunAsync(IProgress<string> progress, CancellationToken ct, IReadOnlyCollection<string> enabledStepKeys)
        {
            var report = new OptimizationReportItem
            {
                Id = Guid.NewGuid().ToString("N").Substring(0, 8),
                StartedAt = DateTime.Now,
                ReportType = "AutoOptimize",
                Steps = new List<OptimizationStep>()
            };

            var stepDefinitions = BuildStepDefinitions();
            if (enabledStepKeys != null)
            {
                var enabled = new HashSet<string>(
                    enabledStepKeys.Where(x => !string.IsNullOrWhiteSpace(x)),
                    StringComparer.OrdinalIgnoreCase);
                stepDefinitions = stepDefinitions
                    .Where(x => enabled.Contains(x.Key))
                    .ToList();
            }

            for (var i = 0; i < stepDefinitions.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                var stepName = stepDefinitions[i].Name;
                progress?.Report($"[{i + 1}/{stepDefinitions.Count}] {stepName}...");

                OptimizationStep result;
                var startedAt = Stopwatch.StartNew();
                try
                {
                    result = await stepDefinitions[i].Execute(ct).ConfigureAwait(false);
                    if (result == null)
                    {
                        result = new OptimizationStep
                        {
                            Name = stepName,
                            Status = "Skipped",
                            Detail = "未返回执行结果。",
                            BytesFreed = 0
                        };
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    AppLogService.Error(ex, "Auto optimize step failed: {StepName}", stepName);
                    result = new OptimizationStep
                    {
                        Name = stepName,
                        Status = "Failed",
                        Detail = ex.Message,
                        BytesFreed = 0
                    };
                }
                finally
                {
                    startedAt.Stop();
                }

                result.Name = string.IsNullOrWhiteSpace(result.Name) ? stepName : result.Name;
                result.Duration = startedAt.Elapsed;
                report.Steps.Add(result);
            }

            report.FinishedAt = DateTime.Now;
            report.TotalBytesFreed = report.Steps.Sum(x => x.BytesFreed);
            var ok = report.Steps.Count(x => string.Equals(x.Status, "OK", StringComparison.OrdinalIgnoreCase));
            var failed = report.Steps.Count(x => string.Equals(x.Status, "Failed", StringComparison.OrdinalIgnoreCase));
            var skipped = report.Steps.Count(x => string.Equals(x.Status, "Skipped", StringComparison.OrdinalIgnoreCase));
            report.Summary = $"成功 {ok} 步，失败 {failed} 步，跳过 {skipped} 步。";
            return report;
        }

        public async Task<IReadOnlyList<OptimizationPlanItem>> ScanAsync(IProgress<string> progress, CancellationToken ct)
        {
            var items = new List<OptimizationPlanItem>();
            const int totalSteps = 14;
            var step = 0;

            progress?.Report($"[{++step}/{totalSteps}] 扫描当前用户 Temp...");
            var userTempPath = Path.GetTempPath();
            var userTempSafe = IsSafeCleanupRoot(userTempPath, userTempPath);
            var userTemp = userTempSafe
                ? CountDeletableFiles(userTempPath, DateTime.Now.AddHours(-24))
                : (count: 0, bytes: 0L);
            items.Add(CreatePlanItem(
                StepKeyUserTemp,
                "清空当前用户 Temp（24 小时前）",
                "磁盘清理",
                false,
                userTemp.bytes,
                userTempSafe && userTemp.count > 0,
                userTempSafe
                    ? (userTemp.count > 0
                    ? $"发现 {userTemp.count} 个超过 24 小时的临时文件。"
                    : "未发现超过 24 小时的用户临时文件。")
                    : "当前用户 Temp 路径未通过白名单校验，已禁止执行。",
                RiskLow,
                "仅限当前用户 Temp；跳过 24 小时内文件。",
                BuildCountEvidence(userTempPath, userTemp.count, userTemp.bytes, "24 小时前"),
                true));

            progress?.Report($"[{++step}/{totalSteps}] 扫描 Windows Temp...");
            var windowsTempPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
            var windowsTempSafe = IsSafeCleanupRoot(windowsTempPath, windowsTempPath);
            var windowsTemp = windowsTempSafe
                ? CountDeletableFiles(windowsTempPath, DateTime.Now.AddDays(-7))
                : (count: 0, bytes: 0L);
            items.Add(CreatePlanItem(
                StepKeyWindowsTemp,
                "清空 Windows Temp（7 天前）",
                "磁盘清理",
                true,
                windowsTemp.bytes,
                windowsTempSafe && windowsTemp.count > 0,
                windowsTempSafe
                    ? (windowsTemp.count > 0
                    ? $"发现 {windowsTemp.count} 个超过 7 天的系统临时文件。"
                    : "未发现超过 7 天的系统临时文件。")
                    : "Windows Temp 路径未通过白名单校验，已禁止执行。",
                RiskLow,
                "需 UAC；仅限 Windows\\Temp；跳过 7 天内文件。",
                BuildCountEvidence(windowsTempPath, windowsTemp.count, windowsTemp.bytes, "7 天前"),
                true));

            progress?.Report($"[{++step}/{totalSteps}] 扫描缩略图缓存...");
            var explorerCacheRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft",
                "Windows",
                "Explorer");
            var explorerCacheSafe = IsSafeCleanupRoot(explorerCacheRoot, explorerCacheRoot);
            var thumbnailFiles = explorerCacheSafe ? CountFiles(explorerCacheRoot, "thumbcache_*.db") : (count: 0, bytes: 0L);
            items.Add(CreatePlanItem(
                StepKeyThumbnailCache,
                "清理缩略图缓存",
                "缓存清理",
                false,
                thumbnailFiles.bytes,
                explorerCacheSafe && thumbnailFiles.count > 0,
                explorerCacheSafe
                    ? (thumbnailFiles.count > 0
                    ? $"发现 {thumbnailFiles.count} 个缩略图缓存文件，优化时需要临时重启资源管理器。"
                    : "未发现可清理的缩略图缓存文件。")
                    : "缩略图缓存路径未通过白名单校验，已禁止执行。",
                RiskMedium,
                "需用户确认重启 Explorer；仅限 thumbcache_*.db。",
                BuildCountEvidence(explorerCacheRoot, thumbnailFiles.count, thumbnailFiles.bytes, "thumbcache_*.db"),
                false));

            progress?.Report($"[{++step}/{totalSteps}] 检查 DNS 缓存刷新项...");
            items.Add(CreatePlanItem(
                StepKeyDnsCache,
                "刷新 DNS 缓存",
                "网络维护",
                false,
                0,
                true,
                "可刷新 DNS 解析缓存，不释放磁盘空间。",
                RiskLow,
                "只调用 ipconfig /flushdns，不删除文件。",
                "无文件扫描；执行命令为 ipconfig /flushdns。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 扫描字体缓存...");
            var fontCachePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "ServiceProfiles",
                "LocalService",
                "AppData",
                "Local",
                "FontCache");
            var fontCacheSafe = IsSafeCleanupRoot(fontCachePath, fontCachePath);
            var fontCacheFiles = fontCacheSafe ? CountFiles(fontCachePath, "*.dat") : (count: 0, bytes: 0L);
            items.Add(CreatePlanItem(
                StepKeyFontCache,
                "重置字体缓存",
                "缓存清理",
                true,
                fontCacheFiles.bytes,
                fontCacheSafe && fontCacheFiles.count > 0,
                fontCacheSafe
                    ? (fontCacheFiles.count > 0
                    ? $"发现 {fontCacheFiles.count} 个字体缓存文件。"
                    : "未发现可清理的字体缓存文件。")
                    : "字体缓存路径未通过白名单校验，已禁止执行。",
                RiskMedium,
                "需 UAC；会临时停止 FontCache 服务。",
                BuildCountEvidence(fontCachePath, fontCacheFiles.count, fontCacheFiles.bytes, "*.dat"),
                false));

            progress?.Report($"[{++step}/{totalSteps}] 扫描 Windows 更新缓存...");
            var updateCacheRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "SoftwareDistribution",
                "Download");
            var updateCacheSafe = IsSafeCleanupRoot(updateCacheRoot, updateCacheRoot);
            var updateCacheBytes = updateCacheSafe ? CalculateDirectorySize(updateCacheRoot) : 0L;
            items.Add(CreatePlanItem(
                StepKeyWindowsUpdateCache,
                "清理 Windows 更新缓存",
                "更新缓存",
                true,
                updateCacheBytes,
                updateCacheSafe && updateCacheBytes > 0,
                updateCacheSafe
                    ? (updateCacheBytes > 0
                    ? "发现 Windows Update 下载缓存。"
                    : "未发现 Windows Update 下载缓存。")
                    : "Windows Update 下载缓存路径未通过白名单校验，已禁止执行。",
                RiskMedium,
                "需 UAC；仅限 SoftwareDistribution\\Download。",
                BuildCountEvidence(updateCacheRoot, 0, updateCacheBytes, "全部缓存项"),
                updateCacheBytes >= 100L * 1024L * 1024L));

            progress?.Report($"[{++step}/{totalSteps}] 扫描 Delivery Optimization 缓存...");
            var deliveryCacheRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "ServiceProfiles",
                "NetworkService",
                "AppData",
                "Local",
                "Microsoft",
                "Windows",
                "DeliveryOptimization",
                "Cache");
            var deliveryCacheSafe = IsSafeCleanupRoot(deliveryCacheRoot, deliveryCacheRoot);
            var deliveryCacheBytes = deliveryCacheSafe ? CalculateDirectorySize(deliveryCacheRoot) : 0L;
            items.Add(CreatePlanItem(
                StepKeyDeliveryOptimizationCache,
                "清理传递优化缓存",
                "更新缓存",
                true,
                deliveryCacheBytes,
                deliveryCacheSafe && deliveryCacheBytes > 0,
                deliveryCacheSafe
                    ? (deliveryCacheBytes > 0
                        ? "发现 Delivery Optimization 下载缓存。"
                        : "未发现 Delivery Optimization 下载缓存。")
                    : "Delivery Optimization 缓存路径未通过白名单校验，已禁止执行。",
                RiskMedium,
                "需 UAC；仅限 DeliveryOptimization\\Cache。",
                BuildCountEvidence(deliveryCacheRoot, 0, deliveryCacheBytes, "全部缓存项"),
                deliveryCacheBytes >= 100L * 1024L * 1024L));

            progress?.Report($"[{++step}/{totalSteps}] 扫描 Windows 错误报告...");
            var wer = CountWindowsErrorReports(DateTime.Now.AddDays(-14));
            items.Add(CreatePlanItem(
                StepKeyWindowsErrorReports,
                "清理 Windows 错误报告（14 天前）",
                "诊断清理",
                true,
                wer.bytes,
                wer.count > 0,
                wer.count > 0
                    ? $"发现 {wer.count} 个 14 天前的 Windows 错误报告文件。"
                    : "未发现 14 天前的 Windows 错误报告文件。",
                RiskMedium,
                "需 UAC；仅限 WER ReportArchive/ReportQueue。",
                $"白名单路径 {wer.roots} 个；文件 {wer.count} 个；阈值 14 天前。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 扫描崩溃转储文件...");
            var crashDumps = CountCrashDumps(DateTime.Now.AddDays(-30));
            items.Add(CreatePlanItem(
                StepKeyCrashDumps,
                "清理崩溃转储（30 天前）",
                "诊断清理",
                true,
                crashDumps.bytes,
                crashDumps.count > 0,
                crashDumps.count > 0
                    ? $"发现 {crashDumps.count} 个 30 天前的崩溃转储文件。"
                    : "未发现 30 天前的崩溃转储文件。",
                RiskHigh,
                "需 UAC；会删除旧 .dmp，可能影响故障追溯。",
                $"白名单路径 {crashDumps.roots} 个；文件 {crashDumps.count} 个；阈值 30 天前。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 检查系统盘优化项...");
            var systemDrive = Environment.GetEnvironmentVariable("SystemDrive");
            var driveLetter = string.IsNullOrWhiteSpace(systemDrive)
                ? string.Empty
                : systemDrive.Trim().TrimEnd('\\').Replace(":", string.Empty);
            var mediaType = string.IsNullOrWhiteSpace(driveLetter)
                ? "Unspecified"
                : await QuerySystemDriveMediaTypeAsync(driveLetter, ct).ConfigureAwait(false);
            var canOptimizeDrive = mediaType.Equals("SSD", StringComparison.OrdinalIgnoreCase)
                || mediaType.Equals("HDD", StringComparison.OrdinalIgnoreCase);
            items.Add(CreatePlanItem(
                StepKeySystemDrive,
                "系统盘 TRIM/碎片整理",
                "磁盘维护",
                true,
                0,
                canOptimizeDrive,
                canOptimizeDrive
                    ? (mediaType.Equals("SSD", StringComparison.OrdinalIgnoreCase)
                        ? "系统盘识别为 SSD，可执行 TRIM。"
                        : "系统盘识别为 HDD，可执行碎片整理。")
                    : $"无法识别系统盘类型（{mediaType ?? "Unspecified"}）。",
                mediaType.Equals("HDD", StringComparison.OrdinalIgnoreCase) ? RiskHigh : RiskMedium,
                "需 UAC；仅对系统盘执行 Optimize-Volume。",
                $"系统盘 {driveLetter}:；MediaType={mediaType ?? "Unspecified"}。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 检查组件存储清理项...");
            var componentCleanupAvailable = IsComponentStoreCleanupAvailable();
            items.Add(CreatePlanItem(
                StepKeyComponentStoreCleanup,
                "组件存储清理（WinSxS）",
                "系统维护",
                true,
                0,
                componentCleanupAvailable,
                componentCleanupAvailable
                    ? "可执行 DISM 组件存储清理，耗时可能较长。"
                    : "当前系统不支持或未找到 DISM 组件存储清理。",
                RiskHigh,
                "需 UAC；执行 DISM StartComponentCleanup，不删除用户文件。",
                componentCleanupAvailable ? "检测到 DISM 和受支持的 Windows 版本。" : "DISM 或系统版本检查未通过。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 检查回收站...");
            var recycleBin = QueryRecycleBin();
            items.Add(CreatePlanItem(
                StepKeyRecycleBin,
                "清空回收站",
                "磁盘清理",
                false,
                recycleBin.bytes,
                recycleBin.count > 0 || recycleBin.bytes > 0,
                recycleBin.count > 0 || recycleBin.bytes > 0
                    ? $"回收站约 {recycleBin.count} 项。"
                    : "回收站为空。",
                RiskHigh,
                "会永久清空回收站；默认不勾选。",
                $"SHQueryRecycleBin：{recycleBin.count} 项，{FileSizeFormatter.Format(recycleBin.bytes)}。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 检查事件日志大小...");
            var largeLogs = await QueryLargeEventLogsAsync(ct).ConfigureAwait(false);
            items.Add(CreatePlanItem(
                StepKeyEventLogs,
                "清理事件日志（大于 500MB）",
                "日志维护",
                true,
                largeLogs.bytes,
                largeLogs.count > 0,
                largeLogs.count > 0
                    ? $"发现 {largeLogs.count} 个超过 500 MB 的事件日志。"
                    : "Application/System/Security 事件日志未超过 500 MB。",
                RiskHigh,
                "需 UAC；仅清理超过 500MB 的 Application/System/Security。",
                $"超过阈值日志 {largeLogs.count} 个；阈值 500MB。",
                false));

            progress?.Report($"[{++step}/{totalSteps}] 检查当前进程工作集...");
            var workingSet = GetCurrentWorkingSetBytes();
            items.Add(CreatePlanItem(
                StepKeyWorkingSet,
                "压缩当前进程工作集",
                "内存维护",
                false,
                0,
                workingSet > 150L * 1024L * 1024L,
                workingSet > 150L * 1024L * 1024L
                    ? $"当前进程工作集约 {FileSizeFormatter.Format(workingSet)}，可尝试压缩。"
                    : $"当前进程工作集约 {FileSizeFormatter.Format(workingSet)}，无需压缩。",
                RiskLow,
                "只调用 EmptyWorkingSet 作用于 MyTools 当前进程。",
                $"当前工作集 {FileSizeFormatter.Format(workingSet)}；阈值 150MB。",
                false));

            return items;
        }

        private List<OptimizationStepDefinition> BuildStepDefinitions()
        {
            return new List<OptimizationStepDefinition>
            {
                new OptimizationStepDefinition(StepKeyUserTemp, "清空当前用户 Temp（24 小时前）", StepClearUserTempAsync),
                new OptimizationStepDefinition(StepKeyWindowsTemp, "清空 Windows Temp（7 天前）", StepClearWindowsTempAsync),
                new OptimizationStepDefinition(StepKeyThumbnailCache, "清理缩略图缓存", StepClearThumbnailCacheAsync),
                new OptimizationStepDefinition(StepKeyDnsCache, "刷新 DNS 缓存", StepFlushDnsAsync),
                new OptimizationStepDefinition(StepKeyFontCache, "重置字体缓存", StepResetFontCacheAsync),
                new OptimizationStepDefinition(StepKeyWindowsUpdateCache, "清理 Windows 更新缓存", StepCleanWindowsUpdateCacheAsync),
                new OptimizationStepDefinition(StepKeyDeliveryOptimizationCache, "清理传递优化缓存", StepCleanDeliveryOptimizationCacheAsync),
                new OptimizationStepDefinition(StepKeyWindowsErrorReports, "清理 Windows 错误报告（14 天前）", StepCleanWindowsErrorReportsAsync),
                new OptimizationStepDefinition(StepKeyCrashDumps, "清理崩溃转储（30 天前）", StepCleanCrashDumpsAsync),
                new OptimizationStepDefinition(StepKeySystemDrive, "系统盘 TRIM/碎片整理", StepOptimizeSystemDriveAsync),
                new OptimizationStepDefinition(StepKeyComponentStoreCleanup, "组件存储清理（WinSxS）", StepComponentStoreCleanupAsync),
                new OptimizationStepDefinition(StepKeyRecycleBin, "清空回收站", StepEmptyRecycleBinAsync),
                new OptimizationStepDefinition(StepKeyEventLogs, "清理事件日志（大于 500MB）", StepCleanEventLogsAsync),
                new OptimizationStepDefinition(StepKeyWorkingSet, "压缩当前进程工作集", StepTrimWorkingSetAsync)
            };
        }

        private static OptimizationPlanItem CreatePlanItem(
            string key,
            string name,
            string category,
            bool requiresAdmin,
            long estimatedBytes,
            bool canOptimize,
            string detail,
            string riskLevel = RiskLow,
            string safetyBoundary = "",
            string evidence = "",
            bool defaultSelected = false)
        {
            return new OptimizationPlanItem
            {
                Key = key,
                Name = name,
                Category = category,
                RequiresAdmin = requiresAdmin,
                EstimatedBytes = estimatedBytes,
                CanOptimize = canOptimize,
                Detail = detail ?? string.Empty,
                RiskLevel = riskLevel ?? RiskLow,
                SafetyBoundary = safetyBoundary ?? string.Empty,
                Evidence = evidence ?? string.Empty,
                IsSelected = canOptimize && defaultSelected
            };
        }

        private async Task<OptimizationStep> StepClearUserTempAsync(CancellationToken ct)
        {
            var tempPath = Path.GetTempPath();
            EnsureSafeCleanupRoot(tempPath, "当前用户 Temp", tempPath);
            var threshold = DateTime.Now.AddHours(-24);
            var planned = CountDeletableFiles(tempPath, threshold);
            AppLogService.Information("Auto optimize plan: clear user temp path {Path}, files {Count}, bytes {Bytes}",
                tempPath, planned.count, planned.bytes);

            var deleted = await DeleteFilesOlderThanAsync(tempPath, threshold, ct).ConfigureAwait(false);
            AppLogService.Information("Auto optimize done: clear user temp deleted {Count} items, freed {Bytes}B",
                deleted.count, deleted.bytes);

            return new OptimizationStep
            {
                Name = "清空当前用户 Temp（24 小时前）",
                Status = "OK",
                Detail = $"已删除 {deleted.count} 项。",
                BytesFreed = deleted.bytes
            };
        }

        private async Task<OptimizationStep> StepClearWindowsTempAsync(CancellationToken ct)
        {
            var windowsTemp = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
            var threshold = DateTime.Now.AddDays(-7);
            var planned = CountDeletableFiles(windowsTemp, threshold);
            AppLogService.Information("Auto optimize plan: clear windows temp path {Path}, files {Count}, bytes {Bytes}",
                windowsTemp, planned.count, planned.bytes);

            var script = BuildSafeDeleteScriptPrelude() + @"
$target = Join-Path $env:windir 'Temp'
$threshold = (Get-Date).AddDays(-7)
Remove-FilesOlderThanSafely $target $threshold '*'
Remove-EmptyDirectoriesSafely $target";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "清空 Windows Temp（7 天前）",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            AppLogService.Information("Auto optimize done: clear windows temp deleted planned {Count} items, freed approx {Bytes}B",
                planned.count, planned.bytes);
            return new OptimizationStep
            {
                Name = "清空 Windows Temp（7 天前）",
                Status = "OK",
                Detail = $"已执行系统临时目录清理（估算 {planned.count} 项）。",
                BytesFreed = planned.bytes
            };
        }

        private async Task<OptimizationStep> StepClearThumbnailCacheAsync(CancellationToken ct)
        {
            if (!AllowExplorerRestartForThumbnailCleanup)
            {
                return new OptimizationStep
                {
                    Name = "清理缩略图缓存",
                    Status = "Skipped",
                    Detail = "用户未同意重启资源管理器，已跳过。",
                    BytesFreed = 0
                };
            }

            var explorerCacheRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft",
                "Windows",
                "Explorer");
            EnsureSafeCleanupRoot(explorerCacheRoot, "缩略图缓存", explorerCacheRoot);

            var files = SafeDirectoryEnumerateFiles(explorerCacheRoot, "thumbcache_*.db").ToList();
            var bytes = files.Sum(file =>
            {
                try { return new FileInfo(file).Length; } catch { return 0L; }
            });

            AppLogService.Information("Auto optimize plan: clear thumbnail cache files {Count}, bytes {Bytes}",
                files.Count, bytes);

            try
            {
                await RunProcessSimpleAsync("taskkill", "/im explorer.exe /f", ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Stopping explorer before thumbcache cleanup failed.");
            }

            var deleted = 0;
            foreach (var file in files)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    File.Delete(file);
                    deleted++;
                }
                catch (Exception ex)
                {
                    AppLogService.Error(ex, "Deleting thumbnail cache failed for {Path}", file);
                }
            }

            try
            {
                await RunProcessSimpleAsync("explorer.exe", string.Empty, ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Restarting explorer after thumbcache cleanup failed.");
            }

            AppLogService.Information("Auto optimize done: thumbnail cache deleted {Count} items, freed {Bytes}B",
                deleted, bytes);

            return new OptimizationStep
            {
                Name = "清理缩略图缓存",
                Status = "OK",
                Detail = $"已删除 {deleted} 个缩略图缓存文件。",
                BytesFreed = bytes
            };
        }

        private async Task<OptimizationStep> StepFlushDnsAsync(CancellationToken ct)
        {
            await RunProcessSimpleAsync("ipconfig", "/flushdns", ct).ConfigureAwait(false);
            return new OptimizationStep
            {
                Name = "刷新 DNS 缓存",
                Status = "OK",
                Detail = "DNS 解析缓存已刷新。",
                BytesFreed = 0
            };
        }

        private async Task<OptimizationStep> StepResetFontCacheAsync(CancellationToken ct)
        {
            AppLogService.Information("Auto optimize plan: reset font cache service and cache files");
            var script = @"
Stop-Service -Name FontCache -Force -ErrorAction SilentlyContinue
Stop-Service -Name FontCache3.0.0.0 -Force -ErrorAction SilentlyContinue
$fontCachePath = Join-Path $env:windir 'ServiceProfiles\LocalService\AppData\Local\FontCache'
Get-ChildItem -LiteralPath $fontCachePath -Filter '*.dat' -Force -ErrorAction SilentlyContinue |
ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch { } }
Start-Service -Name FontCache -ErrorAction SilentlyContinue
Start-Service -Name FontCache3.0.0.0 -ErrorAction SilentlyContinue";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "重置字体缓存",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            AppLogService.Information("Auto optimize done: font cache reset completed");
            return new OptimizationStep
            {
                Name = "重置字体缓存",
                Status = "OK",
                Detail = "字体缓存服务已重置。",
                BytesFreed = 0
            };
        }

        private async Task<OptimizationStep> StepCleanWindowsUpdateCacheAsync(CancellationToken ct)
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "Download");
            var estimatedBytes = CalculateDirectorySize(root);
            AppLogService.Information("Auto optimize plan: clear windows update cache path {Path}, bytes {Bytes}", root, estimatedBytes);

            var script = BuildSafeDeleteScriptPrelude() + @"
$target = Join-Path $env:windir 'SoftwareDistribution\Download'
$service = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
$wasRunning = $null -ne $service -and $service.Status -eq 'Running'
Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
Remove-DirectoryChildrenSafely $target
if ($wasRunning) { Start-Service -Name wuauserv -ErrorAction SilentlyContinue }";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "清理 Windows 更新缓存",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            AppLogService.Information("Auto optimize done: windows update cache cleaned, freed approx {Bytes}B", estimatedBytes);
            return new OptimizationStep
            {
                Name = "清理 Windows 更新缓存",
                Status = "OK",
                Detail = "Windows Update 缓存目录已清理。",
                BytesFreed = estimatedBytes
            };
        }

        private async Task<OptimizationStep> StepCleanDeliveryOptimizationCacheAsync(CancellationToken ct)
        {
            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "ServiceProfiles",
                "NetworkService",
                "AppData",
                "Local",
                "Microsoft",
                "Windows",
                "DeliveryOptimization",
                "Cache");
            var estimatedBytes = CalculateDirectorySize(root);
            AppLogService.Information("Auto optimize plan: clear delivery optimization cache path {Path}, bytes {Bytes}", root, estimatedBytes);

            var script = BuildSafeDeleteScriptPrelude() + @"
$target = Join-Path $env:windir 'ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache'
$service = Get-Service -Name DoSvc -ErrorAction SilentlyContinue
$wasRunning = $null -ne $service -and $service.Status -eq 'Running'
Stop-Service -Name DoSvc -Force -ErrorAction SilentlyContinue
Remove-DirectoryChildrenSafely $target
if ($wasRunning) { Start-Service -Name DoSvc -ErrorAction SilentlyContinue }";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "清理传递优化缓存",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            return new OptimizationStep
            {
                Name = "清理传递优化缓存",
                Status = "OK",
                Detail = "Delivery Optimization 缓存目录已清理。",
                BytesFreed = estimatedBytes
            };
        }

        private async Task<OptimizationStep> StepCleanWindowsErrorReportsAsync(CancellationToken ct)
        {
            var threshold = DateTime.Now.AddDays(-14);
            var planned = CountWindowsErrorReports(threshold);
            AppLogService.Information("Auto optimize plan: clear WER files {Count}, bytes {Bytes}", planned.count, planned.bytes);

            var script = BuildSafeDeleteScriptPrelude() + @"
$roots = @(
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportArchive'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportQueue'),
    (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportArchive'),
    (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportQueue')
)
$threshold = (Get-Date).AddDays(-14)
foreach ($root in $roots) {
    Remove-FilesOlderThanSafely $root $threshold '*'
    Remove-EmptyDirectoriesSafely $root
}";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "清理 Windows 错误报告（14 天前）",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            return new OptimizationStep
            {
                Name = "清理 Windows 错误报告（14 天前）",
                Status = "OK",
                Detail = $"已清理 14 天前的 Windows 错误报告（估算 {planned.count} 项）。",
                BytesFreed = planned.bytes
            };
        }

        private async Task<OptimizationStep> StepCleanCrashDumpsAsync(CancellationToken ct)
        {
            var threshold = DateTime.Now.AddDays(-30);
            var planned = CountCrashDumps(threshold);
            AppLogService.Information("Auto optimize plan: clear crash dumps {Count}, bytes {Bytes}", planned.count, planned.bytes);

            var script = BuildSafeDeleteScriptPrelude() + @"
$roots = @(
    (Join-Path $env:windir 'Minidump'),
    (Join-Path $env:LOCALAPPDATA 'CrashDumps')
)
$threshold = (Get-Date).AddDays(-30)
foreach ($root in $roots) {
    Remove-FilesOlderThanSafely $root $threshold '*.dmp'
}
$memoryDump = Join-Path $env:windir 'MEMORY.DMP'
if (Test-Path $memoryDump) {
    $info = Get-Item -LiteralPath $memoryDump -Force -ErrorAction SilentlyContinue
    if ($null -ne $info -and $info.LastWriteTime -lt $threshold -and -not (Test-ReparsePoint $info)) {
        try { Remove-Item -LiteralPath $memoryDump -Force -ErrorAction SilentlyContinue } catch { }
    }
}";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "清理崩溃转储（30 天前）",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            return new OptimizationStep
            {
                Name = "清理崩溃转储（30 天前）",
                Status = "OK",
                Detail = $"已清理 30 天前的崩溃转储（估算 {planned.count} 项）。",
                BytesFreed = planned.bytes
            };
        }

        private async Task<OptimizationStep> StepOptimizeSystemDriveAsync(CancellationToken ct)
        {
            var systemDrive = Environment.GetEnvironmentVariable("SystemDrive");
            if (string.IsNullOrWhiteSpace(systemDrive))
            {
                return new OptimizationStep
                {
                    Name = "系统盘 TRIM/碎片整理",
                    Status = "Skipped",
                    Detail = "无法读取系统盘盘符。",
                    BytesFreed = 0
                };
            }

            var driveLetter = systemDrive.Trim().TrimEnd('\\').Replace(":", string.Empty);
            if (string.IsNullOrWhiteSpace(driveLetter))
            {
                return new OptimizationStep
                {
                    Name = "系统盘 TRIM/碎片整理",
                    Status = "Skipped",
                    Detail = "系统盘盘符无效。",
                    BytesFreed = 0
                };
            }

            var mediaType = await QuerySystemDriveMediaTypeAsync(driveLetter, ct).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(mediaType)
                || (!mediaType.Equals("SSD", StringComparison.OrdinalIgnoreCase)
                    && !mediaType.Equals("HDD", StringComparison.OrdinalIgnoreCase)))
            {
                return new OptimizationStep
                {
                    Name = "系统盘 TRIM/碎片整理",
                    Status = "Skipped",
                    Detail = $"无法识别磁盘类型（{mediaType ?? "Unspecified"}），已跳过。",
                    BytesFreed = 0
                };
            }

            var script = mediaType.Equals("SSD", StringComparison.OrdinalIgnoreCase)
                ? $"Optimize-Volume -DriveLetter {driveLetter} -ReTrim -Verbose"
                : $"Optimize-Volume -DriveLetter {driveLetter} -Defrag -Verbose";

            AppLogService.Information("Auto optimize plan: optimize volume drive {Drive}, mode {Mode}",
                driveLetter, mediaType);
            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "系统盘 TRIM/碎片整理",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            return new OptimizationStep
            {
                Name = "系统盘 TRIM/碎片整理",
                Status = "OK",
                Detail = mediaType.Equals("SSD", StringComparison.OrdinalIgnoreCase)
                    ? "系统盘为 SSD，已执行 TRIM。"
                    : "系统盘为 HDD，已执行碎片整理。",
                BytesFreed = 0
            };
        }

        private async Task<OptimizationStep> StepComponentStoreCleanupAsync(CancellationToken ct)
        {
            if (!IsComponentStoreCleanupAvailable())
            {
                return new OptimizationStep
                {
                    Name = "组件存储清理（WinSxS）",
                    Status = "Skipped",
                    Detail = "当前系统不支持或未找到 DISM 组件存储清理。",
                    BytesFreed = 0
                };
            }

            AppLogService.Information("Auto optimize plan: run DISM component store cleanup");
            var script = @"
$dism = Join-Path $env:windir 'System32\dism.exe'
if (-not (Test-Path $dism)) { exit 2 }
& $dism /Online /Cleanup-Image /StartComponentCleanup /NoRestart
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "组件存储清理（WinSxS）",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            return new OptimizationStep
            {
                Name = "组件存储清理（WinSxS）",
                Status = "OK",
                Detail = "已执行 DISM 组件存储清理。",
                BytesFreed = 0
            };
        }

        private static string BuildSafeDeleteScriptPrelude()
        {
            return @"
function Normalize-PathSafe([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return '' }
    try {
        return [System.IO.Path]::GetFullPath($path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    } catch {
        return ''
    }
}

function Test-ReparsePoint($item) {
    if ($null -eq $item) { return $true }
    try {
        return (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint)
    } catch {
        return $true
    }
}

function Test-PathInsideRoot([string]$path, [string]$root) {
    $fullPath = Normalize-PathSafe $path
    $fullRoot = Normalize-PathSafe $root
    if ([string]::IsNullOrWhiteSpace($fullPath) -or [string]::IsNullOrWhiteSpace($fullRoot)) { return $false }
    if ([string]::Equals($fullPath, $fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $fullPath.StartsWith($fullRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-SafeChildItems([string]$root) {
    $safeRoot = Normalize-PathSafe $root
    if ([string]::IsNullOrWhiteSpace($safeRoot) -or -not (Test-Path $safeRoot)) { return @() }
    $pending = New-Object System.Collections.Stack
    $pending.Push($safeRoot)
    $results = New-Object System.Collections.ArrayList
    while ($pending.Count -gt 0) {
        $current = [string]$pending.Pop()
        if (-not (Test-PathInsideRoot $current $safeRoot)) { continue }
        $currentItem = Get-Item -LiteralPath $current -Force -ErrorAction SilentlyContinue
        if (Test-ReparsePoint $currentItem) { continue }
        $children = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction SilentlyContinue)
        foreach ($child in $children) {
            if ($null -eq $child) { continue }
            if (-not (Test-PathInsideRoot $child.FullName $safeRoot)) { continue }
            if (Test-ReparsePoint $child) { continue }
            [void]$results.Add($child)
            if ($child.PSIsContainer) { $pending.Push($child.FullName) }
        }
    }
    return $results
}

function Remove-SafeItem($item, [string]$root) {
    if ($null -eq $item) { return }
    if (-not (Test-PathInsideRoot $item.FullName $root)) { return }
    if (Test-ReparsePoint $item) { return }
    try { Remove-Item -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue } catch { }
}

function Remove-FilesOlderThanSafely([string]$root, [datetime]$threshold, [string]$filter) {
    foreach ($item in (Get-SafeChildItems $root)) {
        if ($item.PSIsContainer) { continue }
        if (-not [string]::IsNullOrWhiteSpace($filter) -and $filter -ne '*' -and $item.Name -notlike $filter) { continue }
        if ($item.LastWriteTime -lt $threshold) { Remove-SafeItem $item $root }
    }
}

function Remove-EmptyDirectoriesSafely([string]$root) {
    $dirs = @(Get-SafeChildItems $root | Where-Object { $_.PSIsContainer } | Sort-Object FullName -Descending)
    foreach ($dir in $dirs) {
        if (-not (Test-PathInsideRoot $dir.FullName $root)) { continue }
        if (Test-ReparsePoint $dir) { continue }
        try {
            if ((Get-ChildItem -LiteralPath $dir.FullName -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
                Remove-Item -LiteralPath $dir.FullName -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}

function Remove-DirectoryChildrenSafely([string]$root) {
    $items = @(Get-SafeChildItems $root | Sort-Object { $_.FullName.Length } -Descending)
    foreach ($item in $items) {
        Remove-SafeItem $item $root
    }
}

";
        }

        private Task<OptimizationStep> StepEmptyRecycleBinAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            AppLogService.Information("Auto optimize plan: empty recycle bin");
            var planned = QueryRecycleBin();
            var result = NativeMethods.SHEmptyRecycleBin(IntPtr.Zero, null,
                NativeMethods.SHERB_NOCONFIRMATION
                | NativeMethods.SHERB_NOPROGRESSUI
                | NativeMethods.SHERB_NOSOUND);

            if (result == 0 || result == NativeMethods.S_FALSE)
            {
                AppLogService.Information("Auto optimize done: recycle bin emptied");
                return Task.FromResult(new OptimizationStep
                {
                    Name = "清空回收站",
                    Status = "OK",
                    Detail = "回收站已清空。",
                    BytesFreed = planned.bytes
                });
            }

            return Task.FromResult(new OptimizationStep
            {
                Name = "清空回收站",
                Status = "Failed",
                Detail = $"回收站清理失败（HRESULT=0x{result:X8}）。",
                BytesFreed = 0
            });
        }

        private async Task<OptimizationStep> StepCleanEventLogsAsync(CancellationToken ct)
        {
            AppLogService.Information("Auto optimize plan: clear event logs if size > 500MB for Application/System/Security");
            var planned = await QueryLargeEventLogsAsync(ct).ConfigureAwait(false);
            var script = @"
$logs = @('Application', 'System', 'Security')
$threshold = 500MB
foreach($name in $logs) {
    $info = Get-WinEvent -ListLog $name -ErrorAction SilentlyContinue
    if ($null -eq $info) { continue }
    if ($info.FileSize -gt $threshold) {
        try { wevtutil cl $name } catch { }
    }
}";

            try
            {
                await ElevatedScriptRunner.RunElevatedScriptAsync(script, true, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return new OptimizationStep
                {
                    Name = "清理事件日志（大于 500MB）",
                    Status = "Failed",
                    Detail = "用户取消了 UAC 授权",
                    BytesFreed = 0
                };
            }

            return new OptimizationStep
            {
                Name = "清理事件日志（大于 500MB）",
                Status = "OK",
                Detail = "已按阈值检查并清理事件日志。",
                BytesFreed = planned.bytes
            };
        }

        private Task<OptimizationStep> StepTrimWorkingSetAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            AppLogService.Information("Auto optimize plan: trim current process working set");
            var process = Process.GetCurrentProcess();
            var ok = NativeMethods.EmptyWorkingSet(process.Handle);
            if (!ok)
            {
                var code = Marshal.GetLastWin32Error();
                return Task.FromResult(new OptimizationStep
                {
                    Name = "压缩当前进程工作集",
                    Status = "Failed",
                    Detail = $"调用 EmptyWorkingSet 失败（Win32={code}）。",
                    BytesFreed = 0
                });
            }

            AppLogService.Information("Auto optimize done: trim current process working set completed");
            return Task.FromResult(new OptimizationStep
            {
                Name = "压缩当前进程工作集",
                Status = "OK",
                Detail = "已执行当前进程内存工作集压缩。",
                BytesFreed = 0
            });
        }

        private static async Task<string> QuerySystemDriveMediaTypeAsync(string driveLetter, CancellationToken ct)
        {
            var script = $@"
$letter = '{driveLetter}'
$part = Get-Partition -DriveLetter $letter -ErrorAction SilentlyContinue
if ($null -eq $part) {{ 'Unspecified'; exit 0 }}
$disk = Get-Disk -Number $part.DiskNumber -ErrorAction SilentlyContinue
if ($null -eq $disk) {{ 'Unspecified'; exit 0 }}
$media = $disk.MediaType.ToString()
if ([string]::IsNullOrWhiteSpace($media)) {{ $media = 'Unspecified' }}
Write-Output $media";

            var output = await ElevatedScriptRunner.RunNonElevatedPowerShellAsync(script, ct).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(output))
            {
                return "Unspecified";
            }

            var line = output
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault()
                ?.Trim();
            return string.IsNullOrWhiteSpace(line) ? "Unspecified" : line;
        }

        private static async Task RunProcessSimpleAsync(string fileName, string arguments, CancellationToken ct)
        {
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments ?? string.Empty,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                if (!process.Start())
                {
                    throw new InvalidOperationException($"无法启动进程：{fileName}");
                }

                await WaitForExitAsync(process, MaxProcessWaitMs, ct).ConfigureAwait(false);
            }
        }

        private static async Task WaitForExitAsync(Process process, int timeoutMs, CancellationToken ct)
        {
            var started = DateTime.UtcNow;
            while (!process.HasExited)
            {
                ct.ThrowIfCancellationRequested();
                if ((DateTime.UtcNow - started).TotalMilliseconds > timeoutMs)
                {
                    throw new TimeoutException("进程执行超时。");
                }

                await Task.Delay(150, ct).ConfigureAwait(false);
            }
        }

        private static (int count, long bytes) CountDeletableFiles(string root, DateTime threshold)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return (0, 0);
            }

            var count = 0;
            long bytes = 0;
            foreach (var file in SafeDirectoryEnumerateFiles(root))
            {
                try
                {
                    var info = new FileInfo(file);
                    if (info.LastWriteTime < threshold)
                    {
                        count++;
                        bytes += info.Length;
                    }
                }
                catch
                {
                }
            }

            return (count, bytes);
        }

        private static (int count, long bytes, int roots) CountWindowsErrorReports(DateTime threshold)
        {
            var count = 0;
            long bytes = 0;
            var roots = 0;
            foreach (var root in GetWindowsErrorReportRoots())
            {
                if (!IsSafeCleanupRoot(root, root))
                {
                    continue;
                }

                roots++;
                var result = CountDeletableFiles(root, threshold);
                count += result.count;
                bytes += result.bytes;
            }

            return (count, bytes, roots);
        }

        private static IEnumerable<string> GetWindowsErrorReportRoots()
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            if (!string.IsNullOrWhiteSpace(local))
            {
                yield return Path.Combine(local, "Microsoft", "Windows", "WER", "ReportArchive");
                yield return Path.Combine(local, "Microsoft", "Windows", "WER", "ReportQueue");
            }

            if (!string.IsNullOrWhiteSpace(programData))
            {
                yield return Path.Combine(programData, "Microsoft", "Windows", "WER", "ReportArchive");
                yield return Path.Combine(programData, "Microsoft", "Windows", "WER", "ReportQueue");
            }
        }

        private static (int count, long bytes, int roots) CountCrashDumps(DateTime threshold)
        {
            var count = 0;
            long bytes = 0;
            var roots = 0;
            foreach (var root in GetCrashDumpRoots())
            {
                if (!IsSafeCleanupRoot(root, root))
                {
                    continue;
                }

                roots++;
                foreach (var file in SafeDirectoryEnumerateFiles(root, "*.dmp"))
                {
                    AddFileIfOlder(file, threshold, ref count, ref bytes);
                }
            }

            var memoryDump = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "MEMORY.DMP");
            if (File.Exists(memoryDump))
            {
                AddFileIfOlder(memoryDump, threshold, ref count, ref bytes);
            }

            return (count, bytes, roots);
        }

        private static IEnumerable<string> GetCrashDumpRoots()
        {
            var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (!string.IsNullOrWhiteSpace(windows))
            {
                yield return Path.Combine(windows, "Minidump");
            }

            if (!string.IsNullOrWhiteSpace(local))
            {
                yield return Path.Combine(local, "CrashDumps");
            }
        }

        private static void AddFileIfOlder(string file, DateTime threshold, ref int count, ref long bytes)
        {
            try
            {
                var info = new FileInfo(file);
                if (info.Exists && info.LastWriteTime < threshold)
                {
                    count++;
                    bytes += info.Length;
                }
            }
            catch
            {
            }
        }

        private static string BuildCountEvidence(string root, int count, long bytes, string scope)
        {
            var path = string.IsNullOrWhiteSpace(root) ? "未知路径" : root;
            return $"{scope}；路径：{path}；文件 {count} 项；估算 {FileSizeFormatter.Format(bytes)}。";
        }

        private static bool IsComponentStoreCleanupAvailable()
        {
            try
            {
                var version = Environment.OSVersion.Version;
                if (version.Major < 6 || (version.Major == 6 && version.Minor < 2))
                {
                    return false;
                }

                var dismPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "dism.exe");
                return File.Exists(dismPath);
            }
            catch
            {
                return false;
            }
        }

        private static void EnsureSafeCleanupRoot(string root, string operationName, params string[] approvedRoots)
        {
            if (!IsSafeCleanupRoot(root, approvedRoots))
            {
                throw new InvalidOperationException($"{operationName} 路径未通过白名单校验，已停止执行。");
            }
        }

        private static bool IsSafeCleanupRoot(string root, params string[] approvedRoots)
        {
            var fullRoot = NormalizeDirectoryPath(root);
            if (string.IsNullOrWhiteSpace(fullRoot) || IsUnsafeCleanupRoot(fullRoot))
            {
                return false;
            }

            if (approvedRoots == null || approvedRoots.Length == 0)
            {
                return false;
            }

            foreach (var approvedRoot in approvedRoots)
            {
                var approved = NormalizeDirectoryPath(approvedRoot);
                if (string.IsNullOrWhiteSpace(approved) || IsUnsafeCleanupRoot(approved))
                {
                    continue;
                }

                if (PathsEqual(fullRoot, approved) || IsSubPathOf(fullRoot, approved))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsUnsafeCleanupRoot(string fullRoot)
        {
            var driveRoot = NormalizeDirectoryPath(Path.GetPathRoot(fullRoot));
            if (!string.IsNullOrWhiteSpace(driveRoot) && PathsEqual(fullRoot, driveRoot))
            {
                return true;
            }

            foreach (var criticalRoot in GetCriticalCleanupRoots())
            {
                if (PathsEqual(fullRoot, criticalRoot))
                {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<string> GetCriticalCleanupRoots()
        {
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var roots = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                userProfile,
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                string.IsNullOrWhiteSpace(userProfile) ? string.Empty : Path.Combine(userProfile, "Downloads")
            };

            return roots
                .Select(NormalizeDirectoryPath)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static string NormalizeDirectoryPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                var full = Path.GetFullPath(path.Trim());
                return full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool PathsEqual(string left, string right)
        {
            return string.Equals(
                NormalizeDirectoryPath(left),
                NormalizeDirectoryPath(right),
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSubPathOf(string child, string parent)
        {
            var normalizedChild = NormalizeDirectoryPath(child);
            var normalizedParent = NormalizeDirectoryPath(parent);
            if (string.IsNullOrWhiteSpace(normalizedChild) || string.IsNullOrWhiteSpace(normalizedParent))
            {
                return false;
            }

            return normalizedChild.StartsWith(
                normalizedParent + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsPathInsideRoot(string path, string root)
        {
            return PathsEqual(path, root) || IsSubPathOf(path, root);
        }

        private static bool IsReparsePointDirectory(string path)
        {
            try
            {
                return (new DirectoryInfo(path).Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint;
            }
            catch
            {
                return true;
            }
        }

        private static bool IsReparsePointFile(string path)
        {
            try
            {
                return (new FileInfo(path).Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint;
            }
            catch
            {
                return true;
            }
        }

        private static (int count, long bytes) CountFiles(string root, string pattern = "*")
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return (0, 0);
            }

            var count = 0;
            long bytes = 0;
            foreach (var file in SafeDirectoryEnumerateFiles(root, pattern))
            {
                try
                {
                    var info = new FileInfo(file);
                    count++;
                    bytes += info.Length;
                }
                catch
                {
                }
            }

            return (count, bytes);
        }

        private static async Task<(int count, long bytes)> DeleteFilesOlderThanAsync(string root, DateTime threshold, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return (0, 0);
            }

            var count = 0;
            long bytes = 0;
            foreach (var file in SafeDirectoryEnumerateFiles(root))
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var info = new FileInfo(file);
                    if (info.LastWriteTime >= threshold)
                    {
                        continue;
                    }

                    bytes += info.Length;
                    info.Attributes = FileAttributes.Normal;
                    File.Delete(file);
                    count++;
                }
                catch
                {
                }

                if (count % 40 == 0)
                {
                    await Task.Yield();
                }
            }

            DeleteEmptyDirectories(root);
            return (count, bytes);
        }

        private static IEnumerable<string> SafeDirectoryEnumerateFiles(string root, string pattern = "*")
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                yield break;
            }

            var safeRoot = NormalizeDirectoryPath(root);
            if (string.IsNullOrWhiteSpace(safeRoot))
            {
                yield break;
            }

            var pending = new Stack<string>();
            pending.Push(safeRoot);

            while (pending.Count > 0)
            {
                var current = pending.Pop();
                if (!PathsEqual(current, safeRoot) && !IsSubPathOf(current, safeRoot))
                {
                    continue;
                }

                if (IsReparsePointDirectory(current))
                {
                    continue;
                }

                string[] subDirs;
                try
                {
                    subDirs = Directory.GetDirectories(current);
                }
                catch
                {
                    subDirs = Array.Empty<string>();
                }

                foreach (var dir in subDirs)
                {
                    if (!IsReparsePointDirectory(dir))
                    {
                        pending.Push(dir);
                    }
                }

                string[] files;
                try
                {
                    files = Directory.GetFiles(current, pattern, SearchOption.TopDirectoryOnly);
                }
                catch
                {
                    files = Array.Empty<string>();
                }

                foreach (var file in files)
                {
                    if (IsPathInsideRoot(file, safeRoot) && !IsReparsePointFile(file))
                    {
                        yield return file;
                    }
                }
            }
        }

        private static void DeleteEmptyDirectories(string root)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return;
            }

            var dirs = SafeDirectoryEnumerateDirectories(root)
                .OrderByDescending(x => x.Length)
                .ToList();

            foreach (var dir in dirs)
            {
                try
                {
                    if ((Directory.EnumerateFileSystemEntries(dir).Any()))
                    {
                        continue;
                    }

                    Directory.Delete(dir, false);
                }
                catch
                {
                }
            }
        }

        private static IEnumerable<string> SafeDirectoryEnumerateDirectories(string root)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                yield break;
            }

            var safeRoot = NormalizeDirectoryPath(root);
            if (string.IsNullOrWhiteSpace(safeRoot))
            {
                yield break;
            }

            var pending = new Stack<string>();
            pending.Push(safeRoot);
            while (pending.Count > 0)
            {
                var current = pending.Pop();
                if (!PathsEqual(current, safeRoot) && !IsSubPathOf(current, safeRoot))
                {
                    continue;
                }

                if (IsReparsePointDirectory(current))
                {
                    continue;
                }

                string[] subDirs;
                try
                {
                    subDirs = Directory.GetDirectories(current);
                }
                catch
                {
                    subDirs = Array.Empty<string>();
                }

                foreach (var dir in subDirs)
                {
                    if (IsPathInsideRoot(dir, safeRoot) && !IsReparsePointDirectory(dir))
                    {
                        yield return dir;
                        pending.Push(dir);
                    }
                }
            }
        }

        private static long CalculateDirectorySize(string root)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return 0;
            }

            long total = 0;
            foreach (var file in SafeDirectoryEnumerateFiles(root))
            {
                try
                {
                    total += new FileInfo(file).Length;
                }
                catch
                {
                }
            }

            return total;
        }

        private static (int count, long bytes) QueryRecycleBin()
        {
            try
            {
                var info = new NativeMethods.SHQUERYRBINFO();
                info.cbSize = Marshal.SizeOf(typeof(NativeMethods.SHQUERYRBINFO));
                var result = NativeMethods.SHQueryRecycleBin(null, ref info);
                if (result == 0)
                {
                    return ((int)Math.Min(info.i64NumItems, int.MaxValue), Math.Max(0, info.i64Size));
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Query recycle bin failed.");
            }

            return (0, 0);
        }

        private static async Task<(int count, long bytes)> QueryLargeEventLogsAsync(CancellationToken ct)
        {
            const long threshold = 500L * 1024L * 1024L;
            const string script = @"
$logs = @('Application', 'System', 'Security')
foreach($name in $logs) {
    $info = Get-WinEvent -ListLog $name -ErrorAction SilentlyContinue
    if ($null -eq $info) { continue }
    if ($info.FileSize -gt 500MB) {
        '{0}|{1}' -f $name, $info.FileSize
    }
}";

            try
            {
                var output = await ElevatedScriptRunner.RunNonElevatedPowerShellAsync(script, ct).ConfigureAwait(false);
                var count = 0;
                long bytes = 0;
                foreach (var line in (output ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = line.Split('|');
                    if (parts.Length != 2)
                    {
                        continue;
                    }

                    long size;
                    if (!long.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out size)
                        || size <= threshold)
                    {
                        continue;
                    }

                    count++;
                    bytes += size;
                }

                return (count, bytes);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Query large event logs failed.");
                return (0, 0);
            }
        }

        private static long GetCurrentWorkingSetBytes()
        {
            try
            {
                using (var process = Process.GetCurrentProcess())
                {
                    return process.WorkingSet64;
                }
            }
            catch
            {
                return 0;
            }
        }

        private sealed class OptimizationStepDefinition
        {
            public OptimizationStepDefinition(string key, string name, Func<CancellationToken, Task<OptimizationStep>> execute)
            {
                Key = key;
                Name = name;
                Execute = execute;
            }

            public string Key { get; }
            public string Name { get; }
            public Func<CancellationToken, Task<OptimizationStep>> Execute { get; }
        }

        private static class NativeMethods
        {
            public const uint SHERB_NOCONFIRMATION = 0x00000001;
            public const uint SHERB_NOPROGRESSUI = 0x00000002;
            public const uint SHERB_NOSOUND = 0x00000004;
            public const int S_FALSE = 1;

            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
            public static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

            [StructLayout(LayoutKind.Sequential, Pack = 4)]
            public struct SHQUERYRBINFO
            {
                public int cbSize;
                public long i64Size;
                public long i64NumItems;
            }

            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
            public static extern int SHQueryRecycleBin(string pszRootPath, ref SHQUERYRBINFO pSHQueryRBInfo);

            [DllImport("psapi.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool EmptyWorkingSet(IntPtr hProcess);
        }
    }

    public sealed class OptimizationPlanItem : INotifyPropertyChanged
    {
        private bool _isSelected;

        public event PropertyChangedEventHandler PropertyChanged;

        public string Key { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public bool RequiresAdmin { get; set; }
        public long EstimatedBytes { get; set; }
        public bool CanOptimize { get; set; }
        public string Detail { get; set; }
        public string RiskLevel { get; set; }
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
                OnPropertyChanged();
                OnPropertyChanged(nameof(StatusDisplay));
            }
        }

        public string StatusDisplay
        {
            get
            {
                if (!CanOptimize)
                {
                    return "无需处理";
                }

                return IsSelected ? "已选择" : "可选择";
            }
        }

        public string RequiresAdminDisplay => RequiresAdmin ? "需要" : "否";
        public string EstimatedBytesDisplay => FileSizeFormatter.Format(EstimatedBytes);

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal static class ElevatedScriptRunner
    {
        private const int ElevatedScriptTimeoutMs = 900000;

        public static async Task RunElevatedScriptAsync(string script, bool waitForExit, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(script))
            {
                throw new ArgumentException("script is empty", nameof(script));
            }

            var tempPath = Path.Combine(Path.GetTempPath(), $"mytools_{Guid.NewGuid():N}.ps1");
            using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                await writer.WriteAsync(script).ConfigureAwait(false);
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempPath}\"",
                    UseShellExecute = true,
                    Verb = "runas"
                };

                var process = Process.Start(psi);
                if (process == null || !waitForExit)
                {
                    return;
                }

                await WaitForExitAsync(process, ct, ElevatedScriptTimeoutMs).ConfigureAwait(false);
                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException($"提权脚本执行失败（ExitCode={process.ExitCode}）。");
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                throw new OperationCanceledException("用户取消了 UAC 授权");
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }

        public static async Task<string> RunNonElevatedPowerShellAsync(string script, CancellationToken ct)
        {
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{EscapeForPowerShellCommand(script)}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                if (!process.Start())
                {
                    throw new InvalidOperationException("无法启动 powershell 进程。");
                }

                var outputTask = process.StandardOutput.ReadToEndAsync();
                var errorTask = process.StandardError.ReadToEndAsync();
                await WaitForExitAsync(process, ct).ConfigureAwait(false);
                var output = await outputTask.ConfigureAwait(false);
                var error = await errorTask.ConfigureAwait(false);
                if (!string.IsNullOrWhiteSpace(error))
                {
                    AppLogService.Information("PowerShell stderr: {Message}", error.Trim());
                }

                return output ?? string.Empty;
            }
        }

        private static string EscapeForPowerShellCommand(string script)
        {
            return (script ?? string.Empty)
                .Replace("`", "``")
                .Replace("\"", "`\"");
        }

        private static async Task WaitForExitAsync(Process process, CancellationToken ct, int timeoutMs = 120000)
        {
            var started = DateTime.UtcNow;
            while (!process.HasExited)
            {
                ct.ThrowIfCancellationRequested();
                if ((DateTime.UtcNow - started).TotalMilliseconds > timeoutMs)
                {
                    try { process.Kill(); } catch { }
                    throw new TimeoutException("PowerShell 脚本执行超时。");
                }

                await Task.Delay(150, ct).ConfigureAwait(false);
            }
        }
    }
}
