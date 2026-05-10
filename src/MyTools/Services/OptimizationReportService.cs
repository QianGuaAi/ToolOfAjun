using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public sealed class OptimizationReportService
    {
        private static readonly SemaphoreSlim FileLock = new SemaphoreSlim(1, 1);

        private static readonly string ReportRoot =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyTools", "OptimizationReports");

        private static readonly string IndexPath = Path.Combine(ReportRoot, "index.json");

        public async Task<IReadOnlyList<OptimizationReportItem>> LoadAllAsync()
        {
            await FileLock.WaitAsync().ConfigureAwait(false);
            try
            {
                var index = await LoadIndexCoreAsync().ConfigureAwait(false);
                var reports = new List<OptimizationReportItem>();

                foreach (var item in index.OrderByDescending(x => x.StartedAt))
                {
                    try
                    {
                        var reportPath = ResolveReportPath(item);
                        if (!File.Exists(reportPath))
                        {
                            reports.Add(BuildFallbackReport(item));
                            continue;
                        }

                        using (var stream = new FileStream(reportPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                        using (var reader = new StreamReader(stream, Encoding.UTF8))
                        {
                            var json = await reader.ReadToEndAsync().ConfigureAwait(false);
                            var report = JsonConvert.DeserializeObject<OptimizationReportItem>(json);
                            if (report == null)
                            {
                                continue;
                            }

                            NormalizeReport(report, item.FileName);
                            reports.Add(report);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLogService.Error(ex, "Loading optimization report file failed for {ReportId}", item.Id ?? string.Empty);
                    }
                }

                return reports;
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading optimization reports failed.");
                return Array.Empty<OptimizationReportItem>();
            }
            finally
            {
                FileLock.Release();
            }
        }

        public async Task SaveAsync(OptimizationReportItem report)
        {
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            await FileLock.WaitAsync().ConfigureAwait(false);
            try
            {
                Directory.CreateDirectory(ReportRoot);
                NormalizeReport(report, report.FileName);

                if (string.IsNullOrWhiteSpace(report.Id))
                {
                    report.Id = Guid.NewGuid().ToString("N").Substring(0, 8);
                }

                if (report.StartedAt == default(DateTime))
                {
                    report.StartedAt = DateTime.Now;
                }

                if (report.FinishedAt == default(DateTime))
                {
                    report.FinishedAt = DateTime.Now;
                }

                if (string.IsNullOrWhiteSpace(report.FileName))
                {
                    report.FileName = $"report-{report.StartedAt:yyyyMMdd-HHmmss}-{report.Id}.json";
                }

                var reportPath = ResolveReportPath(report.FileName);
                var json = JsonConvert.SerializeObject(report, Formatting.Indented);
                using (var stream = new FileStream(reportPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(true)))
                {
                    await writer.WriteAsync(json).ConfigureAwait(false);
                }

                var index = await LoadIndexCoreAsync().ConfigureAwait(false);
                var existing = index.FirstOrDefault(x => string.Equals(x.Id, report.Id, StringComparison.OrdinalIgnoreCase));
                var summary = new OptimizationReportIndexItem
                {
                    Id = report.Id,
                    StartedAt = report.StartedAt,
                    FinishedAt = report.FinishedAt,
                    ReportType = report.ReportType,
                    TotalBytesFreed = report.TotalBytesFreed,
                    Summary = report.Summary ?? string.Empty,
                    FileName = report.FileName
                };

                if (existing == null)
                {
                    index.Add(summary);
                }
                else
                {
                    existing.StartedAt = summary.StartedAt;
                    existing.FinishedAt = summary.FinishedAt;
                    existing.ReportType = summary.ReportType;
                    existing.TotalBytesFreed = summary.TotalBytesFreed;
                    existing.Summary = summary.Summary;
                    existing.FileName = summary.FileName;
                }

                index = index
                    .OrderByDescending(x => x.StartedAt)
                    .ToList();

                await SaveIndexCoreAsync(index).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Saving optimization report failed for {ReportId}", report.Id ?? string.Empty);
                throw;
            }
            finally
            {
                FileLock.Release();
            }
        }

        public async Task DeleteAsync(IEnumerable<string> ids)
        {
            var deletingIds = new HashSet<string>((ids ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim()), StringComparer.OrdinalIgnoreCase);

            if (deletingIds.Count == 0)
            {
                return;
            }

            await FileLock.WaitAsync().ConfigureAwait(false);
            try
            {
                var index = await LoadIndexCoreAsync().ConfigureAwait(false);
                var removed = index.Where(x => deletingIds.Contains(x.Id)).ToList();

                foreach (var entry in removed)
                {
                    var reportPath = ResolveReportPath(entry);
                    try
                    {
                        if (File.Exists(reportPath))
                        {
                            File.Delete(reportPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLogService.Error(ex, "Deleting optimization report file failed for {ReportId}", entry.Id ?? string.Empty);
                    }
                }

                index = index
                    .Where(x => !deletingIds.Contains(x.Id))
                    .OrderByDescending(x => x.StartedAt)
                    .ToList();

                await SaveIndexCoreAsync(index).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Deleting optimization reports failed.");
                throw;
            }
            finally
            {
                FileLock.Release();
            }
        }

        private static async Task<List<OptimizationReportIndexItem>> LoadIndexCoreAsync()
        {
            try
            {
                if (!File.Exists(IndexPath))
                {
                    return new List<OptimizationReportIndexItem>();
                }

                using (var stream = new FileStream(IndexPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    var json = await reader.ReadToEndAsync().ConfigureAwait(false);
                    return JsonConvert.DeserializeObject<List<OptimizationReportIndexItem>>(json) ?? new List<OptimizationReportIndexItem>();
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading optimization report index failed.");
                return new List<OptimizationReportIndexItem>();
            }
        }

        private static async Task SaveIndexCoreAsync(List<OptimizationReportIndexItem> index)
        {
            Directory.CreateDirectory(ReportRoot);
            var json = JsonConvert.SerializeObject(index ?? new List<OptimizationReportIndexItem>(), Formatting.Indented);
            using (var stream = new FileStream(IndexPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(true)))
            {
                await writer.WriteAsync(json).ConfigureAwait(false);
            }
        }

        private static OptimizationReportItem BuildFallbackReport(OptimizationReportIndexItem index)
        {
            return new OptimizationReportItem
            {
                Id = index.Id,
                StartedAt = index.StartedAt,
                FinishedAt = index.FinishedAt,
                ReportType = index.ReportType,
                TotalBytesFreed = index.TotalBytesFreed,
                Summary = index.Summary,
                Steps = new List<OptimizationStep>
                {
                    new OptimizationStep
                    {
                        Name = "报告文件缺失",
                        Status = "Skipped",
                        Detail = "仅保留索引信息，详细步骤文件不存在。",
                        Duration = TimeSpan.Zero,
                        BytesFreed = 0
                    }
                },
                FileName = index.FileName
            };
        }

        private static void NormalizeReport(OptimizationReportItem report, string fileName)
        {
            report.ReportType = report.ReportType ?? "AutoOptimize";
            report.Summary = report.Summary ?? string.Empty;
            report.Steps = report.Steps ?? new List<OptimizationStep>();
            report.FileName = fileName ?? report.FileName;
        }

        private static string ResolveReportPath(OptimizationReportIndexItem index)
        {
            return ResolveReportPath(index?.FileName, index?.Id, index?.StartedAt);
        }

        private static string ResolveReportPath(string fileName, string id = null, DateTime? startedAt = null)
        {
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                return Path.Combine(ReportRoot, fileName);
            }

            var dt = startedAt ?? DateTime.Now;
            var safeId = string.IsNullOrWhiteSpace(id) ? Guid.NewGuid().ToString("N").Substring(0, 8) : id;
            return Path.Combine(ReportRoot, $"report-{dt:yyyyMMdd-HHmmss}-{safeId}.json");
        }
    }

    public sealed class OptimizationReportItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private long _totalBytesFreed;

        public string Id { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime FinishedAt { get; set; }
        public string ReportType { get; set; }
        public List<OptimizationStep> Steps { get; set; } = new List<OptimizationStep>();

        public long TotalBytesFreed
        {
            get => _totalBytesFreed;
            set
            {
                if (_totalBytesFreed == value)
                {
                    return;
                }

                _totalBytesFreed = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TotalBytesFreedDisplay));
            }
        }

        public string Summary { get; set; }
        public string FileName { get; set; }

        [JsonIgnore]
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
            }
        }

        [JsonIgnore]
        public string TotalBytesFreedDisplay => FileSizeFormatter.Format(TotalBytesFreed);

        [JsonIgnore]
        public string ReportTypeDisplay
        {
            get
            {
                if (string.Equals(ReportType, "JunkCleanup", StringComparison.OrdinalIgnoreCase))
                {
                    return "垃圾清理";
                }

                if (string.Equals(ReportType, "WeChatCleanup", StringComparison.OrdinalIgnoreCase))
                {
                    return "微信清理";
                }

                return "自动优化";
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public sealed class OptimizationStep
    {
        public string Name { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }
        public long BytesFreed { get; set; }
        public TimeSpan Duration { get; set; }
    }

    public sealed class OptimizationReportIndexItem
    {
        public string Id { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime FinishedAt { get; set; }
        public string ReportType { get; set; }
        public long TotalBytesFreed { get; set; }
        public string Summary { get; set; }
        public string FileName { get; set; }
    }

    internal static class FileSizeFormatter
    {
        public static string Format(long bytes)
        {
            if (bytes <= 0)
            {
                return "0 B";
            }

            const double kb = 1024d;
            const double mb = kb * 1024d;
            const double gb = mb * 1024d;
            var value = (double)bytes;

            if (value >= gb)
            {
                return (value / gb).ToString("F2") + " GB";
            }

            if (value >= mb)
            {
                return (value / mb).ToString("F2") + " MB";
            }

            if (value >= kb)
            {
                return (value / kb).ToString("F2") + " KB";
            }

            return bytes + " B";
        }
    }
}
