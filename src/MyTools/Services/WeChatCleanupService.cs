using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;

namespace MyTools.Services
{
    public sealed class WeChatCleanupService
    {
        private static readonly Regex ForbiddenDbFileRegex =
            new Regex(@"^(Msg.*\.db|MicroMsg\.db)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly HashSet<string> ImageExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"
        };
        private static readonly HashSet<string> VideoExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv", ".m4v"
        };
        private static readonly HashSet<string> AudioExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".amr", ".silk", ".aac", ".mp3", ".wav", ".m4a", ".ogg"
        };

        public async Task<WeChatCleanupScanResult> ScanAsync(
            IEnumerable<WeChatRoot> roots,
            WeChatCleanupScanOptions options,
            IProgress<string> progress,
            CancellationToken ct)
        {
            var rootList = (roots ?? Enumerable.Empty<WeChatRoot>()).Where(x => x != null).ToList();
            var result = new WeChatCleanupScanResult();
            if (rootList.Count == 0)
            {
                return result;
            }

            var normalizedOptions = options ?? WeChatCleanupScanOptions.CreateDefault();
            var timeStart = normalizedOptions.StartDate.Date;
            var timeEnd = normalizedOptions.EndDate.Date.AddDays(1).AddMilliseconds(-1);
            var categories = normalizedOptions.Categories.Count == 0
                ? new HashSet<WeChatDataCategory> { WeChatDataCategory.Image, WeChatDataCategory.Video, WeChatDataCategory.Voice, WeChatDataCategory.File, WeChatDataCategory.Cache }
                : normalizedOptions.Categories;

            var totalRoots = rootList.Count;
            for (var rootIndex = 0; rootIndex < totalRoots; rootIndex++)
            {
                ct.ThrowIfCancellationRequested();
                var root = rootList[rootIndex];
                progress?.Report($"扫描微信目录 [{rootIndex + 1}/{totalRoots}] {root.WxIdOrUserName} ...");
                if (!Directory.Exists(root.RootPath))
                {
                    continue;
                }

                if (root.Variant == WeChatVariant.XWechat)
                {
                    ScanXWechatRoot(root, categories, timeStart, timeEnd, result, ct);
                    if (categories.Contains(WeChatDataCategory.Text))
                    {
                        var note = "[待确认] 新版 xwechat_files 文字类数据尚未识别。";
                        if (!result.PendingNotes.Contains(note))
                        {
                            result.PendingNotes.Add(note);
                        }
                    }

                    continue;
                }

                foreach (var category in categories)
                {
                    ct.ThrowIfCancellationRequested();
                    foreach (var path in ResolveCategoryPaths(root, category))
                    {
                        if (!Directory.Exists(path))
                        {
                            continue;
                        }

                        foreach (var file in SafePathHelper.EnumerateFilesSafe(path))
                        {
                            ct.ThrowIfCancellationRequested();
                            FileInfo info;
                            try
                            {
                                info = new FileInfo(file);
                            }
                            catch
                            {
                                continue;
                            }

                            if (IsForbiddenWeChatPath(info.FullName))
                            {
                                continue;
                            }

                            if (info.LastWriteTime < timeStart || info.LastWriteTime > timeEnd)
                            {
                                continue;
                            }

                            result.Candidates.Add(new WeChatCleanupCandidate
                            {
                                Path = info.FullName,
                                Category = category,
                                LastWriteTime = info.LastWriteTime,
                                Bytes = info.Length,
                                WxIdOrUserName = root.WxIdOrUserName,
                                Variant = root.Variant,
                                IsSelected = true
                            });
                        }
                    }

                    if (root.Variant == WeChatVariant.XWechat && category == WeChatDataCategory.Text)
                    {
                        var note = "[待确认] 文字类别在新版 xwechat_files 尚未识别";
                        if (!result.PendingNotes.Contains(note))
                        {
                            result.PendingNotes.Add(note);
                        }
                    }
                }
            }

            await Task.Yield();
            return result;
        }

        public async Task<WeChatCleanupResult> CleanupAsync(
            IEnumerable<WeChatCleanupCandidate> selectedCandidates,
            IEnumerable<WeChatRoot> roots,
            IProgress<string> progress,
            CancellationToken ct)
        {
            var candidates = (selectedCandidates ?? Enumerable.Empty<WeChatCleanupCandidate>())
                .Where(x => x != null && x.IsSelected)
                .ToList();

            var rootPaths = (roots ?? Enumerable.Empty<WeChatRoot>())
                .Select(x => x?.RootPath)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            var result = new WeChatCleanupResult();
            if (candidates.Count == 0 || rootPaths.Count == 0)
            {
                return result;
            }

            var groups = candidates
                .GroupBy(x => x.Category)
                .OrderBy(x => x.Key.ToString())
                .ToList();

            foreach (var group in groups)
            {
                ct.ThrowIfCancellationRequested();
                var step = new OptimizationStep
                {
                    Name = $"微信清理 - {group.Key}",
                    Status = "OK",
                    Detail = string.Empty,
                    BytesFreed = 0
                };

                var deletedCount = 0;
                var skippedCount = 0;
                var failedCount = 0;

                foreach (var candidate in group)
                {
                    ct.ThrowIfCancellationRequested();
                    progress?.Report($"清理中：{candidate.Category} - {Path.GetFileName(candidate.Path)}");
                    try
                    {
                        var fullPath = Path.GetFullPath(candidate.Path);
                        if (!SafePathHelper.IsPathInsideAny(fullPath, rootPaths))
                        {
                            throw new InvalidOperationException("路径不在微信根目录内。");
                        }

                        if (IsForbiddenWeChatPath(fullPath))
                        {
                            skippedCount++;
                            continue;
                        }

                        if (!File.Exists(fullPath))
                        {
                            skippedCount++;
                            continue;
                        }

                        AppLogService.Information("WeChat cleanup plan: recycle file {Path}", fullPath);
                        var bytes = new FileInfo(fullPath).Length;
                        File.SetAttributes(fullPath, FileAttributes.Normal);
                        // 发送到回收站而非永久删除，给用户最后一道保险
                        FileSystem.DeleteFile(fullPath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin, UICancelOption.ThrowException);
                        AppLogService.Information("WeChat cleanup done: recycled file {Path}, freed {Bytes}B", fullPath, bytes);
                        step.BytesFreed += bytes;
                        deletedCount++;
                    }
                    catch (IOException)
                    {
                        skippedCount++;
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        AppLogService.Error(ex, "WeChat cleanup failed for {Path}", candidate.Path ?? string.Empty);
                    }
                }

                if (failedCount > 0)
                {
                    step.Status = "Failed";
                }
                else if (deletedCount == 0 && skippedCount > 0)
                {
                    step.Status = "Skipped";
                }

                step.Detail = $"成功 {deletedCount}，跳过 {skippedCount}，失败 {failedCount}";
                result.Steps.Add(step);
            }

            result.DeletedCount = result.Steps.Sum(x => ParseStepCount(x.Detail, "成功"));
            result.FreedBytes = result.Steps.Sum(x => x.BytesFreed);
            await Task.Yield();
            return result;
        }

        public static bool IsForbiddenWeChatPath(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath))
            {
                return true;
            }

            var normalized = SafeNormalize(fullPath);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return true;
            }

            var fileName = Path.GetFileName(normalized);
            if (ForbiddenDbFileRegex.IsMatch(fileName))
            {
                return true;
            }

            if (fileName.Equals("Login.dat", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (fileName.StartsWith("SafeStorage", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var segments = normalized.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var segment in segments)
            {
                if (segment.Equals("MMKV", StringComparison.OrdinalIgnoreCase)
                    || segment.Equals("config", StringComparison.OrdinalIgnoreCase)
                    || segment.Equals("db_storage", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static HashSet<WeChatDataCategory> BuildCategories(
            bool includeText,
            bool includeImage,
            bool includeVideo,
            bool includeVoice,
            bool includeFile,
            bool includeCache)
        {
            var categories = new HashSet<WeChatDataCategory>();
            if (includeText) categories.Add(WeChatDataCategory.Text);
            if (includeImage) categories.Add(WeChatDataCategory.Image);
            if (includeVideo) categories.Add(WeChatDataCategory.Video);
            if (includeVoice) categories.Add(WeChatDataCategory.Voice);
            if (includeFile) categories.Add(WeChatDataCategory.File);
            if (includeCache) categories.Add(WeChatDataCategory.Cache);
            return categories;
        }

        private static int ParseStepCount(string detail, string key)
        {
            if (string.IsNullOrWhiteSpace(detail) || string.IsNullOrWhiteSpace(key))
            {
                return 0;
            }

            var start = detail.IndexOf(key, StringComparison.OrdinalIgnoreCase);
            if (start < 0)
            {
                return 0;
            }

            start += key.Length;
            var end = detail.IndexOf('，', start);
            if (end < 0)
            {
                end = detail.Length;
            }

            var segment = detail.Substring(start, end - start).Trim();
            if (int.TryParse(segment, out var value))
            {
                return value;
            }

            return 0;
        }

        private static IEnumerable<string> ResolveCategoryPaths(WeChatRoot root, WeChatDataCategory category)
        {
            if (root == null || string.IsNullOrWhiteSpace(root.RootPath))
            {
                return Enumerable.Empty<string>();
            }

            if (root.Variant == WeChatVariant.LegacyWeChat)
            {
                switch (category)
                {
                    case WeChatDataCategory.Image:
                        return new[] { Path.Combine(root.RootPath, "FileStorage", "Image") };
                    case WeChatDataCategory.Video:
                        return new[] { Path.Combine(root.RootPath, "FileStorage", "Video") };
                    case WeChatDataCategory.Voice:
                        return new[] { Path.Combine(root.RootPath, "FileStorage", "Voice2") };
                    case WeChatDataCategory.File:
                        return new[] { Path.Combine(root.RootPath, "FileStorage", "File") };
                    case WeChatDataCategory.Cache:
                        return new[] { Path.Combine(root.RootPath, "FileStorage", "Cache") };
                    case WeChatDataCategory.Text:
                        return new[]
                        {
                            Path.Combine(root.RootPath, "FileStorage", "CustomEmotion"),
                            Path.Combine(root.RootPath, "FileStorage", "MsgTemp")
                        };
                    default:
                        return Enumerable.Empty<string>();
                }
            }

            switch (category)
            {
                case WeChatDataCategory.Image:
                    return new[] { Path.Combine(root.RootPath, "msg", "attach") };
                case WeChatDataCategory.Video:
                    return new[] { Path.Combine(root.RootPath, "msg", "video") };
                case WeChatDataCategory.Voice:
                    return new[] { Path.Combine(root.RootPath, "msg", "audio") };
                case WeChatDataCategory.File:
                    return new[] { Path.Combine(root.RootPath, "msg", "file") };
                case WeChatDataCategory.Cache:
                    return new[] { Path.Combine(root.RootPath, "temp") };
                case WeChatDataCategory.Text:
                    return Enumerable.Empty<string>();
                default:
                    return Enumerable.Empty<string>();
            }
        }

        private static void ScanXWechatRoot(
            WeChatRoot root,
            HashSet<WeChatDataCategory> categories,
            DateTime timeStart,
            DateTime timeEnd,
            WeChatCleanupScanResult result,
            CancellationToken ct)
        {
            foreach (var path in ResolveXWechatScanPaths(root))
            {
                ct.ThrowIfCancellationRequested();
                if (!Directory.Exists(path))
                {
                    continue;
                }

                foreach (var file in SafePathHelper.EnumerateFilesSafe(path))
                {
                    ct.ThrowIfCancellationRequested();
                    FileInfo info;
                    try
                    {
                        info = new FileInfo(file);
                    }
                    catch
                    {
                        continue;
                    }

                    if (IsForbiddenWeChatPath(info.FullName))
                    {
                        continue;
                    }

                    if (info.LastWriteTime < timeStart || info.LastWriteTime > timeEnd)
                    {
                        continue;
                    }

                    var category = ResolveXWechatCategory(root, info.FullName);
                    if (!category.HasValue || !categories.Contains(category.Value))
                    {
                        continue;
                    }

                    result.Candidates.Add(new WeChatCleanupCandidate
                    {
                        Path = info.FullName,
                        Category = category.Value,
                        LastWriteTime = info.LastWriteTime,
                        Bytes = info.Length,
                        WxIdOrUserName = root.WxIdOrUserName,
                        Variant = root.Variant,
                        IsSelected = true
                    });
                }
            }
        }

        private static IEnumerable<string> ResolveXWechatScanPaths(WeChatRoot root)
        {
            return new[]
            {
                Path.Combine(root.RootPath, "msg", "attach"),
                Path.Combine(root.RootPath, "msg", "video"),
                Path.Combine(root.RootPath, "msg", "file"),
                Path.Combine(root.RootPath, "cache"),
                Path.Combine(root.RootPath, "temp"),
                Path.Combine(root.RootPath, "resource"),
                Path.Combine(root.RootPath, "business", "emoticon", "Thumb"),
                Path.Combine(root.RootPath, "business", "favorite", "temp")
            };
        }

        private static WeChatDataCategory? ResolveXWechatCategory(WeChatRoot root, string fullPath)
        {
            var extension = Path.GetExtension(fullPath);

            if (IsInside(fullPath, Path.Combine(root.RootPath, "cache"))
                || IsInside(fullPath, Path.Combine(root.RootPath, "temp"))
                || IsInside(fullPath, Path.Combine(root.RootPath, "resource"))
                || IsInside(fullPath, Path.Combine(root.RootPath, "business", "emoticon", "Thumb"))
                || IsInside(fullPath, Path.Combine(root.RootPath, "business", "favorite", "temp")))
            {
                return WeChatDataCategory.Cache;
            }

            if (ContainsPathSegment(fullPath, "Img") || ImageExtensions.Contains(extension))
            {
                return WeChatDataCategory.Image;
            }

            if (ContainsPathSegment(fullPath, "Video")
                || IsInside(fullPath, Path.Combine(root.RootPath, "msg", "video"))
                || VideoExtensions.Contains(extension))
            {
                return WeChatDataCategory.Video;
            }

            if (ContainsPathSegment(fullPath, "Audio")
                || ContainsPathSegment(fullPath, "Voice")
                || AudioExtensions.Contains(extension))
            {
                return WeChatDataCategory.Voice;
            }

            if (IsInside(fullPath, Path.Combine(root.RootPath, "msg", "attach"))
                || IsInside(fullPath, Path.Combine(root.RootPath, "msg", "file")))
            {
                return WeChatDataCategory.File;
            }

            return null;
        }

        private static bool ContainsPathSegment(string fullPath, string segmentName)
        {
            var segments = (fullPath ?? string.Empty).Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            return segments.Any(x => x.Equals(segmentName, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsInside(string fullPath, string directory)
        {
            if (string.IsNullOrWhiteSpace(fullPath) || string.IsNullOrWhiteSpace(directory))
            {
                return false;
            }

            try
            {
                var normalizedFile = Path.GetFullPath(fullPath).TrimEnd('\\');
                var normalizedDirectory = Path.GetFullPath(directory).TrimEnd('\\') + "\\";
                return normalizedFile.StartsWith(normalizedDirectory, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string SafeNormalize(string path)
        {
            try
            {
                return Path.GetFullPath(path);
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    public sealed class WeChatCleanupScanOptions
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public HashSet<WeChatDataCategory> Categories { get; set; } = new HashSet<WeChatDataCategory>();

        public static WeChatCleanupScanOptions CreateDefault()
        {
            return new WeChatCleanupScanOptions
            {
                StartDate = DateTime.Today.AddDays(-30),
                EndDate = DateTime.Today,
                Categories = new HashSet<WeChatDataCategory>
                {
                    WeChatDataCategory.Image,
                    WeChatDataCategory.Video,
                    WeChatDataCategory.Voice,
                    WeChatDataCategory.File,
                    WeChatDataCategory.Cache
                }
            };
        }
    }

    public sealed class WeChatCleanupScanResult
    {
        public List<WeChatCleanupCandidate> Candidates { get; } = new List<WeChatCleanupCandidate>();
        public List<string> PendingNotes { get; } = new List<string>();
    }

    public sealed class WeChatCleanupResult
    {
        public List<OptimizationStep> Steps { get; } = new List<OptimizationStep>();
        public int DeletedCount { get; set; }
        public long FreedBytes { get; set; }
    }

    public sealed class WeChatCleanupCandidate : System.ComponentModel.INotifyPropertyChanged
    {
        private bool _isSelected;

        public string Path { get; set; }
        public WeChatDataCategory Category { get; set; }
        public DateTime LastWriteTime { get; set; }
        public long Bytes { get; set; }
        public string WxIdOrUserName { get; set; }
        public WeChatVariant Variant { get; set; }

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
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public string BytesDisplay => FileSizeFormatter.Format(Bytes);
        public string CategoryDisplay => GetCategoryDisplay(Category);
        public string RiskDisplay => GetRiskDisplay(Category);
        public string AdviceDisplay => GetAdviceDisplay(Category);

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        private static string GetCategoryDisplay(WeChatDataCategory category)
        {
            switch (category)
            {
                case WeChatDataCategory.Text: return "文字";
                case WeChatDataCategory.Image: return "图像";
                case WeChatDataCategory.Video: return "视频";
                case WeChatDataCategory.Voice: return "语音";
                case WeChatDataCategory.File: return "文件";
                case WeChatDataCategory.Cache: return "缓存";
                default: return category.ToString();
            }
        }

        private static string GetRiskDisplay(WeChatDataCategory category)
        {
            switch (category)
            {
                case WeChatDataCategory.Cache:
                    return "低";
                case WeChatDataCategory.Text:
                    return "高";
                default:
                    return "中";
            }
        }

        private static string GetAdviceDisplay(WeChatDataCategory category)
        {
            switch (category)
            {
                case WeChatDataCategory.Cache:
                    return "可优先清理";
                case WeChatDataCategory.Text:
                    return "先备份再清理";
                case WeChatDataCategory.Image:
                case WeChatDataCategory.Video:
                case WeChatDataCategory.Voice:
                case WeChatDataCategory.File:
                    return "按日期确认，建议先备份";
                default:
                    return "确认后清理";
            }
        }
    }

    public enum WeChatDataCategory
    {
        Text,
        Image,
        Video,
        Voice,
        File,
        Cache
    }
}
