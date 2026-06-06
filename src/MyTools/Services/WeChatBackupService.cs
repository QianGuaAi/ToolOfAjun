using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public sealed class WeChatBackupService
    {
        private readonly WeChatCleanupService _cleanupService = new WeChatCleanupService();

        public async Task<WeChatBackupResult> BackupAsync(
            WeChatBackupOptions options,
            IProgress<WeChatBackupProgress> progress,
            CancellationToken ct)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Root == null || string.IsNullOrWhiteSpace(options.Root.RootPath) || !Directory.Exists(options.Root.RootPath))
            {
                throw new InvalidOperationException("未检测到可用的微信根目录。");
            }

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new InvalidOperationException("请选择备份输出目录。");
            }

            if (options.Categories == null || options.Categories.Count == 0)
            {
                throw new InvalidOperationException("请至少选择一种备份类别。");
            }

            Directory.CreateDirectory(options.OutputDirectory);

            var scanResult = await _cleanupService.ScanAsync(
                new[] { options.Root },
                new WeChatCleanupScanOptions
                {
                    StartDate = options.StartDate,
                    EndDate = options.EndDate,
                    Categories = new HashSet<WeChatDataCategory>(options.Categories ?? new HashSet<WeChatDataCategory>())
                },
                null,
                ct).ConfigureAwait(false);

            var candidates = scanResult.Candidates
                .Where(x => x != null && !WeChatCleanupService.IsForbiddenWeChatPath(x.Path))
                .ToList();

            var estimatedSize = candidates.Sum(x => x.Bytes);
            EnsureEnoughDiskSpace(options.OutputDirectory, estimatedSize);

            var zipPath = string.IsNullOrWhiteSpace(options.OutputZipPath)
                ? BuildDefaultZipPath(options.OutputDirectory, options.StartDate, options.EndDate)
                : options.OutputZipPath;

            var manifest = new WeChatBackupManifest
            {
                SchemaVersion = 1,
                CreatedAt = DateTimeOffset.Now,
                Machine = Environment.MachineName,
                WechatVariant = options.Root.Variant.ToString(),
                WxId = options.Root.WxIdOrUserName,
                WechatRoot = options.Root.RootPath,
                DateRange = new WeChatBackupDateRange
                {
                    Start = options.StartDate.ToString("yyyy-MM-dd"),
                    End = options.EndDate.ToString("yyyy-MM-dd")
                },
                Categories = (options.Categories ?? new HashSet<WeChatDataCategory>())
                    .Select(x => x.ToString())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList()
            };

            var result = new WeChatBackupResult
            {
                ZipPath = zipPath,
                FileCount = 0,
                TotalBytes = 0
            };

            var normalizedRoot = Normalize(options.Root.RootPath);
            using (var stream = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 81920, useAsync: true))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false, entryNameEncoding: Encoding.UTF8))
            {
                for (var i = 0; i < candidates.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var candidate = candidates[i];
                    if (!File.Exists(candidate.Path))
                    {
                        continue;
                    }

                    var fullPath = Normalize(candidate.Path);
                    if (!SafePathHelper.IsPathInside(fullPath, normalizedRoot))
                    {
                        continue;
                    }

                    var rel = GetRelativePath(normalizedRoot, fullPath).Replace('\\', '/');
                    var archiveEntryPath = $"WeChatBackup/files/{options.Root.WxIdOrUserName}/{rel}";
                    var zipEntry = archive.CreateEntry(archiveEntryPath, CompressionLevel.Optimal);
                    var fileInfo = new FileInfo(fullPath);
                    zipEntry.LastWriteTime = fileInfo.LastWriteTime;

                    string sha256;
                    using (var input = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
                    using (var output = zipEntry.Open())
                    {
                        sha256 = await CopyAndHashAsync(input, output, ct).ConfigureAwait(false);
                    }

                    manifest.Entries.Add(new WeChatBackupManifestEntry
                    {
                        Rel = rel,
                        Category = candidate.Category.ToString(),
                        Size = fileInfo.Length,
                        LastWriteTime = fileInfo.LastWriteTime,
                        Sha256 = sha256
                    });

                    result.FileCount++;
                    result.TotalBytes += fileInfo.Length;
                    progress?.Report(new WeChatBackupProgress
                    {
                        Current = result.FileCount,
                        Total = candidates.Count,
                        RelativePath = rel
                    });
                }

                var manifestEntry = archive.CreateEntry("WeChatBackup/manifest.json", CompressionLevel.Optimal);
                using (var manifestStream = manifestEntry.Open())
                using (var writer = new StreamWriter(manifestStream, new UTF8Encoding(false)))
                {
                    var json = JsonConvert.SerializeObject(manifest, Formatting.Indented);
                    await writer.WriteAsync(json).ConfigureAwait(false);
                }
            }

            return result;
        }

        public async Task<WeChatBackupManifest> ReadManifestAsync(string zipPath, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(zipPath) || !File.Exists(zipPath))
            {
                throw new FileNotFoundException("备份文件不存在。", zipPath);
            }

            using (var stream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false, entryNameEncoding: Encoding.UTF8))
            {
                var entry = archive.GetEntry("WeChatBackup/manifest.json");
                if (entry == null)
                {
                    throw new InvalidDataException("备份文件中缺少 manifest.json。");
                }

                using (var entryStream = entry.Open())
                using (var reader = new StreamReader(entryStream, Encoding.UTF8))
                {
                    var json = await reader.ReadToEndAsync().ConfigureAwait(false);
                    ct.ThrowIfCancellationRequested();
                    var manifest = JsonConvert.DeserializeObject<WeChatBackupManifest>(json);
                    if (manifest == null)
                    {
                        throw new InvalidDataException("manifest.json 格式无效。");
                    }

                    return manifest;
                }
            }
        }

        public async Task<WeChatRestoreResult> RestoreAsync(
            WeChatRestoreOptions options,
            IProgress<string> progress,
            CancellationToken ct)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.ZipPath) || !File.Exists(options.ZipPath))
            {
                throw new FileNotFoundException("备份文件不存在。", options.ZipPath);
            }

            var manifest = await ReadManifestAsync(options.ZipPath, ct).ConfigureAwait(false);
            var targetRoot = options.RestoreToOriginal
                ? manifest.WechatRoot
                : options.CustomTargetRoot;

            if (string.IsNullOrWhiteSpace(targetRoot))
            {
                throw new InvalidOperationException("恢复目标目录无效。");
            }

            if (options.RestoreToOriginal && !Directory.Exists(targetRoot))
            {
                throw new InvalidOperationException("原微信目录不存在，请改用自定义目录恢复。");
            }

            Directory.CreateDirectory(targetRoot);
            var targetRootFull = Normalize(targetRoot);

            var result = new WeChatRestoreResult();
            using (var stream = new FileStream(options.ZipPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false, entryNameEncoding: Encoding.UTF8))
            {
                foreach (var entry in archive.Entries)
                {
                    ct.ThrowIfCancellationRequested();
                    if (ContainsZipSlip(entry.FullName))
                    {
                        throw new InvalidDataException("备份包含非法路径（zip slip），已终止恢复。");
                    }
                }

                foreach (var item in manifest.Entries ?? Enumerable.Empty<WeChatBackupManifestEntry>())
                {
                    ct.ThrowIfCancellationRequested();
                    if (!ShouldRestoreEntry(item, options.Categories))
                    {
                        result.SkippedByCategory++;
                        continue;
                    }

                    progress?.Report("恢复中：" + (item.Rel ?? string.Empty));

                    if (ContainsZipSlip(item.Rel))
                    {
                        throw new InvalidDataException("备份包含非法路径（zip slip），已终止恢复。");
                    }

                    if (IsForbiddenRelativePath(item.Rel))
                    {
                        result.Failed++;
                        result.Errors.Add($"禁止恢复受保护路径：{item.Rel}");
                        continue;
                    }

                    var expectedEntryName = $"WeChatBackup/files/{manifest.WxId}/{(item.Rel ?? string.Empty).Replace('\\', '/')}";
                    var zipEntry = archive.Entries.FirstOrDefault(e =>
                        string.Equals(e.FullName, expectedEntryName, StringComparison.OrdinalIgnoreCase));

                    if (zipEntry == null)
                    {
                        result.Failed++;
                        result.Errors.Add($"备份缺少文件条目：{item.Rel}");
                        continue;
                    }

                    var destination = Normalize(Path.Combine(targetRootFull, item.Rel ?? string.Empty));
                    if (!SafePathHelper.IsPathInside(destination, targetRootFull))
                    {
                        throw new InvalidDataException("检测到非法解压目标路径，已终止恢复。");
                    }

                    if (WeChatCleanupService.IsForbiddenWeChatPath(destination))
                    {
                        result.Failed++;
                        result.Errors.Add($"禁止写入受保护文件：{item.Rel}");
                        continue;
                    }

                    var finalDestination = destination;
                    if (File.Exists(finalDestination))
                    {
                        var ext = Path.GetExtension(finalDestination);
                        var baseName = finalDestination.Substring(0, finalDestination.Length - ext.Length);
                        finalDestination = $"{baseName}.restored-{DateTime.Now.Ticks}{ext}";
                    }

                    Directory.CreateDirectory(Path.GetDirectoryName(finalDestination) ?? targetRootFull);
                    AppLogService.Information("WeChat restore plan: write file {Path}", finalDestination);

                    using (var input = zipEntry.Open())
                    using (var output = new FileStream(finalDestination, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                    {
                        var actualHash = await CopyAndHashAsync(input, output, ct).ConfigureAwait(false);
                        if (!string.Equals(actualHash, item.Sha256, StringComparison.OrdinalIgnoreCase))
                        {
                            result.Failed++;
                            result.Errors.Add($"哈希校验失败：{item.Rel}");
                            try { File.Delete(finalDestination); } catch { }
                            continue;
                        }
                    }

                    result.Success++;
                }
            }

            return result;
        }

        private static bool ShouldRestoreEntry(WeChatBackupManifestEntry item, HashSet<WeChatDataCategory> categories)
        {
            if (categories == null || categories.Count == 0)
            {
                return true;
            }

            WeChatDataCategory category;
            return TryResolveEntryCategory(item, out category) && categories.Contains(category);
        }

        private static bool TryResolveEntryCategory(WeChatBackupManifestEntry item, out WeChatDataCategory category)
        {
            category = WeChatDataCategory.File;
            if (item == null)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(item.Category)
                && Enum.TryParse(item.Category, true, out category))
            {
                return true;
            }

            var rel = (item.Rel ?? string.Empty).Replace('\\', '/');
            var extension = Path.GetExtension(rel);

            if (ContainsPathSegment(rel, "Cache")
                || ContainsPathSegment(rel, "temp")
                || ContainsPathSegment(rel, "resource")
                || ContainsPathSegment(rel, "Thumb"))
            {
                category = WeChatDataCategory.Cache;
                return true;
            }

            if (ContainsPathSegment(rel, "Image")
                || ContainsPathSegment(rel, "Img")
                || IsImageExtension(extension))
            {
                category = WeChatDataCategory.Image;
                return true;
            }

            if (ContainsPathSegment(rel, "Video") || IsVideoExtension(extension))
            {
                category = WeChatDataCategory.Video;
                return true;
            }

            if (ContainsPathSegment(rel, "Voice2")
                || ContainsPathSegment(rel, "audio")
                || ContainsPathSegment(rel, "Audio")
                || ContainsPathSegment(rel, "Voice")
                || IsAudioExtension(extension))
            {
                category = WeChatDataCategory.Voice;
                return true;
            }

            if (ContainsPathSegment(rel, "File") || ContainsPathSegment(rel, "file"))
            {
                category = WeChatDataCategory.File;
                return true;
            }

            if (ContainsPathSegment(rel, "CustomEmotion") || ContainsPathSegment(rel, "MsgTemp"))
            {
                category = WeChatDataCategory.Text;
                return true;
            }

            return false;
        }

        private static bool IsImageExtension(string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".jpg":
                case ".jpeg":
                case ".png":
                case ".gif":
                case ".webp":
                case ".bmp":
                case ".heic":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsVideoExtension(string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".mp4":
                case ".mov":
                case ".avi":
                case ".mkv":
                case ".webm":
                case ".wmv":
                case ".m4v":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsAudioExtension(string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".amr":
                case ".silk":
                case ".aac":
                case ".mp3":
                case ".wav":
                case ".m4a":
                case ".ogg":
                    return true;
                default:
                    return false;
            }
        }

        private static bool ContainsPathSegment(string relativePath, string segment)
        {
            return (relativePath ?? string.Empty)
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Any(item => item.Equals(segment, StringComparison.OrdinalIgnoreCase));
        }

        private static string BuildDefaultZipPath(string outputDirectory, DateTime startDate, DateTime endDate)
        {
            var fileName = $"WeChatBackup_{DateTime.Now:yyyyMMdd-HHmmss}_{startDate:yyyyMMdd}_{endDate:yyyyMMdd}.zip";
            return Path.Combine(outputDirectory, fileName);
        }

        private static void EnsureEnoughDiskSpace(string outputDirectory, long estimatedSizeBytes)
        {
            if (estimatedSizeBytes <= 0)
            {
                return;
            }

            var root = Path.GetPathRoot(Path.GetFullPath(outputDirectory));
            if (string.IsNullOrWhiteSpace(root))
            {
                return;
            }

            var drive = new DriveInfo(root);
            var required = (long)Math.Ceiling(estimatedSizeBytes * 1.1d);
            if (drive.AvailableFreeSpace < required)
            {
                throw new InvalidOperationException("目标磁盘空间不足，请更换位置后重试。");
            }
        }

        private static async Task<string> CopyAndHashAsync(Stream source, Stream destination, CancellationToken ct)
        {
            using (var sha = SHA256.Create())
            {
                var buffer = new byte[81920];
                int read;
                while ((read = await source.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(false)) > 0)
                {
                    await destination.WriteAsync(buffer, 0, read, ct).ConfigureAwait(false);
                    sha.TransformBlock(buffer, 0, read, null, 0);
                }

                sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return BitConverter.ToString(sha.Hash ?? Array.Empty<byte>()).Replace("-", string.Empty).ToLowerInvariant();
            }
        }

        private static bool ContainsZipSlip(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return true;
            }

            if (Path.IsPathRooted(relativePath))
            {
                return true;
            }

            var normalized = relativePath.Replace('/', '\\');
            return normalized.Contains("..\\") || normalized.Contains("../");
        }

        private static bool IsForbiddenRelativePath(string rel)
        {
            if (string.IsNullOrWhiteSpace(rel))
            {
                return true;
            }

            var parts = rel.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Any(x => x.Equals("MMKV", StringComparison.OrdinalIgnoreCase)
                               || x.Equals("config", StringComparison.OrdinalIgnoreCase)
                               || x.Equals("db_storage", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            var fileName = Path.GetFileName(rel);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return true;
            }

            if (fileName.Equals("MicroMsg.db", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (fileName.StartsWith("Msg", StringComparison.OrdinalIgnoreCase) && fileName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (fileName.Equals("Login.dat", StringComparison.OrdinalIgnoreCase) || fileName.StartsWith("SafeStorage", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static string GetRelativePath(string rootPath, string fullPath)
        {
            var root = Normalize(rootPath).TrimEnd('\\') + "\\";
            var full = Normalize(fullPath);
            var rootUri = new Uri(root, UriKind.Absolute);
            var fullUri = new Uri(full, UriKind.Absolute);
            var rel = Uri.UnescapeDataString(rootUri.MakeRelativeUri(fullUri).ToString());
            return rel.Replace('/', '\\');
        }

        private static string Normalize(string path)
        {
            return Path.GetFullPath(path ?? string.Empty).TrimEnd('\\');
        }
    }

    public sealed class WeChatBackupOptions
    {
        public WeChatRoot Root { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public HashSet<WeChatDataCategory> Categories { get; set; } = new HashSet<WeChatDataCategory>();
        public string OutputDirectory { get; set; }
        public string OutputZipPath { get; set; }
    }

    public sealed class WeChatRestoreOptions
    {
        public string ZipPath { get; set; }
        public bool RestoreToOriginal { get; set; } = true;
        public string CustomTargetRoot { get; set; }
        public HashSet<WeChatDataCategory> Categories { get; set; } = new HashSet<WeChatDataCategory>();
    }

    public sealed class WeChatBackupProgress
    {
        public int Current { get; set; }
        public int Total { get; set; }
        public string RelativePath { get; set; }
    }

    public sealed class WeChatBackupResult
    {
        public string ZipPath { get; set; }
        public int FileCount { get; set; }
        public long TotalBytes { get; set; }
    }

    public sealed class WeChatRestoreResult
    {
        public int Success { get; set; }
        public int Failed { get; set; }
        public int SkippedByCategory { get; set; }
        public List<string> Errors { get; } = new List<string>();
    }

    public sealed class WeChatBackupManifest
    {
        [JsonProperty("schemaVersion")]
        public int SchemaVersion { get; set; }

        [JsonProperty("createdAt")]
        public DateTimeOffset CreatedAt { get; set; }

        [JsonProperty("machine")]
        public string Machine { get; set; }

        [JsonProperty("wechatVariant")]
        public string WechatVariant { get; set; }

        [JsonProperty("wxId")]
        public string WxId { get; set; }

        [JsonProperty("wechatRoot")]
        public string WechatRoot { get; set; }

        [JsonProperty("dateRange")]
        public WeChatBackupDateRange DateRange { get; set; } = new WeChatBackupDateRange();

        [JsonProperty("categories")]
        public List<string> Categories { get; set; } = new List<string>();

        [JsonProperty("entries")]
        public List<WeChatBackupManifestEntry> Entries { get; set; } = new List<WeChatBackupManifestEntry>();
    }

    public sealed class WeChatBackupDateRange
    {
        [JsonProperty("start")]
        public string Start { get; set; }

        [JsonProperty("end")]
        public string End { get; set; }
    }

    public sealed class WeChatBackupManifestEntry
    {
        [JsonProperty("rel")]
        public string Rel { get; set; }

        [JsonProperty("category")]
        public string Category { get; set; }

        [JsonProperty("size")]
        public long Size { get; set; }

        [JsonProperty("lastWriteTime")]
        public DateTime LastWriteTime { get; set; }

        [JsonProperty("sha256")]
        public string Sha256 { get; set; }
    }
}
