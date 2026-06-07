using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MyTools.Shared;

namespace MyTools.Services
{
    public enum MediaKind
    {
        Other,
        Image,
        Audio,
        Video,
        Pdf,
        Text,
        Markdown,
        Word,
        Excel,
        PowerPoint
    }

    public sealed class MediaFileDescriptor
    {
        public string Path { get; set; }
        public string Name { get; set; }
        public MediaKind Kind { get; set; }
        public long SizeBytes { get; set; }
        public DateTime ModifiedAt { get; set; }
    }

    public static class MediaFileTypeHelper
    {
        private static readonly string[] ImageExtensions =
        {
            ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp", ".heic"
        };

        private static readonly string[] TextExtensions =
        {
            ".txt", ".log", ".csv", ".json", ".xml", ".ini", ".cfg", ".conf", ".yaml", ".yml"
        };

        private static readonly string[] MarkdownExtensions =
        {
            ".md", ".markdown"
        };

        private static readonly string[] WordExtensions =
        {
            ".doc", ".docx", ".rtf"
        };

        private static readonly string[] ExcelExtensions =
        {
            ".xls", ".xlsx", ".xlsm", ".xlsb"
        };

        private static readonly string[] PowerPointExtensions =
        {
            ".ppt", ".pptx", ".pptm", ".pps", ".ppsx"
        };

        public static bool IsImage(string extension)
        {
            return ImageExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsAudio(string extension)
        {
            return MediaFileAssociationCore.AudioExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsVideo(string extension)
        {
            return MediaFileAssociationCore.VideoExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsAnyMedia(string extension)
        {
            return Classify(extension) != MediaKind.Other;
        }

        public static bool IsPdf(string extension)
        {
            return string.Equals(NormalizeExtension(extension), ".pdf", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsText(string extension)
        {
            return TextExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsMarkdown(string extension)
        {
            return MarkdownExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsWord(string extension)
        {
            return WordExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsExcel(string extension)
        {
            return ExcelExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static bool IsPowerPoint(string extension)
        {
            return PowerPointExtensions.Contains(NormalizeExtension(extension), StringComparer.OrdinalIgnoreCase);
        }

        public static MediaKind Classify(string extension)
        {
            if (IsImage(extension)) return MediaKind.Image;
            if (IsAudio(extension)) return MediaKind.Audio;
            if (IsVideo(extension)) return MediaKind.Video;
            if (IsPdf(extension)) return MediaKind.Pdf;
            if (IsMarkdown(extension)) return MediaKind.Markdown;
            if (IsText(extension)) return MediaKind.Text;
            if (IsWord(extension)) return MediaKind.Word;
            if (IsExcel(extension)) return MediaKind.Excel;
            if (IsPowerPoint(extension)) return MediaKind.PowerPoint;
            return MediaKind.Other;
        }

        public static string FormatFileSize(long bytes)
        {
            if (bytes < 0) bytes = 0;
            string[] units = { "B", "KB", "MB", "GB", "TB" };
            var value = (double)bytes;
            var unitIndex = 0;
            while (value >= 1024 && unitIndex < units.Length - 1)
            {
                value /= 1024;
                unitIndex++;
            }

            return unitIndex == 0
                ? string.Format("{0:0} {1}", value, units[unitIndex])
                : string.Format("{0:0.0} {1}", value, units[unitIndex]);
        }

        public static Task<IList<MediaFileDescriptor>> EnumerateMediaFilesAsync(string folder, CancellationToken cancellationToken)
        {
            return Task.Run<IList<MediaFileDescriptor>>(() =>
            {
                var result = new List<MediaFileDescriptor>();
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    return result;
                }

                try
                {
                    foreach (var file in Directory.EnumerateFiles(folder))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        FileInfo info;
                        try
                        {
                            info = new FileInfo(file);
                            if ((info.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0) continue;
                        }
                        catch
                        {
                            continue;
                        }

                        var kind = Classify(info.Extension);
                        if (kind == MediaKind.Other) continue;
                        result.Add(new MediaFileDescriptor
                        {
                            Path = info.FullName,
                            Name = info.Name,
                            Kind = kind,
                            SizeBytes = info.Length,
                            ModifiedAt = info.LastWriteTime
                        });
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    return result;
                }
                catch (IOException)
                {
                    return result;
                }

                return result
                    .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
            }, cancellationToken);
        }

        private static string NormalizeExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension)) return string.Empty;
            return extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;
        }
    }
}
