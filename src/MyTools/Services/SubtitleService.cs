using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public sealed class SubtitleCue
    {
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public string Text { get; set; }
    }

    public static class SubtitleService
    {
        private static readonly Regex TimeRangeRegex = new Regex(
            @"(?<start>\d{1,2}:\d{2}:\d{2}[,.]\d{1,3})\s*-->\s*(?<end>\d{1,2}:\d{2}:\d{2}[,.]\d{1,3})",
            RegexOptions.Compiled);

        public static async Task<IReadOnlyList<SubtitleCue>> LoadSrtAsync(string filePath, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return new List<SubtitleCue>();
            }

            var bytes = await ReadAllBytesAsync(filePath, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            var text = DetectEncoding(bytes).GetString(bytes);
            return ParseSrt(text).ToList();
        }

        public static IEnumerable<SubtitleCue> ParseSrt(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                yield break;
            }

            var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            var blocks = normalized.Split(new[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var block in blocks)
            {
                var lines = block
                    .Split(new[] { '\n' }, StringSplitOptions.None)
                    .Select(line => line.TrimEnd())
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .ToList();
                if (lines.Count < 2)
                {
                    continue;
                }

                var timeLineIndex = lines.FindIndex(line => TimeRangeRegex.IsMatch(line));
                if (timeLineIndex < 0)
                {
                    continue;
                }

                var match = TimeRangeRegex.Match(lines[timeLineIndex]);
                if (!TryParseSrtTime(match.Groups["start"].Value, out var start)
                    || !TryParseSrtTime(match.Groups["end"].Value, out var end)
                    || end <= start)
                {
                    continue;
                }

                var captionLines = lines
                    .Skip(timeLineIndex + 1)
                    .Select(CleanCaptionLine)
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .ToList();
                if (captionLines.Count == 0)
                {
                    continue;
                }

                yield return new SubtitleCue
                {
                    Start = start,
                    End = end,
                    Text = string.Join(Environment.NewLine, captionLines)
                };
            }
        }

        public static string FindSiblingSrt(string mediaPath)
        {
            if (string.IsNullOrWhiteSpace(mediaPath))
            {
                return null;
            }

            var directory = Path.GetDirectoryName(mediaPath);
            var baseName = Path.GetFileNameWithoutExtension(mediaPath);
            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(baseName))
            {
                return null;
            }

            var exact = Path.Combine(directory, baseName + ".srt");
            if (File.Exists(exact))
            {
                return exact;
            }

            return Directory.EnumerateFiles(directory, "*.srt")
                .FirstOrDefault(path => string.Equals(Path.GetFileNameWithoutExtension(path), baseName, StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryParseSrtTime(string value, out TimeSpan time)
        {
            time = TimeSpan.Zero;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var parts = value.Replace(',', '.').Split(':', '.');
            if (parts.Length != 4)
            {
                return false;
            }

            if (!int.TryParse(parts[0], out var hours)
                || !int.TryParse(parts[1], out var minutes)
                || !int.TryParse(parts[2], out var seconds)
                || !int.TryParse(parts[3].PadRight(3, '0').Substring(0, 3), out var milliseconds))
            {
                return false;
            }

            try
            {
                time = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        private static string CleanCaptionLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return string.Empty;
            }

            var cleaned = Regex.Replace(line.Trim(), @"</?[^>]+>", string.Empty);
            return cleaned.Replace(@"\N", Environment.NewLine);
        }

        private static async Task<byte[]> ReadAllBytesAsync(string filePath, CancellationToken cancellationToken)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 8192, true))
            {
                if (stream.Length > int.MaxValue)
                {
                    throw new InvalidOperationException("字幕文件过大。");
                }

                var buffer = new byte[(int)stream.Length];
                var offset = 0;
                while (offset < buffer.Length)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var read = await stream.ReadAsync(buffer, offset, buffer.Length - offset, cancellationToken).ConfigureAwait(false);
                    if (read == 0)
                    {
                        break;
                    }

                    offset += read;
                }

                if (offset == buffer.Length)
                {
                    return buffer;
                }

                var trimmed = new byte[offset];
                Buffer.BlockCopy(buffer, 0, trimmed, 0, offset);
                return trimmed;
            }
        }

        private static Encoding DetectEncoding(byte[] bytes)
        {
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                return new UTF8Encoding(true);
            }

            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                return Encoding.Unicode;
            }

            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode;
            }

            return LooksLikeUtf8(bytes) ? new UTF8Encoding(false) : Encoding.Default;
        }

        private static bool LooksLikeUtf8(byte[] bytes)
        {
            var i = 0;
            while (i < bytes.Length)
            {
                var b = bytes[i];
                if (b <= 0x7F)
                {
                    i++;
                    continue;
                }

                int extraBytes;
                if ((b & 0xE0) == 0xC0)
                {
                    extraBytes = 1;
                }
                else if ((b & 0xF0) == 0xE0)
                {
                    extraBytes = 2;
                }
                else if ((b & 0xF8) == 0xF0)
                {
                    extraBytes = 3;
                }
                else
                {
                    return false;
                }

                if (i + extraBytes >= bytes.Length)
                {
                    return false;
                }

                for (var j = 1; j <= extraBytes; j++)
                {
                    if ((bytes[i + j] & 0xC0) != 0x80)
                    {
                        return false;
                    }
                }

                i += extraBytes + 1;
            }

            return true;
        }
    }
}
