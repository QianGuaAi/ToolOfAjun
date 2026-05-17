using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public class ConvertResult
    {
        public bool Success { get; set; }
        public string OutputPath { get; set; }
        public string Message { get; set; }
        public long InputSize { get; set; }
        public long OutputSize { get; set; }
    }

    public static class MediaConvertService
    {
        // ======================== Image ========================
        public static Task<ConvertResult> ConvertImageAsync(
            string inputPath,
            string outputFormat,
            int maxWidth,
            int maxHeight,
            int quality,
            string outputDirectory,
            CancellationToken ct)
        {
            return Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (!File.Exists(inputPath))
                    return new ConvertResult { Success = false, Message = "源文件不存在。" };

                var inputSize = new FileInfo(inputPath).Length;
                var ext = NormalizeImageExtension(outputFormat);
                var outDir = ResolveOutputDirectory(inputPath, outputDirectory);
                var baseName = Path.GetFileNameWithoutExtension(inputPath);
                var outputPath = BuildUniqueOutputPath(outDir, baseName, ext);

                ImageFormat targetFormat;
                switch (ext)
                {
                    case "jpg":
                    case "jpeg":
                        targetFormat = ImageFormat.Jpeg;
                        break;
                    case "png":
                        targetFormat = ImageFormat.Png;
                        break;
                    case "bmp":
                        targetFormat = ImageFormat.Bmp;
                        break;
                    case "gif":
                        targetFormat = ImageFormat.Gif;
                        break;
                    case "tiff":
                    case "tif":
                        targetFormat = ImageFormat.Tiff;
                        break;
                    default:
                        targetFormat = ImageFormat.Png;
                        break;
                }

                using (var original = Image.FromFile(inputPath))
                {
                    ct.ThrowIfCancellationRequested();

                    int newWidth = original.Width;
                    int newHeight = original.Height;

                    if (maxWidth > 0 && maxHeight > 0 && (original.Width > maxWidth || original.Height > maxHeight))
                    {
                        double ratioW = (double)maxWidth / original.Width;
                        double ratioH = (double)maxHeight / original.Height;
                        double ratio = Math.Min(ratioW, ratioH);
                        newWidth = (int)(original.Width * ratio);
                        newHeight = (int)(original.Height * ratio);
                    }

                    using (var bmp = new Bitmap(newWidth, newHeight))
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.DrawImage(original, 0, 0, newWidth, newHeight);

                        ct.ThrowIfCancellationRequested();

                        if (targetFormat == ImageFormat.Jpeg && quality > 0 && quality <= 100)
                        {
                            using (var encoderParams = new EncoderParameters(1))
                            {
                                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
                                var jpegCodec = GetEncoder(ImageFormat.Jpeg);
                                if (jpegCodec != null)
                                {
                                    bmp.Save(outputPath, jpegCodec, encoderParams);
                                }
                                else
                                {
                                    bmp.Save(outputPath, targetFormat);
                                }
                            }
                        }
                        else
                        {
                            bmp.Save(outputPath, targetFormat);
                        }
                    }
                }

                var outputSize = new FileInfo(outputPath).Length;
                return new ConvertResult
                {
                    Success = true,
                    OutputPath = outputPath,
                    InputSize = inputSize,
                    OutputSize = outputSize,
                    Message = $"转换完成：{FormatSize(inputSize)} → {FormatSize(outputSize)}"
                };
            }, ct);
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            foreach (var codec in ImageCodecInfo.GetImageDecoders())
            {
                if (codec.FormatID == format.Guid) return codec;
            }
            return null;
        }

        // ======================== FFmpeg (Audio/Video) ========================
        public static string FindFfmpeg()
        {
            foreach (var candidate in EnumerateFfmpegCandidates())
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = candidate,
                        Arguments = "-version",
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };
                    using (var p = Process.Start(psi))
                    {
                        if (p == null)
                        {
                            continue;
                        }

                        var stdoutTask = Task.Run(() => p.StandardOutput.ReadToEnd());
                        var stderrTask = Task.Run(() => p.StandardError.ReadToEnd());
                        var exited = p.WaitForExit(5000);
                        if (!exited) try { p.Kill(); } catch { }
                        try { Task.WaitAll(new Task[] { stdoutTask, stderrTask }, 1000); } catch { }
                        if (exited && p.ExitCode == 0) return candidate;
                    }
                }
                catch { }
            }
            return null;
        }

        public static async Task<ConvertResult> ConvertMediaAsync(
            string ffmpegPath,
            string inputPath,
            string outputFormat,
            string extraArgs,
            string outputDirectory,
            IProgress<string> progress,
            CancellationToken ct)
        {
            if (string.IsNullOrEmpty(ffmpegPath))
                return new ConvertResult { Success = false, Message = "未找到 ffmpeg.exe，请先安装 FFmpeg。" };
            if (!File.Exists(inputPath))
                return new ConvertResult { Success = false, Message = "源文件不存在。" };

            var inputSize = new FileInfo(inputPath).Length;
            var ext = NormalizeMediaExtension(outputFormat);
            var outDir = ResolveOutputDirectory(inputPath, outputDirectory);
            var baseName = Path.GetFileNameWithoutExtension(inputPath);
            var outputPath = BuildUniqueOutputPath(outDir, baseName, ext);

            var args = $"-i \"{inputPath}\" {extraArgs} -y \"{outputPath}\"";
            progress?.Report($"正在转换… 输出：{Path.GetFileName(outputPath)}");

            var result = await Task.Run(() =>
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = ffmpegPath,
                        Arguments = args,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true
                    };
                    using (var p = Process.Start(psi))
                    {
                        if (p == null)
                        {
                            return new ConvertResult { Success = false, Message = "ffmpeg 启动失败。" };
                        }

                        using (ct.Register(() =>
                        {
                            try
                            {
                                if (!p.HasExited)
                                {
                                    p.Kill();
                                }
                            }
                            catch
                            {
                            }
                        }))
                        {
                            // Read stderr async to avoid deadlock
                            string stderr = null;
                            var stderrTask = Task.Run(() => p.StandardError.ReadToEnd());
                            p.StandardOutput.ReadToEnd(); // drain stdout

                            if (!p.WaitForExit(300_000)) // 5 min timeout
                            {
                                try { p.Kill(); } catch { }
                                return new ConvertResult { Success = false, Message = "ffmpeg 超时（超过 5 分钟）。" };
                            }
                            stderr = stderrTask.Result;
                            ct.ThrowIfCancellationRequested();

                            if (p.ExitCode != 0)
                                return new ConvertResult { Success = false, Message = $"ffmpeg 返回错误码 {p.ExitCode}：\n{TruncateEnd(stderr, 500)}" };

                            if (!File.Exists(outputPath))
                                return new ConvertResult { Success = false, Message = "输出文件未生成。" };

                            var outputSize = new FileInfo(outputPath).Length;
                            return new ConvertResult
                            {
                                Success = true,
                                OutputPath = outputPath,
                                InputSize = inputSize,
                                OutputSize = outputSize,
                                Message = $"转换完成：{FormatSize(inputSize)} → {FormatSize(outputSize)}"
                            };
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    return new ConvertResult { Success = false, Message = $"执行 ffmpeg 出错：{ex.Message}" };
                }
            }, ct).ConfigureAwait(false);

            return result;
        }

        public static async Task<ConvertResult> CaptureVideoFrameAsync(
            string ffmpegPath,
            string inputPath,
            double positionSeconds,
            string outputDirectory,
            CancellationToken ct)
        {
            if (string.IsNullOrEmpty(ffmpegPath))
                return new ConvertResult { Success = false, Message = "未找到 ffmpeg.exe，请先安装 FFmpeg。" };
            if (!File.Exists(inputPath))
                return new ConvertResult { Success = false, Message = "源文件不存在。" };

            if (double.IsNaN(positionSeconds) || double.IsInfinity(positionSeconds) || positionSeconds < 0)
            {
                positionSeconds = 0;
            }

            var outDir = ResolveOutputDirectory(inputPath, outputDirectory);
            var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(inputPath));
            var timeName = TimeSpan.FromSeconds(positionSeconds).ToString(@"hhmmss");
            var outputPath = BuildUniqueOutputPath(outDir, $"{baseName}_frame_{timeName}", "png", string.Empty);
            var args = $"-ss {positionSeconds.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)} -i \"{inputPath}\" -frames:v 1 -q:v 2 -y \"{outputPath}\"";

            return await Task.Run(() =>
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = ffmpegPath,
                        Arguments = args,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true
                    };
                    using (var p = Process.Start(psi))
                    {
                        if (p == null)
                        {
                            return new ConvertResult { Success = false, Message = "ffmpeg 启动失败。" };
                        }

                        using (ct.Register(() =>
                        {
                            try
                            {
                                if (!p.HasExited)
                                {
                                    p.Kill();
                                }
                            }
                            catch
                            {
                            }
                        }))
                        {
                            var stderrTask = Task.Run(() => p.StandardError.ReadToEnd());
                            p.StandardOutput.ReadToEnd();
                            if (!p.WaitForExit(60_000))
                            {
                                try { p.Kill(); } catch { }
                                return new ConvertResult { Success = false, Message = "ffmpeg 截帧超时。" };
                            }

                            var stderr = stderrTask.Result;
                            ct.ThrowIfCancellationRequested();

                            if (p.ExitCode != 0)
                            {
                                return new ConvertResult { Success = false, Message = $"ffmpeg 返回错误码 {p.ExitCode}：\n{TruncateEnd(stderr, 500)}" };
                            }

                            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
                            {
                                return new ConvertResult { Success = false, Message = "截帧文件未生成。" };
                            }

                            return new ConvertResult
                            {
                                Success = true,
                                OutputPath = outputPath,
                                OutputSize = new FileInfo(outputPath).Length,
                                Message = "截帧完成。"
                            };
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    return new ConvertResult { Success = false, Message = $"执行 ffmpeg 截帧出错：{ex.Message}" };
                }
            }, ct).ConfigureAwait(false);
        }

        public static async Task<ConvertResult> GenerateAudioWaveformAsync(
            string ffmpegPath,
            string inputPath,
            string outputDirectory,
            int width,
            int height,
            CancellationToken ct)
        {
            if (string.IsNullOrEmpty(ffmpegPath))
                return new ConvertResult { Success = false, Message = "未找到 ffmpeg.exe，请先安装 FFmpeg。" };
            if (!File.Exists(inputPath))
                return new ConvertResult { Success = false, Message = "源文件不存在。" };

            width = Math.Max(320, Math.Min(1920, width));
            height = Math.Max(80, Math.Min(480, height));
            var outDir = ResolveOutputDirectory(inputPath, outputDirectory);
            var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(inputPath));
            var outputPath = BuildUniqueOutputPath(outDir, $"{baseName}_waveform", "png", string.Empty);
            var args = $"-i \"{inputPath}\" -filter_complex \"aformat=channel_layouts=mono,showwavespic=s={width}x{height}:colors=#2563EB\" -frames:v 1 -y \"{outputPath}\"";

            return await Task.Run(() =>
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = ffmpegPath,
                        Arguments = args,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true
                    };
                    using (var p = Process.Start(psi))
                    {
                        if (p == null)
                        {
                            return new ConvertResult { Success = false, Message = "ffmpeg 启动失败。" };
                        }

                        using (ct.Register(() =>
                        {
                            try
                            {
                                if (!p.HasExited)
                                {
                                    p.Kill();
                                }
                            }
                            catch
                            {
                            }
                        }))
                        {
                            var stderrTask = Task.Run(() => p.StandardError.ReadToEnd());
                            p.StandardOutput.ReadToEnd();
                            if (!p.WaitForExit(90_000))
                            {
                                try { p.Kill(); } catch { }
                                return new ConvertResult { Success = false, Message = "ffmpeg 生成波形超时。" };
                            }

                            var stderr = stderrTask.Result;
                            ct.ThrowIfCancellationRequested();

                            if (p.ExitCode != 0)
                            {
                                return new ConvertResult { Success = false, Message = $"ffmpeg 返回错误码 {p.ExitCode}：\n{TruncateEnd(stderr, 500)}" };
                            }

                            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
                            {
                                return new ConvertResult { Success = false, Message = "波形图片未生成。" };
                            }

                            return new ConvertResult
                            {
                                Success = true,
                                OutputPath = outputPath,
                                OutputSize = new FileInfo(outputPath).Length,
                                Message = "波形生成完成。"
                            };
                        }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    return new ConvertResult { Success = false, Message = $"执行 ffmpeg 生成波形出错：{ex.Message}" };
                }
            }, ct).ConfigureAwait(false);
        }

        private static IEnumerable<string> EnumerateFfmpegCandidates()
        {
            yield return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NativeBinaries", "ffmpeg", "ffmpeg.exe");
            yield return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe");
            yield return @"C:\ffmpeg\bin\ffmpeg.exe";
            yield return @"C:\Program Files\ffmpeg\bin\ffmpeg.exe";

            var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var pathPart in pathValue.Split(new[] { Path.PathSeparator }, StringSplitOptions.RemoveEmptyEntries))
            {
                string candidate;
                try
                {
                    candidate = Path.Combine(pathPart.Trim(), "ffmpeg.exe");
                }
                catch
                {
                    continue;
                }

                yield return candidate;
            }
        }

        private static string BuildUniqueOutputPath(string directory, string baseName, string extension)
        {
            return BuildUniqueOutputPath(directory, baseName, extension, "_converted");
        }

        private static string BuildUniqueOutputPath(string directory, string baseName, string extension, string suffix)
        {
            Directory.CreateDirectory(directory);
            var safeExtension = extension.TrimStart('.');
            var safeBaseName = SanitizeFileName(baseName);
            var safeSuffix = suffix ?? string.Empty;
            var first = Path.Combine(directory, $"{safeBaseName}{safeSuffix}.{safeExtension}");
            if (!File.Exists(first))
            {
                return first;
            }

            for (var i = 2; i < 1000; i++)
            {
                var candidate = Path.Combine(directory, $"{safeBaseName}{safeSuffix}_{i}.{safeExtension}");
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return Path.Combine(directory, $"{safeBaseName}{safeSuffix}_{DateTime.Now:yyyyMMddHHmmssfff}.{safeExtension}");
        }

        private static string SanitizeFileName(string value)
        {
            var safeName = string.IsNullOrWhiteSpace(value) ? "output" : value.Trim();
            foreach (var ch in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(ch, '_');
            }

            return string.IsNullOrWhiteSpace(safeName) ? "output" : safeName;
        }

        private static string ResolveOutputDirectory(string inputPath, string outputDirectory)
        {
            if (!string.IsNullOrWhiteSpace(outputDirectory))
            {
                return outputDirectory;
            }

            return Path.GetDirectoryName(inputPath) ?? AppDomain.CurrentDomain.BaseDirectory;
        }

        private static string NormalizeImageExtension(string extension)
        {
            var ext = (extension ?? string.Empty).Trim().TrimStart('.').ToLowerInvariant();
            switch (ext)
            {
                case "jpg":
                case "jpeg":
                case "png":
                case "bmp":
                case "gif":
                case "tif":
                case "tiff":
                    return ext;
                default:
                    return "png";
            }
        }

        private static string NormalizeMediaExtension(string extension)
        {
            var ext = (extension ?? string.Empty).Trim().TrimStart('.').ToLowerInvariant();
            switch (ext)
            {
                case "mp4":
                case "mkv":
                case "avi":
                case "mov":
                case "flv":
                case "wmv":
                case "mp3":
                case "wav":
                case "aac":
                case "flac":
                case "ogg":
                case "m4a":
                    return ext;
                default:
                    return "mp3";
            }
        }

        private static string FormatSize(long bytes)
        {
            if (bytes >= 1_073_741_824) return $"{bytes / 1_073_741_824.0:0.##} GB";
            if (bytes >= 1_048_576) return $"{bytes / 1_048_576.0:0.##} MB";
            return $"{bytes / 1024.0:0.##} KB";
        }

        private static string TruncateEnd(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            if (s.Length <= max) return s;
            return "…" + s.Substring(s.Length - max);
        }
    }
}
