using System;
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
            CancellationToken ct)
        {
            return Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                if (!File.Exists(inputPath))
                    return new ConvertResult { Success = false, Message = "源文件不存在。" };

                var inputSize = new FileInfo(inputPath).Length;
                var ext = outputFormat.ToLowerInvariant().TrimStart('.');
                var outDir = Path.GetDirectoryName(inputPath);
                var baseName = Path.GetFileNameWithoutExtension(inputPath);
                var outputPath = Path.Combine(outDir, $"{baseName}_converted.{ext}");

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
                        ext = "png";
                        outputPath = Path.Combine(outDir, $"{baseName}_converted.{ext}");
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
            // Check common locations
            var candidates = new[]
            {
                "ffmpeg",
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe"),
                @"C:\ffmpeg\bin\ffmpeg.exe",
                @"C:\Program Files\ffmpeg\bin\ffmpeg.exe"
            };

            foreach (var candidate in candidates)
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
                        p.StandardOutput.ReadToEnd();
                        p.StandardError.ReadToEnd();
                        if (p.WaitForExit(5000) && p.ExitCode == 0) return candidate;
                        if (!p.HasExited) try { p.Kill(); } catch { }
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
            IProgress<string> progress,
            CancellationToken ct)
        {
            if (string.IsNullOrEmpty(ffmpegPath))
                return new ConvertResult { Success = false, Message = "未找到 ffmpeg.exe，请先安装 FFmpeg。" };
            if (!File.Exists(inputPath))
                return new ConvertResult { Success = false, Message = "源文件不存在。" };

            var inputSize = new FileInfo(inputPath).Length;
            var ext = outputFormat.ToLowerInvariant().TrimStart('.');
            var outDir = Path.GetDirectoryName(inputPath);
            var baseName = Path.GetFileNameWithoutExtension(inputPath);
            var outputPath = Path.Combine(outDir, $"{baseName}_converted.{ext}");

            var args = $"-i \"{inputPath}\" {extraArgs} -y \"{outputPath}\"";
            progress?.Report($"正在转换… ffmpeg {args}");

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
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    return new ConvertResult { Success = false, Message = $"执行 ffmpeg 出错：{ex.Message}" };
                }
            }, ct).ConfigureAwait(false);

            return result;
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
