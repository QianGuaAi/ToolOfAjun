using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public sealed class RecordingService
    {
        private readonly object _syncRoot = new object();
        private Process _activeProcess;
        private RecordingTaskKind _activeTaskKind = RecordingTaskKind.None;
        private string _activeOutputPath = string.Empty;
        private DateTime _startedAtUtc = DateTime.UtcNow;

        public string ExpectedFfmpegPath =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NativeBinaries", "ffmpeg", "ffmpeg.exe");

        public bool TryGetFfmpegPath(out string ffmpegPath)
        {
            if (File.Exists(ExpectedFfmpegPath))
            {
                ffmpegPath = ExpectedFfmpegPath;
                return true;
            }

            var fallback = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe");
            if (File.Exists(fallback))
            {
                ffmpegPath = fallback;
                return true;
            }

            ffmpegPath = string.Empty;
            return false;
        }

        public async Task<string> ResolvePreferredAudioDeviceAsync(CancellationToken cancellationToken)
        {
            if (!TryGetFfmpegPath(out var ffmpegPath))
            {
                return string.Empty;
            }

            var devices = await ListAudioDevicesAsync(ffmpegPath, cancellationToken).ConfigureAwait(false);
            var virtualDevice = devices.FirstOrDefault(device =>
                device.IndexOf("virtual-audio-capturer", StringComparison.OrdinalIgnoreCase) >= 0);
            return virtualDevice ?? devices.FirstOrDefault() ?? string.Empty;
        }

        public async Task StartVideoRecordingAsync(
            RecordingRegion region,
            string outputPath,
            string audioDeviceName,
            CancellationToken cancellationToken)
        {
            if (!TryGetFfmpegPath(out var ffmpegPath))
            {
                throw new FileNotFoundException("未找到 ffmpeg.exe。", ExpectedFfmpegPath);
            }

            ValidateRecordingRegion(region);
            ValidateOutputPath(outputPath);
            EnsureIdle();

            var arguments = BuildVideoRecordingArguments(region, outputPath, audioDeviceName);
            await StartProcessInternalAsync(ffmpegPath, arguments, outputPath, RecordingTaskKind.Video, cancellationToken).ConfigureAwait(false);
            AppLogService.Information("Video recording started: {Path}", outputPath);
        }

        public async Task StartAudioOnlyAsync(string outputPath, CancellationToken cancellationToken)
        {
            if (!TryGetFfmpegPath(out var ffmpegPath))
            {
                throw new FileNotFoundException("未找到 ffmpeg.exe。", ExpectedFfmpegPath);
            }

            ValidateOutputPath(outputPath);
            EnsureIdle();

            var audioDeviceName = await ResolvePreferredAudioDeviceAsync(cancellationToken).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(audioDeviceName))
            {
                throw new InvalidOperationException("未检测到任何音频输入设备，无法录音。");
            }

            var arguments = BuildAudioRecordingArguments(outputPath, audioDeviceName);
            await StartProcessInternalAsync(ffmpegPath, arguments, outputPath, RecordingTaskKind.Audio, cancellationToken).ConfigureAwait(false);
            AppLogService.Information("Audio recording started: {Path}", outputPath);
        }

        public Task<RecordingStopResult> StopVideoRecordingAsync()
        {
            return StopCurrentAsync(RecordingTaskKind.Video);
        }

        public Task<RecordingStopResult> StopAudioOnlyAsync()
        {
            return StopCurrentAsync(RecordingTaskKind.Audio);
        }

        private async Task<RecordingStopResult> StopCurrentAsync(RecordingTaskKind expectedTaskKind)
        {
            Process process;
            DateTime startedAtUtc;
            string outputPath;
            lock (_syncRoot)
            {
                if (_activeProcess == null || _activeProcess.HasExited || _activeTaskKind != expectedTaskKind)
                {
                    return new RecordingStopResult();
                }

                process = _activeProcess;
                startedAtUtc = _startedAtUtc;
                outputPath = _activeOutputPath;
            }

            var timedOut = false;
            try
            {
                try
                {
                    await process.StandardInput.WriteLineAsync("q").ConfigureAwait(false);
                    await process.StandardInput.FlushAsync().ConfigureAwait(false);
                }
                catch
                {
                    // Ignore stdin errors and continue with exit wait.
                }

                var exited = await WaitForExitAsync(process, 5000).ConfigureAwait(false);
                if (!exited)
                {
                    timedOut = true;
                    process.Kill();
                    await WaitForExitAsync(process, 2000).ConfigureAwait(false);
                }

                var fileSize = 0L;
                try
                {
                    if (!string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath))
                    {
                        fileSize = new FileInfo(outputPath).Length;
                    }
                }
                catch
                {
                    fileSize = 0;
                }

                var durationSeconds = Math.Max(0, (long)(DateTime.UtcNow - startedAtUtc).TotalSeconds);
                if (expectedTaskKind == RecordingTaskKind.Audio)
                {
                    AppLogService.Information("Audio recording stopped: duration {Seconds}s, size {Bytes}B", durationSeconds, fileSize);
                }
                else
                {
                    AppLogService.Information("Video recording stopped: duration {Seconds}s, size {Bytes}B", durationSeconds, fileSize);
                }

                return new RecordingStopResult
                {
                    TimedOut = timedOut,
                    OutputPath = outputPath,
                    DurationSeconds = durationSeconds,
                    FileSizeBytes = fileSize
                };
            }
            finally
            {
                CleanupProcessState();
                process.Dispose();
            }
        }

        private async Task StartProcessInternalAsync(
            string ffmpegPath,
            string arguments,
            string outputPath,
            RecordingTaskKind taskKind,
            CancellationToken cancellationToken)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? AppDomain.CurrentDomain.BaseDirectory);
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = arguments,
                    RedirectStandardInput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                },
                EnableRaisingEvents = true
            };

            process.ErrorDataReceived += HandleFfmpegErrorData;

            cancellationToken.ThrowIfCancellationRequested();
            if (!process.Start())
            {
                process.Dispose();
                throw new InvalidOperationException("无法启动 ffmpeg 进程。");
            }

            process.BeginErrorReadLine();
            lock (_syncRoot)
            {
                _activeProcess = process;
                _activeTaskKind = taskKind;
                _activeOutputPath = outputPath;
                _startedAtUtc = DateTime.UtcNow;
            }

            await Task.CompletedTask.ConfigureAwait(false);
        }

        private static void HandleFfmpegErrorData(object sender, DataReceivedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.Data))
            {
                return;
            }

            if (e.Data.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                AppLogService.Error("ffmpeg: {Line}", e.Data);
            }
            else
            {
                AppLogService.Information("ffmpeg: {Line}", e.Data);
            }
        }

        private async Task<List<string>> ListAudioDevicesAsync(string ffmpegPath, CancellationToken cancellationToken)
        {
            var deviceLines = new List<string>();
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = "-hide_banner -list_devices true -f dshow -i dummy",
                    RedirectStandardError = true,
                    RedirectStandardOutput = false,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                if (!process.Start())
                {
                    return deviceLines;
                }

                var stderr = await process.StandardError.ReadToEndAsync().ConfigureAwait(false);
                await WaitForExitAsync(process, 8000).ConfigureAwait(false);

                var regex = new Regex("\"(?<name>.+?)\"\\s+\\(audio\\)", RegexOptions.IgnoreCase);
                foreach (var line in stderr.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var match = regex.Match(line);
                    if (!match.Success)
                    {
                        continue;
                    }

                    var name = match.Groups["name"].Value.Trim();
                    if (!string.IsNullOrWhiteSpace(name)
                        && !deviceLines.Any(item => string.Equals(item, name, StringComparison.OrdinalIgnoreCase)))
                    {
                        deviceLines.Add(name);
                    }
                }
            }

            return deviceLines;
        }

        private void EnsureIdle()
        {
            lock (_syncRoot)
            {
                if (_activeProcess != null && !_activeProcess.HasExited)
                {
                    throw new InvalidOperationException("当前已有录像/录音任务在进行，请先停止。");
                }
            }
        }

        private void CleanupProcessState()
        {
            lock (_syncRoot)
            {
                _activeProcess = null;
                _activeTaskKind = RecordingTaskKind.None;
                _activeOutputPath = string.Empty;
                _startedAtUtc = DateTime.UtcNow;
            }
        }

        private static void ValidateRecordingRegion(RecordingRegion region)
        {
            if (region.Width <= 0 || region.Height <= 0)
            {
                throw new InvalidOperationException("录制区域无效，请重新调整区域窗口大小后再试。");
            }
        }

        private static void ValidateOutputPath(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new InvalidOperationException("输出路径不能为空。");
            }
        }

        private static string BuildVideoRecordingArguments(RecordingRegion region, string outputPath, string audioDeviceName)
        {
            var args = new StringBuilder();
            args.Append("-y -f gdigrab -framerate 30 ");
            args.AppendFormat("-offset_x {0} -offset_y {1} ", region.X, region.Y);
            args.AppendFormat("-video_size {0}x{1} -i desktop ", region.Width, region.Height);

            var hasAudio = !string.IsNullOrWhiteSpace(audioDeviceName);
            if (hasAudio)
            {
                args.AppendFormat("-f dshow -i audio=\"{0}\" ", EscapeDshowDeviceName(audioDeviceName));
            }

            args.Append("-c:v libx264 -preset veryfast -crf 23 -pix_fmt yuv420p ");
            if (hasAudio)
            {
                args.Append("-c:a aac -b:a 128k ");
            }

            args.AppendFormat("-movflags +faststart \"{0}\"", outputPath);
            return args.ToString();
        }

        private static string BuildAudioRecordingArguments(string outputPath, string audioDeviceName)
        {
            return string.Format(
                "-y -f dshow -i audio=\"{0}\" -c:a aac -b:a 128k -ar 44100 -ac 2 \"{1}\"",
                EscapeDshowDeviceName(audioDeviceName),
                outputPath);
        }

        private static string EscapeDshowDeviceName(string value)
        {
            return (value ?? string.Empty).Replace("\"", "\\\"");
        }

        private static Task<bool> WaitForExitAsync(Process process, int timeoutMilliseconds)
        {
            if (process == null)
            {
                return Task.FromResult(true);
            }

            if (process.HasExited)
            {
                return Task.FromResult(true);
            }

            var tcs = new TaskCompletionSource<bool>();
            EventHandler handler = null;
            var timer = new Timer(_ =>
            {
                process.Exited -= handler;
                tcs.TrySetResult(false);
            }, null, timeoutMilliseconds, Timeout.Infinite);
            handler = (sender, args) =>
            {
                timer.Dispose();
                process.Exited -= handler;
                tcs.TrySetResult(true);
            };

            process.EnableRaisingEvents = true;
            process.Exited += handler;
            if (process.HasExited)
            {
                timer.Dispose();
                process.Exited -= handler;
                tcs.TrySetResult(true);
            }

            return tcs.Task;
        }

        private enum RecordingTaskKind
        {
            None = 0,
            Video = 1,
            Audio = 2
        }
    }

    public struct RecordingRegion
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    public sealed class RecordingStopResult
    {
        public bool TimedOut { get; set; }
        public string OutputPath { get; set; } = string.Empty;
        public long DurationSeconds { get; set; }
        public long FileSizeBytes { get; set; }
    }
}
