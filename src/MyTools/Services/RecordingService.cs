using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public sealed class RecordingService
    {
        private const int MaxImmediateFfmpegInfoLines = 12;
        private const int FfmpegStartProbeMilliseconds = 700;
        private static readonly TimeSpan FfmpegInfoLogInterval = TimeSpan.FromSeconds(30);
        private readonly object _syncRoot = new object();
        private readonly object _ffmpegLogSyncRoot = new object();
        private Process _activeProcess;
        private RecordingTaskKind _activeTaskKind = RecordingTaskKind.None;
        private string _activeOutputPath = string.Empty;
        private string _activeProcessOutputPath = string.Empty;
        private string _activeLoopbackAudioPath = string.Empty;
        private string _activeFfmpegPath = string.Empty;
        private WasapiLoopbackAudioRecorder _activeLoopbackRecorder;
        private bool _activeVideoHasAudio;
        private DateTime _startedAtUtc = DateTime.UtcNow;
        private DateTime _lastFfmpegInfoLogUtc = DateTime.MinValue;
        private int _ffmpegInfoLinesLogged;
        private int _suppressedFfmpegInfoLines;

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

            if (TryFindFfmpegFromPath(out ffmpegPath))
            {
                return true;
            }

            ffmpegPath = string.Empty;
            return false;
        }

        private static bool TryFindFfmpegFromPath(out string ffmpegPath)
        {
            ffmpegPath = string.Empty;
            var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var pathPart in pathValue.Split(new[] { Path.PathSeparator }, StringSplitOptions.RemoveEmptyEntries))
            {
                try
                {
                    var candidate = Path.Combine(pathPart.Trim(), "ffmpeg.exe");
                    if (File.Exists(candidate))
                    {
                        ffmpegPath = candidate;
                        return true;
                    }
                }
                catch
                {
                    // Ignore malformed PATH entries.
                }
            }

            return false;
        }

        public async Task StartVideoRecordingAsync(
            RecordingRegion region,
            string outputPath,
            CancellationToken cancellationToken)
        {
            await StartVideoRecordingAsync(region, outputPath, false, cancellationToken).ConfigureAwait(false);
        }

        public async Task StartVideoRecordingAsync(
            RecordingRegion region,
            string outputPath,
            bool gifMode,
            CancellationToken cancellationToken)
        {
            await StartVideoRecordingAsync(region, outputPath, gifMode, RecordingOptions.Default, cancellationToken).ConfigureAwait(false);
        }

        public async Task StartVideoRecordingAsync(
            RecordingRegion region,
            string outputPath,
            bool gifMode,
            RecordingOptions options,
            CancellationToken cancellationToken)
        {
            if (!TryGetFfmpegPath(out var ffmpegPath))
            {
                throw new FileNotFoundException("未找到 ffmpeg.exe。", ExpectedFfmpegPath);
            }

            ValidateRecordingRegion(region);
            region = NormalizeVideoRecordingRegion(region);
            ValidateOutputPath(outputPath);
            EnsureIdle();

            var loopbackAudioPath = string.Empty;
            var loopbackRecorder = gifMode ? null : TryStartLoopbackRecorder(outputPath, out loopbackAudioPath);
            var processOutputPath = loopbackRecorder != null
                ? BuildTempSiblingPath(outputPath, "video", ".mp4")
                : outputPath;
            options = RecordingOptions.Normalize(options);
            var arguments = gifMode
                ? BuildGifRecordingArguments(region, processOutputPath, options)
                : BuildVideoRecordingArguments(region, processOutputPath, options);

            try
            {
                await StartProcessInternalAsync(
                    ffmpegPath,
                    arguments,
                    outputPath,
                    processOutputPath,
                    RecordingTaskKind.Video,
                    loopbackRecorder != null,
                    loopbackRecorder,
                    loopbackAudioPath,
                    cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                loopbackRecorder?.Stop();
                TryDeleteFile(loopbackAudioPath);
                TryDeleteFile(processOutputPath);
                throw;
            }

            var earlyExitCode = await WaitForActiveProcessEarlyExitCodeAsync(FfmpegStartProbeMilliseconds).ConfigureAwait(false);
            if (earlyExitCode.HasValue)
            {
                CleanupFailedStart();
                TryDeleteEmptyOutput(outputPath);
                throw new InvalidOperationException("ffmpeg 启动后立即退出，录像未开始。请查看日志确认屏幕捕获权限或输出目录是否可写。");
            }

            AppLogService.Information("Video recording started: {Path}, audio={HasAudio}, gif={GifMode}", outputPath, _activeVideoHasAudio, gifMode);
        }

        public async Task StartAudioOnlyAsync(string outputPath, CancellationToken cancellationToken)
        {
            ValidateOutputPath(outputPath);
            EnsureIdle();
            var format = AudioRecordingFormat.FromPath(outputPath);
            var ffmpegPath = string.Empty;
            if (format.RequiresFfmpeg && !TryGetFfmpegPath(out ffmpegPath))
            {
                throw new FileNotFoundException("未找到 ffmpeg.exe。", ExpectedFfmpegPath);
            }

            var loopbackRecorder = TryStartLoopbackRecorder(outputPath, out var loopbackAudioPath);
            if (loopbackRecorder == null)
            {
                throw new InvalidOperationException("无法启动系统声音录制，请确认 Windows 音频服务正在运行并且存在默认播放设备。");
            }

            lock (_syncRoot)
            {
                _activeProcess = null;
                _activeTaskKind = RecordingTaskKind.Audio;
                _activeOutputPath = outputPath;
                _activeProcessOutputPath = string.Empty;
                _activeLoopbackRecorder = loopbackRecorder;
                _activeLoopbackAudioPath = loopbackAudioPath;
                _activeFfmpegPath = ffmpegPath;
                _activeVideoHasAudio = false;
                _startedAtUtc = DateTime.UtcNow;
            }

            await Task.CompletedTask.ConfigureAwait(false);
            AppLogService.Information("System audio recording started: {Path}, format={Format}", outputPath, format.Extension);
        }

        public Task<RecordingStopResult> StopVideoRecordingAsync()
        {
            return StopCurrentAsync(RecordingTaskKind.Video);
        }

        public Task<RecordingStopResult> StopAudioOnlyAsync()
        {
            return StopCurrentAsync(RecordingTaskKind.Audio);
        }

        public RecordingAudioLevelSnapshot GetActiveAudioLevelSnapshot()
        {
            WasapiLoopbackAudioRecorder loopbackRecorder;
            lock (_syncRoot)
            {
                loopbackRecorder = _activeLoopbackRecorder;
            }

            var stats = loopbackRecorder?.ConsumeLiveAudioLevelStats() ?? new AudioLevelStats();
            return new RecordingAudioLevelSnapshot
            {
                CurrentSampleLevel = stats.CurrentSampleLevel,
                MaxSampleLevel = stats.MaxSampleLevel,
                BytesRecorded = stats.BytesRecorded,
                HasAudibleAudio = stats.HasAudibleAudio
            };
        }

        private async Task<RecordingStopResult> StopCurrentAsync(RecordingTaskKind expectedTaskKind)
        {
            Process process;
            DateTime startedAtUtc;
            string outputPath;
            string processOutputPath;
            string loopbackAudioPath;
            string ffmpegPath;
            WasapiLoopbackAudioRecorder loopbackRecorder;
            bool videoHasAudio;
            lock (_syncRoot)
            {
                if (_activeTaskKind != expectedTaskKind)
                {
                    return new RecordingStopResult();
                }

                if (expectedTaskKind == RecordingTaskKind.Video && _activeProcess == null)
                {
                    return new RecordingStopResult();
                }

                if (expectedTaskKind == RecordingTaskKind.Audio && _activeLoopbackRecorder == null)
                {
                    return new RecordingStopResult();
                }

                process = _activeProcess;
                startedAtUtc = _startedAtUtc;
                outputPath = _activeOutputPath;
                processOutputPath = _activeProcessOutputPath;
                loopbackAudioPath = _activeLoopbackAudioPath;
                ffmpegPath = _activeFfmpegPath;
                loopbackRecorder = _activeLoopbackRecorder;
                videoHasAudio = _activeVideoHasAudio;
            }

            var timedOut = false;
            try
            {
                if (process != null)
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
                }

                loopbackRecorder?.Stop();

                if (!timedOut && expectedTaskKind == RecordingTaskKind.Video && videoHasAudio)
                {
                    try
                    {
                        videoHasAudio = await MuxVideoAndLoopbackAudioAsync(ffmpegPath, processOutputPath, loopbackAudioPath, loopbackRecorder, outputPath).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        AppLogService.Warning("Muxing loopback audio failed; keeping video-only output. {Msg}", ex.Message);
                        if (!string.IsNullOrWhiteSpace(processOutputPath) && File.Exists(processOutputPath))
                        {
                            File.Copy(processOutputPath, outputPath, true);
                        }

                        videoHasAudio = false;
                    }
                }
                else if (!timedOut && expectedTaskKind == RecordingTaskKind.Audio)
                {
                    await SaveLoopbackAudioAsync(ffmpegPath, loopbackAudioPath, loopbackRecorder, outputPath).ConfigureAwait(false);
                }

                var fileSize = GetFileSize(outputPath);

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
                    FileSizeBytes = fileSize,
                    VideoHasAudio = expectedTaskKind == RecordingTaskKind.Video && videoHasAudio
                };
            }
            finally
            {
                try
                {
                    if (process != null)
                    {
                        process.ErrorDataReceived -= HandleFfmpegErrorData;
                        process.Dispose();
                    }

                    loopbackRecorder?.Dispose();
                    TryDeleteTempSiblingFile(processOutputPath, outputPath);
                    TryDeleteFile(loopbackAudioPath);
                }
                finally
                {
                    CleanupProcessState();
                }
            }
        }

        private async Task StartProcessInternalAsync(
            string ffmpegPath,
            string arguments,
            string outputPath,
            string processOutputPath,
            RecordingTaskKind taskKind,
            bool videoHasAudio,
            WasapiLoopbackAudioRecorder loopbackRecorder,
            string loopbackAudioPath,
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
                    CreateNoWindow = true,
                    StandardErrorEncoding = Encoding.Default
                },
                EnableRaisingEvents = true
            };

            ResetFfmpegLogState();
            process.ErrorDataReceived += HandleFfmpegErrorData;

            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!process.Start())
                {
                    throw new InvalidOperationException("无法启动 ffmpeg 进程。");
                }

                process.BeginErrorReadLine();
            }
            catch
            {
                process.ErrorDataReceived -= HandleFfmpegErrorData;
                process.Dispose();
                throw;
            }

            lock (_syncRoot)
            {
                _activeProcess = process;
                _activeTaskKind = taskKind;
                _activeOutputPath = outputPath;
                _activeProcessOutputPath = processOutputPath;
                _activeLoopbackRecorder = loopbackRecorder;
                _activeLoopbackAudioPath = loopbackAudioPath;
                _activeFfmpegPath = ffmpegPath;
                _activeVideoHasAudio = taskKind == RecordingTaskKind.Video && videoHasAudio;
                _startedAtUtc = DateTime.UtcNow;
            }

            await Task.CompletedTask.ConfigureAwait(false);
        }

        private void HandleFfmpegErrorData(object sender, DataReceivedEventArgs e)
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
                LogFfmpegInfoLine(e.Data);
            }
        }

        private void LogFfmpegInfoLine(string line)
        {
            var now = DateTime.UtcNow;
            var suppressedCount = 0;
            var shouldLog = false;

            lock (_ffmpegLogSyncRoot)
            {
                if (_ffmpegInfoLinesLogged < MaxImmediateFfmpegInfoLines)
                {
                    _ffmpegInfoLinesLogged++;
                    _lastFfmpegInfoLogUtc = now;
                    shouldLog = true;
                }
                else if (now - _lastFfmpegInfoLogUtc >= FfmpegInfoLogInterval)
                {
                    suppressedCount = _suppressedFfmpegInfoLines;
                    _suppressedFfmpegInfoLines = 0;
                    _lastFfmpegInfoLogUtc = now;
                    shouldLog = true;
                }
                else
                {
                    _suppressedFfmpegInfoLines++;
                }
            }

            if (!shouldLog)
            {
                return;
            }

            if (suppressedCount > 0)
            {
                AppLogService.Information("ffmpeg status ({SuppressedCount} lines suppressed): {Line}", suppressedCount, line);
                return;
            }

            AppLogService.Information("ffmpeg: {Line}", line);
        }

        private void ResetFfmpegLogState()
        {
            lock (_ffmpegLogSyncRoot)
            {
                _lastFfmpegInfoLogUtc = DateTime.MinValue;
                _ffmpegInfoLinesLogged = 0;
                _suppressedFfmpegInfoLines = 0;
            }
        }

        private void EnsureIdle()
        {
            lock (_syncRoot)
            {
                if ((_activeProcess != null && !_activeProcess.HasExited) || _activeLoopbackRecorder != null)
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
                _activeProcessOutputPath = string.Empty;
                _activeLoopbackAudioPath = string.Empty;
                _activeFfmpegPath = string.Empty;
                _activeLoopbackRecorder = null;
                _activeVideoHasAudio = false;
                _startedAtUtc = DateTime.UtcNow;
            }
        }

        private Process GetActiveProcess()
        {
            lock (_syncRoot)
            {
                return _activeProcess;
            }
        }

        private async Task<int?> WaitForActiveProcessEarlyExitCodeAsync(int timeoutMilliseconds)
        {
            var process = GetActiveProcess();
            var exited = await WaitForExitAsync(process, timeoutMilliseconds).ConfigureAwait(false);
            if (!exited)
            {
                return null;
            }

            try
            {
                return process?.ExitCode ?? -1;
            }
            catch
            {
                return -1;
            }
        }

        private void CleanupFailedStart()
        {
            Process process;
            WasapiLoopbackAudioRecorder loopbackRecorder;
            string processOutputPath;
            string loopbackAudioPath;
            lock (_syncRoot)
            {
                process = _activeProcess;
                loopbackRecorder = _activeLoopbackRecorder;
                processOutputPath = _activeProcessOutputPath;
                loopbackAudioPath = _activeLoopbackAudioPath;
                _activeProcess = null;
                _activeTaskKind = RecordingTaskKind.None;
                _activeOutputPath = string.Empty;
                _activeProcessOutputPath = string.Empty;
                _activeLoopbackAudioPath = string.Empty;
                _activeFfmpegPath = string.Empty;
                _activeLoopbackRecorder = null;
                _activeVideoHasAudio = false;
                _startedAtUtc = DateTime.UtcNow;
            }

            loopbackRecorder?.Stop();
            TryDeleteTempSiblingFile(processOutputPath, string.Empty);
            TryDeleteFile(loopbackAudioPath);

            if (process == null)
            {
                return;
            }

            process.ErrorDataReceived -= HandleFfmpegErrorData;
            try { process.CancelErrorRead(); } catch { }
            process.Dispose();
        }

        private static WasapiLoopbackAudioRecorder TryStartLoopbackRecorder(string outputPath, out string loopbackAudioPath)
        {
            loopbackAudioPath = BuildTempSiblingPath(outputPath, "audio", ".wav");
            try
            {
                var recorder = WasapiLoopbackAudioRecorder.Start(loopbackAudioPath);
                AppLogService.Information("WASAPI loopback recording started: {Path}, format={Format}", loopbackAudioPath, recorder.DeviceName);
                return recorder;
            }
            catch (Exception ex)
            {
                TryDeleteFile(loopbackAudioPath);
                AppLogService.Warning("WASAPI loopback recording unavailable: {Msg}", ex.Message);
                return null;
            }
        }

        private async Task<bool> MuxVideoAndLoopbackAudioAsync(string ffmpegPath, string videoPath, string audioPath, WasapiLoopbackAudioRecorder recorder, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(videoPath) || !File.Exists(videoPath))
            {
                return false;
            }

            if (!HasUsableLoopbackAudio(audioPath, recorder, out var audioStats))
            {
                File.Copy(videoPath, outputPath, true);
                AppLogService.Warning(
                    "Loopback audio is empty or silent; saved video without audio: {Path}, bytes={Bytes}, peak={Peak}, audibleSamples={Samples}",
                    outputPath,
                    audioStats.BytesRecorded,
                    audioStats.MaxSampleLevel,
                    audioStats.AudibleSampleCount);
                return false;
            }

            TryDeleteFile(outputPath);
            var arguments = string.Format(
                "-y -i \"{0}\" -i \"{1}\" -map 0:v:0 -map 1:a:0 -c:v copy -c:a aac -b:a 160k -shortest -movflags +faststart \"{2}\"",
                videoPath,
                audioPath,
                outputPath);
            await RunFfmpegUtilityAsync(ffmpegPath, arguments, TimeSpan.FromSeconds(60)).ConfigureAwait(false);
            return true;
        }

        private async Task SaveLoopbackAudioAsync(string ffmpegPath, string audioPath, WasapiLoopbackAudioRecorder recorder, string outputPath)
        {
            if (!HasUsableLoopbackAudio(audioPath, recorder, out var audioStats))
            {
                TryDeleteFile(outputPath);
                AppLogService.Warning(
                    "Loopback audio save cancelled because audio is empty or silent: {Path}, bytes={Bytes}, peak={Peak}, audibleSamples={Samples}",
                    outputPath,
                    audioStats.BytesRecorded,
                    audioStats.MaxSampleLevel,
                    audioStats.AudibleSampleCount);
                throw new InvalidOperationException(
                    "未检测到可听的系统声音，录音文件已取消保存。\n\n"
                    + "请确认电脑正在播放声音、播放器和系统音量未静音，并且当前默认播放设备就是正在出声的设备。");
            }

            var format = AudioRecordingFormat.FromPath(outputPath);
            TryDeleteFile(outputPath);
            if (!format.RequiresFfmpeg)
            {
                File.Move(audioPath, outputPath);
                return;
            }

            var codecArgs = format.FfmpegCodecArguments;
            var arguments = string.Format(
                "-y -i \"{0}\" {1} \"{2}\"",
                audioPath,
                codecArgs,
                outputPath);
            await RunFfmpegUtilityAsync(ffmpegPath, arguments, TimeSpan.FromSeconds(60)).ConfigureAwait(false);
        }

        private static bool HasUsableLoopbackAudio(string audioPath, WasapiLoopbackAudioRecorder recorder, out AudioLevelStats audioStats)
        {
            if (string.IsNullOrWhiteSpace(audioPath) || !File.Exists(audioPath) || GetFileSize(audioPath) <= 44)
            {
                audioStats = new AudioLevelStats(0, 0d, 0);
                return false;
            }

            var liveStats = recorder?.GetLiveAudioLevelStats() ?? new AudioLevelStats(0, 0d, 0);
            if (liveStats.HasAudibleAudio)
            {
                audioStats = liveStats;
                return true;
            }

            audioStats = WasapiLoopbackAudioRecorder.ScanWaveFile(audioPath);
            return audioStats.HasAudibleAudio;
        }

        private async Task RunFfmpegUtilityAsync(string ffmpegPath, string arguments, TimeSpan timeout)
        {
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = arguments,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardErrorEncoding = Encoding.Default,
                    StandardOutputEncoding = Encoding.Default
                };

                if (!process.Start())
                {
                    throw new InvalidOperationException("无法启动 ffmpeg 合成进程。");
                }

                var stderrTask = process.StandardError.ReadToEndAsync();
                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var exited = await WaitForExitAsync(process, (int)Math.Max(1000, timeout.TotalMilliseconds)).ConfigureAwait(false);
                if (!exited)
                {
                    process.Kill();
                    await WaitForExitAsync(process, 3000).ConfigureAwait(false);
                    throw new TimeoutException("ffmpeg 合成超时。");
                }

                var stderr = await stderrTask.ConfigureAwait(false);
                await stdoutTask.ConfigureAwait(false);
                if (process.ExitCode != 0)
                {
                    AppLogService.Error("ffmpeg utility failed ({ExitCode}): {Args}; {Error}", process.ExitCode, arguments, stderr);
                    throw new InvalidOperationException("ffmpeg 合成音频失败，请查看日志。");
                }
            }
        }

        private static long GetFileSize(string path)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(path) && File.Exists(path) ? new FileInfo(path).Length : 0L;
            }
            catch
            {
                return 0L;
            }
        }

        private static string BuildTempSiblingPath(string outputPath, string marker, string extension)
        {
            var folder = Path.GetDirectoryName(outputPath) ?? AppDomain.CurrentDomain.BaseDirectory;
            var name = Path.GetFileNameWithoutExtension(outputPath);
            return Path.Combine(folder, name + "." + marker + ".tmp" + extension);
        }

        private static void TryDeleteEmptyOutput(string outputPath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath) && new FileInfo(outputPath).Length == 0)
                {
                    File.Delete(outputPath);
                }
            }
            catch
            {
            }
        }

        private static void TryDeleteFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
            }
        }

        private static void TryDeleteTempSiblingFile(string filePath, string finalOutputPath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(finalOutputPath)
                && string.Equals(Path.GetFullPath(filePath), Path.GetFullPath(finalOutputPath), StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            TryDeleteFile(filePath);
        }

        private static void ValidateRecordingRegion(RecordingRegion region)
        {
            if (region.Width <= 0 || region.Height <= 0)
            {
                throw new InvalidOperationException("录制区域无效，请重新调整区域窗口大小后再试。");
            }
        }

        private static RecordingRegion NormalizeVideoRecordingRegion(RecordingRegion region)
        {
            var normalized = region;
            if (normalized.Width % 2 != 0)
            {
                normalized.Width -= 1;
            }

            if (normalized.Height % 2 != 0)
            {
                normalized.Height -= 1;
            }

            if (normalized.Width <= 0 || normalized.Height <= 0)
            {
                throw new InvalidOperationException("录制区域过小，请重新调整区域窗口大小后再试。");
            }

            if (normalized.Width != region.Width || normalized.Height != region.Height)
            {
                AppLogService.Information(
                    "Normalized video recording region from {OriginalWidth}x{OriginalHeight} to {Width}x{Height} for H.264 encoding.",
                    region.Width,
                    region.Height,
                    normalized.Width,
                    normalized.Height);
            }

            return normalized;
        }

        private static void ValidateOutputPath(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new InvalidOperationException("输出路径不能为空。");
            }
        }

        private static string BuildVideoRecordingArguments(RecordingRegion region, string outputPath, RecordingOptions options)
        {
            var args = new StringBuilder();
            args.AppendFormat("-y -f gdigrab -framerate {0} ", options.FrameRate);
            args.AppendFormat("-offset_x {0} -offset_y {1} ", region.X, region.Y);
            args.AppendFormat("-video_size {0}x{1} -i desktop ", region.Width, region.Height);

            args.AppendFormat("-c:v libx264 -preset {0} -crf {1} -pix_fmt yuv420p ", options.Preset, options.Crf);
            args.AppendFormat("-movflags +faststart \"{0}\"", outputPath);
            return args.ToString();
        }

        private static string BuildGifRecordingArguments(RecordingRegion region, string outputPath, RecordingOptions options)
        {
            var fps = Math.Min(15, Math.Max(8, options.FrameRate));
            var args = new StringBuilder();
            args.AppendFormat("-y -f gdigrab -framerate {0} ", fps);
            args.AppendFormat("-offset_x {0} -offset_y {1} ", region.X, region.Y);
            args.AppendFormat("-video_size {0}x{1} -i desktop ", region.Width, region.Height);
            args.AppendFormat("-vf \"fps={0},scale=trunc(iw/2)*2:trunc(ih/2)*2:flags=lanczos,split[s0][s1];[s0]palettegen=stats_mode=diff[p];[s1][p]paletteuse=dither=bayer:bayer_scale=5\" ", fps);
            args.AppendFormat("\"{0}\"", outputPath);
            return args.ToString();
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

    public sealed class RecordingOptions
    {
        public static RecordingOptions Default => new RecordingOptions
        {
            FrameRate = 30,
            Crf = 23,
            Preset = "veryfast"
        };

        public int FrameRate { get; set; }
        public int Crf { get; set; }
        public string Preset { get; set; }

        public static RecordingOptions Normalize(RecordingOptions options)
        {
            var value = options ?? Default;
            var frameRate = value.FrameRate <= 0 ? 30 : Math.Min(60, Math.Max(10, value.FrameRate));
            var crf = value.Crf <= 0 ? 23 : Math.Min(30, Math.Max(18, value.Crf));
            var preset = string.IsNullOrWhiteSpace(value.Preset) ? "veryfast" : value.Preset.Trim();
            return new RecordingOptions
            {
                FrameRate = frameRate,
                Crf = crf,
                Preset = preset
            };
        }
    }

    public sealed class RecordingStopResult
    {
        public bool TimedOut { get; set; }
        public string OutputPath { get; set; } = string.Empty;
        public long DurationSeconds { get; set; }
        public long FileSizeBytes { get; set; }
        public bool VideoHasAudio { get; set; }
    }

    public sealed class RecordingAudioLevelSnapshot
    {
        public double CurrentSampleLevel { get; set; }
        public double MaxSampleLevel { get; set; }
        public long BytesRecorded { get; set; }
        public bool HasAudibleAudio { get; set; }
    }

    public sealed class AudioRecordingFormat
    {
        private static readonly AudioRecordingFormat WavFormat = new AudioRecordingFormat("wav", false, string.Empty);
        private static readonly AudioRecordingFormat M4aFormat = new AudioRecordingFormat("m4a", true, "-c:a aac -b:a 160k -ar 44100 -ac 2");
        private static readonly AudioRecordingFormat Mp3Format = new AudioRecordingFormat("mp3", true, "-c:a libmp3lame -b:a 192k -ar 44100 -ac 2");

        private AudioRecordingFormat(string extension, bool requiresFfmpeg, string ffmpegCodecArguments)
        {
            Extension = extension;
            RequiresFfmpeg = requiresFfmpeg;
            FfmpegCodecArguments = ffmpegCodecArguments;
        }

        public string Extension { get; }

        public bool RequiresFfmpeg { get; }

        public string FfmpegCodecArguments { get; }

        public static AudioRecordingFormat FromPath(string outputPath)
        {
            return FromExtension(Path.GetExtension(outputPath));
        }

        public static AudioRecordingFormat FromExtension(string extension)
        {
            var normalized = NormalizeExtension(extension);
            if (string.Equals(normalized, "m4a", StringComparison.OrdinalIgnoreCase))
            {
                return M4aFormat;
            }

            if (string.Equals(normalized, "mp3", StringComparison.OrdinalIgnoreCase))
            {
                return Mp3Format;
            }

            return WavFormat;
        }

        public static string NormalizeExtension(string extension)
        {
            var value = (extension ?? string.Empty).Trim().TrimStart('.').ToLowerInvariant();
            return value == "m4a" || value == "mp3" || value == "wav" ? value : "wav";
        }
    }
}
