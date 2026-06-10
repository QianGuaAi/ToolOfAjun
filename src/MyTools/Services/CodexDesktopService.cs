using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public class CodexDesktopRestartResult
    {
        public bool Success { get; set; }
        public bool WasRunning { get; set; }
        public string Message { get; set; }
    }

    public static class CodexDesktopService
    {
        private const string CodexPackagePrefix = "OpenAI.Codex_";
        private const string FallbackAppUserModelId = "OpenAI.Codex_2p2nqsd0c76g0!App";
        private static readonly TimeSpan GracefulCloseTimeout = TimeSpan.FromSeconds(6);
        private static readonly TimeSpan ForceCloseTimeout = TimeSpan.FromSeconds(8);
        private static readonly TimeSpan LaunchDetectTimeout = TimeSpan.FromSeconds(10);
        private static readonly TimeSpan NoProcessStableDuration = TimeSpan.FromMilliseconds(800);

        public static async Task<CodexDesktopRestartResult> RestartAsync(CancellationToken cancellationToken)
        {
            var codexProcesses = GetCodexDesktopProcesses();
            var wasRunning = codexProcesses.Count > 0;
            try
            {
                var killedCount = 0;
                if (wasRunning)
                {
                    foreach (var process in codexProcesses.Where(HasMainWindow))
                    {
                        TryCloseMainWindow(process);
                    }

                    DisposeProcesses(codexProcesses);
                    codexProcesses = null;

                    var closed = await WaitForNoCodexProcessesAsync(GracefulCloseTimeout, cancellationToken);
                    if (!closed)
                    {
                        var remainingProcesses = GetCodexDesktopProcesses();
                        try
                        {
                            killedCount = TryKillProcesses(remainingProcesses);
                        }
                        finally
                        {
                            DisposeProcesses(remainingProcesses);
                        }

                        closed = await WaitForNoCodexProcessesAsync(ForceCloseTimeout, cancellationToken);
                        if (!closed)
                        {
                            return new CodexDesktopRestartResult
                            {
                                Success = false,
                                WasRunning = true,
                                Message = "Codex App 仍有进程未退出，重启未完成。请在任务管理器结束 Codex 后再重新打开。"
                            };
                        }
                    }

                    await Task.Delay(800, cancellationToken);
                }

                StartCodexDesktop();
                var started = await WaitForCodexDesktopStartedAsync(LaunchDetectTimeout, cancellationToken);
                if (!started)
                {
                    return new CodexDesktopRestartResult
                    {
                        Success = false,
                        WasRunning = wasRunning,
                        Message = "已发出启动请求，但未检测到 Codex App 进程。请从开始菜单手动打开 Codex。"
                    };
                }

                return new CodexDesktopRestartResult
                {
                    Success = true,
                    WasRunning = wasRunning,
                    Message = wasRunning
                        ? (killedCount > 0
                            ? $"已关闭 {killedCount} 个遗留 Codex 进程并重新打开 Codex App。对话历史通常会保留。"
                            : "已重新打开 Codex App。对话历史通常会保留。")
                        : "已打开 Codex App。"
                };
            }
            finally
            {
                DisposeProcesses(codexProcesses);
            }
        }

        private static List<Process> GetCodexDesktopProcesses()
        {
            var result = new List<Process>();
            foreach (var process in Process.GetProcesses())
            {
                try
                {
                    if (!process.HasExited && IsCodexDesktopProcess(process))
                    {
                        result.Add(process);
                    }
                    else
                    {
                        process.Dispose();
                    }
                }
                catch
                {
                    process.Dispose();
                }
            }

            return result;
        }

        private static bool IsCodexDesktopProcess(Process process)
        {
            var processName = process.ProcessName ?? string.Empty;
            if (!string.Equals(processName, "Codex", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(processName, "codex", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var path = GetProcessPath(process);
            if (string.IsNullOrWhiteSpace(path))
            {
                return HasMainWindow(process) && IsCodexWindowTitle(process.MainWindowTitle);
            }

            return path.IndexOf("\\WindowsApps\\OpenAI.Codex_", StringComparison.OrdinalIgnoreCase) >= 0
                || path.IndexOf("\\AppData\\Local\\OpenAI\\Codex\\", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string GetProcessPath(Process process)
        {
            try
            {
                return process.MainModule?.FileName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool HasMainWindow(Process process)
        {
            try
            {
                return process.MainWindowHandle != IntPtr.Zero;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsCodexWindowTitle(string title)
        {
            return !string.IsNullOrWhiteSpace(title)
                && title.IndexOf("Codex", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void TryCloseMainWindow(Process process)
        {
            try
            {
                process.CloseMainWindow();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Failed to request Codex App close: {ErrorType}", ex.GetType().Name);
            }
        }

        private static async Task<bool> WaitForProcessesToExitAsync(
            IEnumerable<Process> processes,
            TimeSpan timeout,
            CancellationToken cancellationToken)
        {
            var processList = processes.ToList();
            var deadline = DateTime.UtcNow.Add(timeout);
            while (DateTime.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (processList.All(HasExited))
                {
                    return true;
                }

                await Task.Delay(250, cancellationToken);
            }

            return processList.All(HasExited);
        }

        private static async Task<bool> WaitForNoCodexProcessesAsync(
            TimeSpan timeout,
            CancellationToken cancellationToken)
        {
            var deadline = DateTime.UtcNow.Add(timeout);
            DateTime? noProcessSince = null;
            while (DateTime.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var processes = GetCodexDesktopProcesses();
                try
                {
                    if (processes.Count == 0)
                    {
                        if (!noProcessSince.HasValue)
                        {
                            noProcessSince = DateTime.UtcNow;
                        }

                        if (DateTime.UtcNow - noProcessSince.Value >= NoProcessStableDuration)
                        {
                            return true;
                        }
                    }
                    else
                    {
                        noProcessSince = null;
                    }
                }
                finally
                {
                    DisposeProcesses(processes);
                }

                await Task.Delay(250, cancellationToken);
            }

            return false;
        }

        private static async Task<bool> WaitForCodexDesktopStartedAsync(
            TimeSpan timeout,
            CancellationToken cancellationToken)
        {
            var deadline = DateTime.UtcNow.Add(timeout);
            while (DateTime.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var processes = GetCodexDesktopProcesses();
                try
                {
                    if (processes.Count > 0)
                    {
                        return true;
                    }
                }
                finally
                {
                    DisposeProcesses(processes);
                }

                await Task.Delay(300, cancellationToken);
            }

            return false;
        }

        private static bool HasExited(Process process)
        {
            try
            {
                process.Refresh();
                return process.HasExited;
            }
            catch
            {
                return true;
            }
        }

        private static void StartCodexDesktop()
        {
            var appUserModelId = ResolveCodexAppUserModelId();
            var startInfo = new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = "\"shell:AppsFolder\\" + appUserModelId + "\"",
                CreateNoWindow = true,
                UseShellExecute = true
            };
            var process = Process.Start(startInfo);
            if (process != null)
            {
                process.Dispose();
            }
        }

        private static string ResolveCodexAppUserModelId()
        {
            try
            {
                var packageRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Packages");
                if (!Directory.Exists(packageRoot))
                {
                    return FallbackAppUserModelId;
                }

                var packageFolder = Directory.GetDirectories(packageRoot, CodexPackagePrefix + "*")
                    .OrderByDescending(Directory.GetLastWriteTimeUtc)
                    .FirstOrDefault();
                if (string.IsNullOrWhiteSpace(packageFolder))
                {
                    return FallbackAppUserModelId;
                }

                return Path.GetFileName(packageFolder) + "!App";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Failed to resolve Codex AppUserModelId: {ErrorType}", ex.GetType().Name);
                return FallbackAppUserModelId;
            }
        }

        private static int TryKillProcesses(IEnumerable<Process> processes)
        {
            var killedCount = 0;
            foreach (var process in processes)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill();
                        killedCount++;
                    }
                }
                catch (Exception ex)
                {
                    AppLogService.Warning(
                        "Failed to terminate Codex App process {ProcessId}: {ErrorType}",
                        SafeProcessId(process),
                        ex.GetType().Name);
                }
            }

            return killedCount;
        }

        private static int SafeProcessId(Process process)
        {
            try
            {
                return process.Id;
            }
            catch
            {
                return 0;
            }
        }

        private static void DisposeProcesses(IEnumerable<Process> processes)
        {
            if (processes == null)
            {
                return;
            }

            foreach (var process in processes)
            {
                try
                {
                    process.Dispose();
                }
                catch
                {
                }
            }
        }
    }
}
