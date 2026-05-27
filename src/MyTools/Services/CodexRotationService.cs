using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Hardcodet.Wpf.TaskbarNotification;
using MyTools.ViewModels;

namespace MyTools.Services
{
    public class CodexRotationSettings
    {
        public bool IsEnabled { get; set; } = false;
        public bool NotifyOnSwitch { get; set; } = true;
    }

    public class CodexRotationResult
    {
        public bool Success { get; set; }
        public string FromProfile { get; set; }
        public string ToProfile { get; set; }
        public string Message { get; set; }
    }

    public static class CodexRotationService
    {
        public static List<CodexProfileItem> GetRotatingPool()
        {
            var profiles = CodexProfileLibraryService.GetCachedProfiles();
            return profiles.items
                .Where(p => p != null && p.EnableRotation && p.Status != CodexProfileLibraryService.StatusExpired)
                .OrderBy(p => p.RotationPriority)
                .ThenBy(p => p.LastAppliedAt ?? DateTime.MinValue)
                .ToList();
        }

        public static string GetNextSwitchPreview(CodexProfileItem current)
        {
            if (current == null)
            {
                return "无可用轮换目标";
            }

            var pool = GetRotatingPool();
            var candidates = pool.Where(p => p.DisplayName != current.DisplayName).ToList();
            if (candidates.Count == 0)
            {
                return "无可用轮换目标";
            }

            var next = candidates.First();
            return $"当前：{current.DisplayName} → 切换至：{next.DisplayName}";
        }

        public static async Task<CodexRotationResult> RotateToNextAsync(
            CodexProfileItem current,
            bool notifyOnSwitch,
            CancellationToken ct)
        {
            if (current == null)
            {
                return new CodexRotationResult { Success = false, Message = "未指定当前账号" };
            }

            var pool = GetRotatingPool();
            var candidates = pool.Where(p => p.DisplayName != current.DisplayName).ToList();

            if (candidates.Count == 0)
            {
                return new CodexRotationResult { Success = false, Message = "没有其他已加入轮换的账号" };
            }

            var next = candidates.First();
            string backupPath = string.Empty;

            try
            {
                backupPath = await CodexProfileLibraryService.BackupCurrentCodexFolderAsync(current.DisplayName, ct).ConfigureAwait(false);

                var authBytes = CodexConfigProfileService.UnprotectBytesFromBase64(next.ProtectedAuthJsonBase64);
                var configBytes = CodexConfigProfileService.UnprotectBytesFromBase64(next.ProtectedConfigTomlBase64);
                await CodexConfigProfileService.ApplyAsync(configBytes, authBytes, ct).ConfigureAwait(false);

                var active = new CodexActiveFile
                {
                    ActiveDisplayName = next.DisplayName,
                    SwitchedAtUtc = DateTime.UtcNow
                };
                await CodexProfileLibraryService.SaveActiveAsync(active, ct).ConfigureAwait(false);

                next.LastAppliedAt = DateTime.Now;
                await CodexProfileLibraryService.SaveAsync(CodexProfileLibraryService.GetCachedProfiles(), ct).ConfigureAwait(false);

                if (notifyOnSwitch)
                {
                    NotifySwitch(current.DisplayName, next.DisplayName);
                }

                AppLogService.Information("Codex 账号轮换完成: {From} -> {To}", current.DisplayName, next.DisplayName);

                return new CodexRotationResult
                {
                    Success = true,
                    FromProfile = current.DisplayName,
                    ToProfile = next.DisplayName,
                    Message = $"已写入 {next.DisplayName}，请重启 Codex App 后使用"
                };
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Codex 账号轮换失败: {From} → {To}, 错误类型: {ErrorType}",
                    current.DisplayName, next?.DisplayName ?? "(未知)", ex.GetType().Name);

                if (!string.IsNullOrEmpty(backupPath))
                {
                    try
                    {
                        await CodexProfileLibraryService.RestoreLatestBackupAsync(ct).ConfigureAwait(false);
                        AppLogService.Information("Codex 账号轮换失败，已回滚备份: {BackupPath}", backupPath);
                    }
                    catch (Exception rollbackEx)
                    {
                        AppLogService.Error(rollbackEx, "Codex 账号回滚失败，错误类型: {ErrorType}", rollbackEx.GetType().Name);
                    }
                }

                return new CodexRotationResult
                {
                    Success = false,
                    FromProfile = current.DisplayName,
                    ToProfile = next?.DisplayName,
                    Message = $"轮换失败：{ex.Message}"
                };
            }
        }

        private static void NotifySwitch(string fromProfile, string toProfile)
        {
            try
            {
                var mainWindow = System.Windows.Application.Current?.MainWindow;
                if (mainWindow == null) return;

                var taskbarIcon = mainWindow.TryFindResource("TrayIcon") as Hardcodet.Wpf.TaskbarNotification.TaskbarIcon;
                if (taskbarIcon == null) return;

                taskbarIcon.ShowBalloonTip(
                    "Codex 配置已写入",
                    $"{fromProfile} → {toProfile}，重启 Codex App 后生效",
                    BalloonIcon.Info);
            }
            catch
            {
            }
        }
    }
}
