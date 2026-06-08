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
        public bool RelayTestExecuted { get; set; }
        public bool RelayTestSucceeded { get; set; }
        public DateTime? RelayTestedAt { get; set; }
        public string RelayTestStatus { get; set; }
        public string RelayTestMessage { get; set; }
        public bool RequiresCodexRestart { get; set; } = true;
        public bool UsedHotTokenRefresh { get; set; }
        public bool UsedLocalRelaySwitch { get; set; }
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
            var relayTestExecuted = false;
            var relayTestSucceeded = false;
            DateTime? relayTestedAt = null;
            string relayTestStatus = null;
            string relayTestMessage = null;

            try
            {
                var authBytes = CodexConfigProfileService.UnprotectBytesFromBase64(
                    SelectProtectedContent(next.ProtectedAuthJsonBase64, next.AuthJsonContentProtected));
                var configBytes = CodexConfigProfileService.UnprotectBytesFromBase64(
                    SelectProtectedContent(next.ProtectedConfigTomlBase64, next.ConfigTomlContentProtected));
                if (configBytes == null || configBytes.Length == 0 || authBytes == null || authBytes.Length == 0)
                {
                    relayTestExecuted = true;
                    relayTestSucceeded = false;
                    relayTestedAt = DateTime.Now;
                    relayTestStatus = CodexProfileItem.RelayStatusFailed;
                    relayTestMessage = "目标档案缺少 config.toml 或 auth.json。";
                    return new CodexRotationResult
                    {
                        Success = false,
                        FromProfile = current.DisplayName,
                        ToProfile = next.DisplayName,
                        Message = "轮换前中转测试未通过：" + relayTestMessage,
                        RelayTestExecuted = relayTestExecuted,
                        RelayTestSucceeded = relayTestSucceeded,
                        RelayTestedAt = relayTestedAt,
                        RelayTestStatus = relayTestStatus,
                        RelayTestMessage = relayTestMessage
                    };
                }

                var relayTest = await CodexRelayTestService.TestAsync(configBytes, authBytes, ct).ConfigureAwait(false);
                relayTestExecuted = true;
                relayTestSucceeded = relayTest.Success;
                relayTestedAt = DateTime.Now;
                relayTestStatus = relayTest.Success ? CodexProfileItem.RelayStatusOk : CodexProfileItem.RelayStatusFailed;
                relayTestMessage = LimitRelayTestMessage(relayTest.Message);
                if (!relayTest.Success)
                {
                    AppLogService.Warning(
                        "Codex 账号轮换前中转测试未通过: {To}, {Message}",
                        next.DisplayName,
                        relayTestMessage);

                    return new CodexRotationResult
                    {
                        Success = false,
                        FromProfile = current.DisplayName,
                        ToProfile = next.DisplayName,
                        Message = "轮换前中转测试未通过：" + relayTestMessage,
                        RelayTestExecuted = relayTestExecuted,
                        RelayTestSucceeded = relayTestSucceeded,
                        RelayTestedAt = relayTestedAt,
                        RelayTestStatus = relayTestStatus,
                        RelayTestMessage = relayTestMessage
                    };
                }

                var localRelay = await CodexLocalRelayService.TryApplyProfileAsync(configBytes, authBytes, next.DisplayName, ct).ConfigureAwait(false);
                if (localRelay.Success)
                {
                    var relayActive = new CodexActiveFile
                    {
                        ActiveDisplayName = next.DisplayName,
                        SwitchedAtUtc = DateTime.UtcNow
                    };
                    await CodexProfileLibraryService.SaveActiveAsync(relayActive, ct).ConfigureAwait(false);

                    next.LastAppliedAt = DateTime.Now;
                    next.RelayTestStatus = CodexProfileItem.RelayStatusOk;
                    next.RelayTestedAt = relayTestedAt;
                    next.RelayTestMessage = relayTestMessage;
                    await CodexProfileLibraryService.SaveAsync(CodexProfileLibraryService.GetCachedProfiles(), ct).ConfigureAwait(false);

                    if (notifyOnSwitch)
                    {
                        NotifySwitch(current.DisplayName, next.DisplayName, "local-relay");
                    }

                    AppLogService.Information("Codex 本地中转轮换完成: {From} -> {To}", current.DisplayName, next.DisplayName);

                    return new CodexRotationResult
                    {
                        Success = true,
                        FromProfile = current.DisplayName,
                        ToProfile = next.DisplayName,
                        Message = localRelay.Message,
                        RelayTestExecuted = relayTestExecuted,
                        RelayTestSucceeded = relayTestSucceeded,
                        RelayTestedAt = relayTestedAt,
                        RelayTestStatus = relayTestStatus,
                        RelayTestMessage = relayTestMessage,
                        RequiresCodexRestart = localRelay.RequiresCodexRestart,
                        UsedLocalRelaySwitch = true
                    };
                }

                if (localRelay.LocalRelayEnabled && !localRelay.AllowFullConfigSwitch)
                {
                    return new CodexRotationResult
                    {
                        Success = false,
                        FromProfile = current.DisplayName,
                        ToProfile = next.DisplayName,
                        Message = localRelay.Message,
                        RelayTestExecuted = relayTestExecuted,
                        RelayTestSucceeded = relayTestSucceeded,
                        RelayTestedAt = relayTestedAt,
                        RelayTestStatus = relayTestStatus,
                        RelayTestMessage = relayTestMessage,
                        RequiresCodexRestart = false
                    };
                }

                var hotToken = await CodexHotTokenService.TryApplyProfileTokenAsync(configBytes, authBytes, ct).ConfigureAwait(false);
                if (hotToken.Success)
                {
                    var hotActive = new CodexActiveFile
                    {
                        ActiveDisplayName = next.DisplayName,
                        SwitchedAtUtc = DateTime.UtcNow
                    };
                    await CodexProfileLibraryService.SaveActiveAsync(hotActive, ct).ConfigureAwait(false);

                    next.LastAppliedAt = DateTime.Now;
                    next.RelayTestStatus = CodexProfileItem.RelayStatusOk;
                    next.RelayTestedAt = relayTestedAt;
                    next.RelayTestMessage = relayTestMessage;
                    await CodexProfileLibraryService.SaveAsync(CodexProfileLibraryService.GetCachedProfiles(), ct).ConfigureAwait(false);

                    if (notifyOnSwitch)
                    {
                        NotifySwitch(current.DisplayName, next.DisplayName, "hot-token");
                    }

                    AppLogService.Information("Codex 账号热轮换完成: {From} -> {To}", current.DisplayName, next.DisplayName);

                    return new CodexRotationResult
                    {
                        Success = true,
                        FromProfile = current.DisplayName,
                        ToProfile = next.DisplayName,
                        Message = hotToken.Message,
                        RelayTestExecuted = relayTestExecuted,
                        RelayTestSucceeded = relayTestSucceeded,
                        RelayTestedAt = relayTestedAt,
                        RelayTestStatus = relayTestStatus,
                        RelayTestMessage = relayTestMessage,
                        RequiresCodexRestart = false,
                        UsedHotTokenRefresh = true
                    };
                }

                if (hotToken.HotModeEnabled && !hotToken.AllowFullConfigSwitch)
                {
                    return new CodexRotationResult
                    {
                        Success = false,
                        FromProfile = current.DisplayName,
                        ToProfile = next.DisplayName,
                        Message = hotToken.Message,
                        RelayTestExecuted = relayTestExecuted,
                        RelayTestSucceeded = relayTestSucceeded,
                        RelayTestedAt = relayTestedAt,
                        RelayTestStatus = relayTestStatus,
                        RelayTestMessage = relayTestMessage,
                        RequiresCodexRestart = false
                    };
                }

                backupPath = await CodexProfileLibraryService.BackupCurrentCodexFolderAsync(current.DisplayName, ct).ConfigureAwait(false);
                await CodexConfigProfileService.ApplyAsync(configBytes, authBytes, ct).ConfigureAwait(false);

                var active = new CodexActiveFile
                {
                    ActiveDisplayName = next.DisplayName,
                    SwitchedAtUtc = DateTime.UtcNow
                };
                await CodexProfileLibraryService.SaveActiveAsync(active, ct).ConfigureAwait(false);

                next.LastAppliedAt = DateTime.Now;
                next.RelayTestStatus = CodexProfileItem.RelayStatusOk;
                next.RelayTestedAt = relayTestedAt;
                next.RelayTestMessage = relayTestMessage;
                await CodexProfileLibraryService.SaveAsync(CodexProfileLibraryService.GetCachedProfiles(), ct).ConfigureAwait(false);

                if (notifyOnSwitch)
                {
                    NotifySwitch(current.DisplayName, next.DisplayName, "full-config");
                }

                AppLogService.Information("Codex 账号轮换完成: {From} -> {To}", current.DisplayName, next.DisplayName);

                return new CodexRotationResult
                {
                    Success = true,
                    FromProfile = current.DisplayName,
                    ToProfile = next.DisplayName,
                    Message = $"已写入 {next.DisplayName}，请重启 Codex App 后使用",
                    RelayTestExecuted = relayTestExecuted,
                    RelayTestSucceeded = relayTestSucceeded,
                    RelayTestedAt = relayTestedAt,
                    RelayTestStatus = relayTestStatus,
                    RelayTestMessage = relayTestMessage,
                    RequiresCodexRestart = true
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
                    Message = $"轮换失败：{ex.Message}",
                    RelayTestExecuted = relayTestExecuted,
                    RelayTestSucceeded = relayTestSucceeded,
                    RelayTestedAt = relayTestedAt,
                    RelayTestStatus = relayTestStatus,
                    RelayTestMessage = relayTestMessage
                };
            }
        }

        private static string SelectProtectedContent(string primary, string fallback)
        {
            return string.IsNullOrWhiteSpace(primary) ? fallback : primary;
        }

        private static string LimitRelayTestMessage(string value)
        {
            var text = string.IsNullOrWhiteSpace(value) ? "中转测试失败。" : value.Trim();
            return text.Length <= 160 ? text : text.Substring(0, 160) + "...";
        }

        private static void NotifySwitch(string fromProfile, string toProfile, string mode)
        {
            try
            {
                var mainWindow = System.Windows.Application.Current?.MainWindow;
                if (mainWindow == null) return;

                var taskbarIcon = mainWindow.TryFindResource("TrayIcon") as Hardcodet.Wpf.TaskbarNotification.TaskbarIcon;
                if (taskbarIcon == null) return;

                var localRelay = string.Equals(mode, "local-relay", StringComparison.OrdinalIgnoreCase);
                var hotToken = string.Equals(mode, "hot-token", StringComparison.OrdinalIgnoreCase);
                taskbarIcon.ShowBalloonTip(
                    localRelay ? "Codex 本地中转已切换" : (hotToken ? "Codex 热轮换已更新" : "Codex 配置已写入"),
                    localRelay
                        ? $"{fromProfile} → {toProfile}，下一次请求使用新上游"
                        : (hotToken
                            ? $"{fromProfile} → {toProfile}，等待 Codex 自动刷新 token"
                            : $"{fromProfile} → {toProfile}，重启 Codex App 后生效"),
                    BalloonIcon.Info);
            }
            catch
            {
            }
        }
    }
}
