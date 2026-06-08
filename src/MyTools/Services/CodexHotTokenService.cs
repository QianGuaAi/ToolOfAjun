using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public static class CodexHotTokenService
    {
        private const string ProviderId = "mytools_hot_newapi";
        private const int DefaultRefreshIntervalMs = 5000;

        private static string HotTokenFolderPath => Path.Combine(CodexProfileLibraryService.RootFolderPath, "HotToken");
        private static string SettingsPath => Path.Combine(HotTokenFolderPath, "hot-token.json");
        private static string TokenPath => Path.Combine(HotTokenFolderPath, "token.dpapi");
        private static string ScriptPath => Path.Combine(HotTokenFolderPath, "ReadCodexHotToken.ps1");

        public static async Task<CodexHotTokenSetupResult> EnableFromCurrentCodexAsync(CancellationToken ct)
        {
            var current = await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
            var runtime = CodexRelayTestService.InspectRuntime(current.ConfigTomlBytes, current.AuthJsonBytes);
            ValidateRuntime(runtime);

            Directory.CreateDirectory(HotTokenFolderPath);
            var settings = new CodexHotTokenSettings
            {
                Enabled = true,
                ProviderId = ProviderId,
                BaseUrl = runtime.BaseUrl,
                TokenPath = TokenPath,
                ScriptPath = ScriptPath,
                RefreshIntervalMs = DefaultRefreshIntervalMs,
                EnabledAtUtc = DateTime.UtcNow
            };

            var backupPath = await CodexProfileLibraryService
                .BackupCurrentCodexFolderAsync("hot-token-before-enable", ct)
                .ConfigureAwait(false);

            await WriteTokenAsync(runtime.Token, settings, ct).ConfigureAwait(false);
            await WriteScriptAsync(settings, ct).ConfigureAwait(false);

            var configText = Encoding.UTF8.GetString(current.ConfigTomlBytes ?? new byte[0]);
            var hotConfig = BuildHotConfig(configText, settings, runtime.WireApi);
            await CodexConfigProfileService
                .ApplyAsync(Encoding.UTF8.GetBytes(hotConfig), current.AuthJsonBytes, ct)
                .ConfigureAwait(false);

            await SaveSettingsAsync(settings, ct).ConfigureAwait(false);
            AppLogService.Information("Codex hot token mode enabled for base URL host: {Host}", SafeHost(runtime.BaseUrl));

            return new CodexHotTokenSetupResult
            {
                Success = true,
                RequiresRestart = true,
                BaseUrl = runtime.BaseUrl,
                BackupPath = backupPath,
                Message = "已启用 Codex 热轮换。请重启一次 Codex App 让 auth.command 生效；之后同一 base_url 的轮换只更新 token 文件。"
            };
        }

        public static async Task<CodexHotTokenApplyResult> TryApplyProfileTokenAsync(
            byte[] configTomlBytes,
            byte[] authJsonBytes,
            CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (settings == null || !settings.Enabled)
            {
                return new CodexHotTokenApplyResult
                {
                    HotModeEnabled = false,
                    AllowFullConfigSwitch = true,
                    Message = "未启用 Codex 热轮换。"
                };
            }

            var runtime = CodexRelayTestService.InspectRuntime(configTomlBytes, authJsonBytes);
            if (string.IsNullOrWhiteSpace(runtime.BaseUrl))
            {
                return Stop("目标档案缺少 base_url，无法热轮换。");
            }

            if (!SameBaseUrl(settings.BaseUrl, runtime.BaseUrl))
            {
                return Stop("目标档案 base_url 与热轮换 base_url 不一致，已停止以避免覆盖热轮换配置。");
            }

            if (!runtime.HasToken)
            {
                return Stop("目标档案未解析到可用于热轮换的 API key 或 token。");
            }

            await WriteTokenAsync(runtime.Token, settings, ct).ConfigureAwait(false);
            await WriteScriptAsync(settings, ct).ConfigureAwait(false);
            AppLogService.Information("Codex hot token updated for base URL host: {Host}", SafeHost(runtime.BaseUrl));
            return new CodexHotTokenApplyResult
            {
                HotModeEnabled = true,
                Success = true,
                AllowFullConfigSwitch = false,
                BaseUrl = runtime.BaseUrl,
                Message = "已更新热轮换 token，Codex App 将在刷新间隔内读取新 token。"
            };
        }

        private static CodexHotTokenApplyResult Stop(string message)
        {
            return new CodexHotTokenApplyResult
            {
                HotModeEnabled = true,
                Success = false,
                AllowFullConfigSwitch = false,
                Message = message
            };
        }

        private static void ValidateRuntime(CodexRelayRuntimeInfo runtime)
        {
            if (runtime == null || string.IsNullOrWhiteSpace(runtime.BaseUrl))
            {
                throw new InvalidOperationException("当前 Codex 配置缺少 base_url，无法启用热轮换。");
            }

            if (!Uri.TryCreate(runtime.BaseUrl, UriKind.Absolute, out var uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("当前 Codex 配置的 base_url 不是有效的 http/https 地址。");
            }

            if (!runtime.HasToken)
            {
                throw new InvalidOperationException("当前 Codex 配置未解析到可用于热轮换的 API key 或 token。");
            }
        }

        private static async Task<CodexHotTokenSettings> LoadSettingsAsync(CancellationToken ct)
        {
            if (!File.Exists(SettingsPath))
            {
                return null;
            }

            try
            {
                var json = await ReadAllTextAsync(SettingsPath, ct).ConfigureAwait(false);
                return NormalizeSettings(JsonConvert.DeserializeObject<CodexHotTokenSettings>(json));
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Codex hot token settings load failed: {ErrorType}", ex.GetType().Name);
                return null;
            }
        }

        private static CodexHotTokenSettings NormalizeSettings(CodexHotTokenSettings settings)
        {
            if (settings == null)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(settings.ProviderId))
            {
                settings.ProviderId = ProviderId;
            }

            if (string.IsNullOrWhiteSpace(settings.TokenPath))
            {
                settings.TokenPath = TokenPath;
            }

            if (string.IsNullOrWhiteSpace(settings.ScriptPath))
            {
                settings.ScriptPath = ScriptPath;
            }

            if (settings.RefreshIntervalMs <= 0)
            {
                settings.RefreshIntervalMs = DefaultRefreshIntervalMs;
            }

            return settings;
        }

        private static async Task SaveSettingsAsync(CodexHotTokenSettings settings, CancellationToken ct)
        {
            Directory.CreateDirectory(HotTokenFolderPath);
            var json = JsonConvert.SerializeObject(settings ?? new CodexHotTokenSettings(), Formatting.Indented);
            await WriteAllTextAsync(SettingsPath, json, ct).ConfigureAwait(false);
        }

        private static async Task WriteTokenAsync(string token, CodexHotTokenSettings settings, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("热轮换 token 不能为空。");
            }

            Directory.CreateDirectory(Path.GetDirectoryName(settings.TokenPath) ?? HotTokenFolderPath);
            var plainBytes = Encoding.UTF8.GetBytes(token.Trim());
            var protectedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            await WriteAllTextAsync(settings.TokenPath, Convert.ToBase64String(protectedBytes), ct).ConfigureAwait(false);
        }

        private static async Task WriteScriptAsync(CodexHotTokenSettings settings, CancellationToken ct)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(settings.ScriptPath) ?? HotTokenFolderPath);
            var script = @"param(
    [Parameter(Mandatory=$true)]
    [string]$TokenPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Security
$text = [string]::Join('', (Get-Content -LiteralPath $TokenPath))
$protectedBytes = [Convert]::FromBase64String($text.Trim())
$plainBytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
    $protectedBytes,
    $null,
    [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
[Console]::Out.Write([System.Text.Encoding]::UTF8.GetString($plainBytes))
";
            await WriteAllTextAsync(settings.ScriptPath, script, ct).ConfigureAwait(false);
        }

        private static string BuildHotConfig(string configText, CodexHotTokenSettings settings, string wireApi)
        {
            var lines = SplitLines(RemoveHotProviderSections(configText, settings.ProviderId)).ToList();
            var withProvider = SetRootModelProvider(lines, settings.ProviderId);
            var builder = new StringBuilder();
            foreach (var line in withProvider)
            {
                builder.AppendLine(line);
            }

            if (builder.Length > 0 && !builder.ToString().EndsWith(Environment.NewLine + Environment.NewLine, StringComparison.Ordinal))
            {
                builder.AppendLine();
            }

            var effectiveWireApi = string.IsNullOrWhiteSpace(wireApi) ? "responses" : wireApi.Trim();
            builder.AppendLine("[model_providers." + settings.ProviderId + "]");
            builder.AppendLine("name = \"MyTools Hot NewAPI\"");
            builder.AppendLine("base_url = " + TomlQuote(settings.BaseUrl));
            builder.AppendLine("wire_api = " + TomlQuote(effectiveWireApi));
            builder.AppendLine();
            builder.AppendLine("[model_providers." + settings.ProviderId + ".auth]");
            builder.AppendLine("command = \"powershell.exe\"");
            builder.AppendLine("args = [\"-NoProfile\", \"-ExecutionPolicy\", \"Bypass\", \"-File\", "
                               + TomlQuote(settings.ScriptPath) + ", " + TomlQuote(settings.TokenPath) + "]");
            builder.AppendLine("timeout_ms = 5000");
            builder.AppendLine("refresh_interval_ms = " + settings.RefreshIntervalMs.ToString());
            return builder.ToString();
        }

        private static IEnumerable<string> SetRootModelProvider(IList<string> lines, string providerId)
        {
            var result = new List<string>();
            var replaced = false;
            var inserted = false;
            var inTable = false;

            foreach (var line in lines ?? new List<string>())
            {
                var trimmed = (line ?? string.Empty).Trim();
                if (trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal))
                {
                    if (!replaced && !inserted)
                    {
                        result.Add("model_provider = " + TomlQuote(providerId));
                        inserted = true;
                    }

                    inTable = true;
                }

                if (!inTable && IsAssignment(trimmed, "model_provider"))
                {
                    result.Add("model_provider = " + TomlQuote(providerId));
                    replaced = true;
                    continue;
                }

                result.Add(line ?? string.Empty);
            }

            if (!replaced && !inserted)
            {
                result.Insert(0, "model_provider = " + TomlQuote(providerId));
            }

            return result;
        }

        private static string RemoveHotProviderSections(string configText, string providerId)
        {
            var result = new List<string>();
            var skip = false;
            foreach (var line in SplitLines(configText))
            {
                var trimmed = (line ?? string.Empty).Trim();
                if (IsProviderSectionHeader(trimmed, providerId))
                {
                    skip = true;
                    continue;
                }

                if (skip && IsAnyTableHeader(trimmed))
                {
                    skip = false;
                }

                if (!skip)
                {
                    result.Add(line ?? string.Empty);
                }
            }

            return string.Join(Environment.NewLine, result);
        }

        private static bool IsProviderSectionHeader(string trimmed, string providerId)
        {
            return string.Equals(trimmed, "[model_providers." + providerId + "]", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(trimmed, "[model_providers." + providerId + ".auth]", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAnyTableHeader(string trimmed)
        {
            return trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal);
        }

        private static bool IsAssignment(string trimmed, string key)
        {
            var index = trimmed.IndexOf('=');
            return index > 0 && string.Equals(trimmed.Substring(0, index).Trim(), key, StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> SplitLines(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }

        private static bool SameBaseUrl(string left, string right)
        {
            return string.Equals(NormalizeBaseUrl(left), NormalizeBaseUrl(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeBaseUrl(string value)
        {
            return (value ?? string.Empty).Trim().TrimEnd('/');
        }

        private static string TomlQuote(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        private static string SafeHost(string baseUrl)
        {
            return Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri) ? uri.Host : string.Empty;
        }

        private static async Task<string> ReadAllTextAsync(string path, CancellationToken ct)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
            using (var reader = new StreamReader(stream, Encoding.UTF8, true))
            {
                return await reader.ReadToEndAsync().ConfigureAwait(false);
            }
        }

        private static async Task WriteAllTextAsync(string path, string value, CancellationToken ct)
        {
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                await writer.WriteAsync(value ?? string.Empty).ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);
            }
        }
    }

    public sealed class CodexHotTokenSettings
    {
        public bool Enabled { get; set; }
        public string ProviderId { get; set; }
        public string BaseUrl { get; set; }
        public string TokenPath { get; set; }
        public string ScriptPath { get; set; }
        public int RefreshIntervalMs { get; set; }
        public DateTime EnabledAtUtc { get; set; }
    }

    public sealed class CodexHotTokenSetupResult
    {
        public bool Success { get; set; }
        public bool RequiresRestart { get; set; }
        public string BaseUrl { get; set; }
        public string BackupPath { get; set; }
        public string Message { get; set; }
    }

    public sealed class CodexHotTokenApplyResult
    {
        public bool HotModeEnabled { get; set; }
        public bool Success { get; set; }
        public bool AllowFullConfigSwitch { get; set; }
        public string BaseUrl { get; set; }
        public string Message { get; set; }
    }
}
