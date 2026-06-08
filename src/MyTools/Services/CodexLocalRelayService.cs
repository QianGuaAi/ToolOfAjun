using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public static class CodexLocalRelayService
    {
        private const string ProviderId = "mytools_local_relay";
        private const string RelayProcessArgument = "--codex-local-relay";
        private const int DefaultPort = 48176;
        private const int MaxHeaderBytes = 64 * 1024;

        private static readonly object SyncRoot = new object();
        private static readonly HttpClient UpstreamClient = CreateHttpClient();

        private static TcpListener _listener;
        private static CancellationTokenSource _listenerCts;
        private static Task _acceptLoopTask;
        private static CodexLocalRelayRuntimeState _runtimeState;

        private static string RelayFolderPath => Path.Combine(CodexProfileLibraryService.RootFolderPath, "LocalRelay");
        private static string SettingsPath => Path.Combine(RelayFolderPath, "local-relay.json");
        private static string LocalTokenPath => Path.Combine(RelayFolderPath, "local-token.dpapi");
        private static string ScriptPath => Path.Combine(RelayFolderPath, "ReadCodexLocalRelayToken.ps1");

        public static async Task<CodexLocalRelaySetupResult> EnableFromCurrentCodexAsync(CancellationToken ct)
        {
            Directory.CreateDirectory(RelayFolderPath);
            var existing = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (existing != null
                && existing.Enabled
                && !string.IsNullOrWhiteSpace(existing.ProtectedUpstreamTokenBase64))
            {
                var restored = NormalizeSettings(existing);
                var restoredLocalToken = !string.IsNullOrWhiteSpace(restored.ProtectedLocalApiTokenBase64)
                    ? UnprotectString(restored.ProtectedLocalApiTokenBase64)
                    : GenerateToken();
                restored.ProtectedLocalApiTokenBase64 = ProtectString(restoredLocalToken);
                restored.UpdatedAtUtc = DateTime.UtcNow;

                await WriteLocalTokenScriptAsync(restored, restoredLocalToken, ct).ConfigureAwait(false);
                await SaveSettingsAsync(restored, ct).ConfigureAwait(false);
                await EnsureBackgroundRelayProcessAsync(restored, ct).ConfigureAwait(false);
                var currentForRepair = await TryReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
                var restoredPin = await EnsureCurrentCodexConfigPinnedAsync(
                    restored,
                    currentForRepair?.ConfigTomlBytes,
                    currentForRepair?.AuthJsonBytes,
                    "local-relay-before-repair",
                    ct).ConfigureAwait(false);

                AppLogService.Information("Codex local relay mode repaired on port {Port}", restored.Port);
                return new CodexLocalRelaySetupResult
                {
                    Success = true,
                    RequiresRestart = restoredPin.RepairedConfig,
                    LocalBaseUrl = restored.LocalBaseUrl,
                    UpstreamBaseUrl = restored.UpstreamBaseUrl,
                    BackupPath = restoredPin.BackupPath,
                    Message = restoredPin.RepairedConfig
                        ? "已修复 Codex 本地中转配置。请重启一次 Codex App 让固定本地 base_url 生效。"
                        : "Codex 本地中转已启用并通过健康检查。"
                };
            }

            var current = await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
            var runtime = CodexRelayTestService.InspectRuntime(current.ConfigTomlBytes, current.AuthJsonBytes);
            ValidateRuntime(runtime);

            var port = existing != null && existing.Port > 0 ? existing.Port : FindAvailablePort(DefaultPort);
            var localToken = !string.IsNullOrWhiteSpace(existing?.ProtectedLocalApiTokenBase64)
                ? UnprotectString(existing.ProtectedLocalApiTokenBase64)
                : GenerateToken();

            var effectiveWireApi = string.IsNullOrWhiteSpace(runtime.WireApi) ? "responses" : runtime.WireApi.Trim();
            var settings = new CodexLocalRelaySettings
            {
                Enabled = true,
                ProviderId = ProviderId,
                Port = port,
                LocalBasePath = "/v1",
                LocalBaseUrl = BuildLocalBaseUrl(port),
                ScriptPath = ScriptPath,
                LocalTokenPath = LocalTokenPath,
                ProtectedLocalApiTokenBase64 = ProtectString(localToken),
                UpstreamBaseUrl = runtime.BaseUrl.Trim(),
                ProtectedUpstreamTokenBase64 = ProtectString(runtime.Token),
                WireApi = effectiveWireApi,
                ActiveDisplayName = "当前 Codex 配置",
                EnabledAtUtc = existing == null || existing.EnabledAtUtc == default(DateTime) ? DateTime.UtcNow : existing.EnabledAtUtc,
                UpdatedAtUtc = DateTime.UtcNow
            };

            await WriteLocalTokenScriptAsync(settings, localToken, ct).ConfigureAwait(false);
            await SaveSettingsAsync(settings, ct).ConfigureAwait(false);
            await EnsureBackgroundRelayProcessAsync(settings, ct).ConfigureAwait(false);
            var pin = await EnsureCurrentCodexConfigPinnedAsync(
                settings,
                current.ConfigTomlBytes,
                current.AuthJsonBytes,
                "local-relay-before-enable",
                ct).ConfigureAwait(false);

            AppLogService.Information("Codex local relay enabled on port {Port} for upstream host {Host}", port, SafeHost(settings.UpstreamBaseUrl));
            return new CodexLocalRelaySetupResult
            {
                Success = true,
                RequiresRestart = pin.RepairedConfig,
                LocalBaseUrl = settings.LocalBaseUrl,
                UpstreamBaseUrl = settings.UpstreamBaseUrl,
                BackupPath = pin.BackupPath,
                Message = pin.RepairedConfig
                    ? "已启用 Codex 本地中转。请重启一次 Codex App 让固定本地 base_url 生效；之后可热切换不同 NewAPI base_url 和 key。"
                    : "Codex 本地中转已启用并通过健康检查。"
            };
        }

        public static async Task<CodexLocalRelayApplyResult> TryApplyProfileAsync(
            byte[] configTomlBytes,
            byte[] authJsonBytes,
            string displayName,
            CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (settings == null || !settings.Enabled)
            {
                return new CodexLocalRelayApplyResult
                {
                    LocalRelayEnabled = false,
                    AllowFullConfigSwitch = true,
                    Message = "未启用 Codex 本地中转。"
                };
            }

            var runtime = CodexRelayTestService.InspectRuntime(configTomlBytes, authJsonBytes);
            if (string.IsNullOrWhiteSpace(runtime.BaseUrl))
            {
                return Stop("目标档案缺少 base_url，无法切换本地中转上游。");
            }

            if (!runtime.HasToken)
            {
                return Stop("目标档案未解析到可用于本地中转的 API key 或 token。");
            }

            var targetWireApi = string.IsNullOrWhiteSpace(runtime.WireApi) ? settings.WireApi : runtime.WireApi.Trim();
            if (!string.IsNullOrWhiteSpace(settings.WireApi)
                && !string.IsNullOrWhiteSpace(targetWireApi)
                && !string.Equals(settings.WireApi, targetWireApi, StringComparison.OrdinalIgnoreCase))
            {
                return Stop("目标档案 wire_api 与本地中转当前配置不一致；该项属于 Codex config.toml 配置，需完整切换并重启。");
            }

            settings.UpstreamBaseUrl = runtime.BaseUrl.Trim();
            settings.ProtectedUpstreamTokenBase64 = ProtectString(runtime.Token);
            settings.WireApi = string.IsNullOrWhiteSpace(settings.WireApi) ? targetWireApi : settings.WireApi;
            settings.ActiveDisplayName = string.IsNullOrWhiteSpace(displayName) ? "Codex 档案" : displayName.Trim();
            settings.UpdatedAtUtc = DateTime.UtcNow;

            await WriteLocalTokenScriptAsync(settings, UnprotectString(settings.ProtectedLocalApiTokenBase64), ct).ConfigureAwait(false);
            await SaveSettingsAsync(settings, ct).ConfigureAwait(false);
            await EnsureBackgroundRelayProcessAsync(settings, ct).ConfigureAwait(false);
            var pin = await EnsureCurrentCodexConfigPinnedAsync(
                settings,
                configTomlBytes,
                authJsonBytes,
                "local-relay-before-repair",
                ct).ConfigureAwait(false);

            AppLogService.Information("Codex local relay switched to upstream host {Host}", SafeHost(settings.UpstreamBaseUrl));
            return new CodexLocalRelayApplyResult
            {
                LocalRelayEnabled = true,
                Success = true,
                AllowFullConfigSwitch = false,
                LocalBaseUrl = settings.LocalBaseUrl,
                UpstreamBaseUrl = settings.UpstreamBaseUrl,
                RequiresCodexRestart = pin.RepairedConfig,
                Message = pin.RepairedConfig
                    ? "已切换本地中转上游，并修复 Codex 固定本地配置。请重启 Codex App 后使用。"
                    : "已切换本地中转上游，Codex App 下一次请求将使用新的 NewAPI 地址和 key。"
            };
        }

        public static async Task<bool> IsEnabledAsync(CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            return settings != null && settings.Enabled;
        }

        public static async Task<CodexLocalRelayStartResult> TryStartEnabledAsync(CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (settings == null || !settings.Enabled)
            {
                return new CodexLocalRelayStartResult
                {
                    Enabled = false,
                    Success = true,
                    Message = "未启用 Codex 本地中转。"
                };
            }

            try
            {
                await EnsureBackgroundRelayProcessAsync(settings, ct).ConfigureAwait(false);
                return new CodexLocalRelayStartResult
                {
                    Enabled = true,
                    Success = true,
                    LocalBaseUrl = settings.LocalBaseUrl,
                    UpstreamBaseUrl = settings.UpstreamBaseUrl,
                    Message = "Codex 本地中转已启动。"
                };
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Starting Codex local relay failed with {ErrorType}", ex.GetType().Name);
                return new CodexLocalRelayStartResult
                {
                    Enabled = true,
                    Success = false,
                    LocalBaseUrl = settings.LocalBaseUrl,
                    UpstreamBaseUrl = settings.UpstreamBaseUrl,
                    Message = ex.Message
                };
            }
        }

        public static async Task<CodexLocalRelayProbeResult> StartFromCurrentCodexAndProbeAsync(CancellationToken ct)
        {
            Directory.CreateDirectory(RelayFolderPath);
            var existing = await LoadSettingsAsync(ct).ConfigureAwait(false);
            var current = await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
            var runtime = CodexRelayTestService.InspectRuntime(current.ConfigTomlBytes, current.AuthJsonBytes);
            CodexLocalRelaySettings settings;
            string model;
            string wireApi;

            if (IsCurrentRuntimePinnedToExistingRelay(runtime, existing))
            {
                settings = NormalizeSettings(existing);
                model = runtime.Model ?? string.Empty;
                wireApi = string.IsNullOrWhiteSpace(runtime.WireApi) ? settings.WireApi : runtime.WireApi.Trim();
            }
            else
            {
                ValidateRuntime(runtime);
                settings = BuildSettingsFromRuntime(existing, runtime, "当前 Codex 配置");
                model = runtime.Model ?? string.Empty;
                wireApi = string.IsNullOrWhiteSpace(runtime.WireApi) ? settings.WireApi : runtime.WireApi.Trim();
            }

            if (string.IsNullOrWhiteSpace(settings?.ProtectedUpstreamTokenBase64))
            {
                throw new InvalidOperationException("本地中转缺少已保存的上游 key，无法测试当前 NewAPI。");
            }

            await WriteLocalTokenScriptAsync(settings, UnprotectString(settings.ProtectedLocalApiTokenBase64), ct).ConfigureAwait(false);
            await SaveSettingsAsync(settings, ct).ConfigureAwait(false);
            await EnsureBackgroundRelayProcessAsync(settings, ct).ConfigureAwait(false);
            var localToken = UnprotectString(settings.ProtectedLocalApiTokenBase64);
            var probeConfig = BuildRelayProbeConfigBytes(settings.LocalBaseUrl, localToken, model, wireApi);
            var probe = await CodexRelayTestService.TestAsync(probeConfig, Encoding.UTF8.GetBytes("{}"), ct).ConfigureAwait(false);

            return new CodexLocalRelayProbeResult
            {
                Enabled = true,
                Success = probe.Success,
                ProbeSuccess = probe.Success,
                LocalBaseUrl = settings.LocalBaseUrl,
                UpstreamBaseUrl = settings.UpstreamBaseUrl,
                ProbeMessage = probe.Message,
                Message = probe.Success
                    ? $"本地中转已启动，当前 NewAPI 经 relay 测试可达：{LimitProbeMessage(probe.Message)}"
                    : $"本地中转已启动，但当前 NewAPI 经 relay 测试不可达：{LimitProbeMessage(probe.Message)}"
            };
        }

        public static async Task<CodexLocalRelaySetupResult> ConfigureCurrentCodexToUseRelayAsync(CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (settings == null || !settings.Enabled || string.IsNullOrWhiteSpace(settings.ProtectedUpstreamTokenBase64))
            {
                return new CodexLocalRelaySetupResult
                {
                    Success = false,
                    Message = "请先点击“启动本地中转”，并确认当前 NewAPI 测试通过。"
                };
            }

            settings = NormalizeSettings(settings);
            await WriteLocalTokenScriptAsync(settings, UnprotectString(settings.ProtectedLocalApiTokenBase64), ct).ConfigureAwait(false);
            await SaveSettingsAsync(settings, ct).ConfigureAwait(false);
            await EnsureBackgroundRelayProcessAsync(settings, ct).ConfigureAwait(false);
            var current = await TryReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
            var pin = await EnsureCurrentCodexConfigPinnedAsync(
                settings,
                current?.ConfigTomlBytes,
                current?.AuthJsonBytes,
                "local-relay-before-codex-restart",
                ct).ConfigureAwait(false);

            return new CodexLocalRelaySetupResult
            {
                Success = true,
                RequiresRestart = true,
                LocalBaseUrl = settings.LocalBaseUrl,
                UpstreamBaseUrl = settings.UpstreamBaseUrl,
                BackupPath = pin.BackupPath,
                Message = pin.RepairedConfig
                    ? "已把 Codex 配置固定到本地中转。正在重启 Codex App..."
                    : "Codex 配置已是本地中转。正在重启 Codex App..."
            };
        }

        private static CodexLocalRelayApplyResult Stop(string message)
        {
            return new CodexLocalRelayApplyResult
            {
                LocalRelayEnabled = true,
                Success = false,
                AllowFullConfigSwitch = false,
                Message = message
            };
        }

        private static CodexLocalRelaySettings BuildSettingsFromRuntime(
            CodexLocalRelaySettings existing,
            CodexRelayRuntimeInfo runtime,
            string activeDisplayName)
        {
            var restored = NormalizeSettings(existing);
            var port = restored != null && restored.Port > 0 ? restored.Port : FindAvailablePort(DefaultPort);
            var localToken = restored != null && !string.IsNullOrWhiteSpace(restored.ProtectedLocalApiTokenBase64)
                ? UnprotectString(restored.ProtectedLocalApiTokenBase64)
                : GenerateToken();
            var effectiveWireApi = string.IsNullOrWhiteSpace(runtime?.WireApi)
                ? (string.IsNullOrWhiteSpace(restored?.WireApi) ? "responses" : restored.WireApi.Trim())
                : runtime.WireApi.Trim();

            return new CodexLocalRelaySettings
            {
                Enabled = true,
                ProviderId = ProviderId,
                Port = port,
                LocalBasePath = "/v1",
                LocalBaseUrl = BuildLocalBaseUrl(port),
                ScriptPath = ScriptPath,
                LocalTokenPath = LocalTokenPath,
                ProtectedLocalApiTokenBase64 = ProtectString(localToken),
                UpstreamBaseUrl = runtime.BaseUrl.Trim(),
                ProtectedUpstreamTokenBase64 = ProtectString(runtime.Token),
                WireApi = effectiveWireApi,
                ActiveDisplayName = string.IsNullOrWhiteSpace(activeDisplayName) ? "当前 Codex 配置" : activeDisplayName.Trim(),
                EnabledAtUtc = restored == null || restored.EnabledAtUtc == default(DateTime) ? DateTime.UtcNow : restored.EnabledAtUtc,
                UpdatedAtUtc = DateTime.UtcNow
            };
        }

        private static bool IsCurrentRuntimePinnedToExistingRelay(
            CodexRelayRuntimeInfo runtime,
            CodexLocalRelaySettings existing)
        {
            var settings = NormalizeSettings(existing);
            if (runtime == null
                || settings == null
                || !settings.Enabled
                || string.IsNullOrWhiteSpace(runtime.BaseUrl)
                || string.IsNullOrWhiteSpace(settings.ProtectedUpstreamTokenBase64))
            {
                return false;
            }

            if (string.Equals(
                    (runtime.BaseUrl ?? string.Empty).TrimEnd('/'),
                    (settings.LocalBaseUrl ?? string.Empty).TrimEnd('/'),
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!Uri.TryCreate(runtime.BaseUrl, UriKind.Absolute, out var runtimeUri))
            {
                return false;
            }

            return runtimeUri.Port == settings.Port
                   && (string.Equals(runtimeUri.Host, "localhost", StringComparison.OrdinalIgnoreCase)
                       || (IPAddress.TryParse(runtimeUri.Host, out var address) && IPAddress.IsLoopback(address)));
        }

        private static byte[] BuildRelayProbeConfigBytes(
            string localBaseUrl,
            string localToken,
            string model,
            string wireApi)
        {
            var builder = new StringBuilder();
            builder.AppendLine("model_provider = " + TomlQuote(ProviderId));
            if (!string.IsNullOrWhiteSpace(model))
            {
                builder.AppendLine("model = " + TomlQuote(model.Trim()));
            }

            builder.AppendLine();
            builder.AppendLine("[model_providers." + ProviderId + "]");
            builder.AppendLine("base_url = " + TomlQuote(localBaseUrl));
            builder.AppendLine("wire_api = " + TomlQuote(string.IsNullOrWhiteSpace(wireApi) ? "responses" : wireApi.Trim()));
            builder.AppendLine("api_key = " + TomlQuote(localToken));
            return Encoding.UTF8.GetBytes(builder.ToString());
        }

        private static string LimitProbeMessage(string value)
        {
            var text = string.IsNullOrWhiteSpace(value) ? "测试失败。" : value.Trim();
            return text.Length <= 160 ? text : text.Substring(0, 160) + "...";
        }

        public static bool IsRelayProcessMode(string[] args)
        {
            return (args ?? new string[0]).Any(arg =>
                string.Equals((arg ?? string.Empty).Trim(), RelayProcessArgument, StringComparison.OrdinalIgnoreCase));
        }

        public static async Task<CodexLocalRelayStartResult> StartRelayProcessModeAsync(CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (settings == null || !settings.Enabled)
            {
                return new CodexLocalRelayStartResult
                {
                    Enabled = false,
                    Success = false,
                    Message = "未启用 Codex 本地中转。"
                };
            }

            await EnsureStartedInCurrentProcessAsync(settings, ct).ConfigureAwait(false);
            return new CodexLocalRelayStartResult
            {
                Enabled = true,
                Success = true,
                LocalBaseUrl = settings.LocalBaseUrl,
                UpstreamBaseUrl = settings.UpstreamBaseUrl,
                Message = "Codex 本地中转后台进程已启动。"
            };
        }

        private static async Task EnsureBackgroundRelayProcessAsync(CodexLocalRelaySettings settings, CancellationToken ct)
        {
            settings = NormalizeSettings(settings);
            if (settings == null || !settings.Enabled)
            {
                return;
            }

            if (IsRelayProcessMode(Environment.GetCommandLineArgs()))
            {
                await EnsureStartedInCurrentProcessAsync(settings, ct).ConfigureAwait(false);
                return;
            }

            if (await IsLocalRelayHealthyAsync(settings, ct).ConfigureAwait(false))
            {
                return;
            }

            var exePath = ResolveCurrentExecutablePath();
            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = RelayProcessArgument,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            var process = Process.Start(startInfo);
            int? exitCode = null;

            for (var i = 0; i < 80; i++)
            {
                if (await IsLocalRelayHealthyAsync(settings, ct).ConfigureAwait(false))
                {
                    return;
                }

                if (!exitCode.HasValue && process != null && process.HasExited)
                {
                    exitCode = process.ExitCode;
                }

                await Task.Delay(500, ct).ConfigureAwait(false);
            }

            if (await IsLocalRelayHealthyAsync(settings, ct).ConfigureAwait(false))
            {
                return;
            }

            if (exitCode.HasValue)
            {
                throw new InvalidOperationException("Codex 本地中转后台进程已退出，退出码：" + exitCode.Value.ToString(CultureInfo.InvariantCulture));
            }

            throw new InvalidOperationException("Codex 本地中转后台进程启动后 40 秒内未通过健康检查。");
        }

        private static async Task<CodexLocalRelayPinResult> EnsureCurrentCodexConfigPinnedAsync(
            CodexLocalRelaySettings settings,
            byte[] fallbackConfigTomlBytes,
            byte[] fallbackAuthJsonBytes,
            string backupName,
            CancellationToken ct)
        {
            byte[] configBytes = null;
            byte[] authBytes = null;
            try
            {
                var current = await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
                configBytes = current.ConfigTomlBytes;
                authBytes = current.AuthJsonBytes;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Reading current Codex config before local relay pin failed: {ErrorType}", ex.GetType().Name);
            }

            configBytes = configBytes != null && configBytes.Length > 0 ? configBytes : fallbackConfigTomlBytes;
            authBytes = authBytes != null && authBytes.Length > 0 ? authBytes : fallbackAuthJsonBytes;
            configBytes = configBytes ?? new byte[0];
            if (authBytes == null || authBytes.Length == 0)
            {
                authBytes = Encoding.UTF8.GetBytes("{}");
            }

            var configText = Encoding.UTF8.GetString(configBytes);
            if (IsCodexConfigPinnedToLocalRelay(configText, settings))
            {
                return new CodexLocalRelayPinResult
                {
                    RepairedConfig = false,
                    Message = "当前 Codex 配置已固定到本地中转。"
                };
            }

            var backupPath = await CodexProfileLibraryService
                .BackupCurrentCodexFolderAsync(backupName, ct)
                .ConfigureAwait(false);
            var relayConfig = BuildRelayConfig(configText, settings);
            await CodexConfigProfileService
                .ApplyAsync(Encoding.UTF8.GetBytes(relayConfig), authBytes, ct)
                .ConfigureAwait(false);

            AppLogService.Information("Codex config pinned to local relay provider {ProviderId}", settings.ProviderId);
            return new CodexLocalRelayPinResult
            {
                RepairedConfig = true,
                BackupPath = backupPath,
                Message = "已把当前 Codex 配置固定到本地中转。"
            };
        }

        private static bool IsCodexConfigPinnedToLocalRelay(string configText, CodexLocalRelaySettings settings)
        {
            var text = configText ?? string.Empty;
            return text.IndexOf("model_provider = \"mytools_local_relay\"", StringComparison.OrdinalIgnoreCase) >= 0
                   && text.IndexOf("[model_providers.mytools_local_relay]", StringComparison.OrdinalIgnoreCase) >= 0
                   && text.IndexOf(settings.LocalBaseUrl ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static async Task<CodexProfileSourceFiles> TryReadCurrentCodexFilesAsync(CancellationToken ct)
        {
            try
            {
                return await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Reading current Codex files for local relay repair failed: {ErrorType}", ex.GetType().Name);
                return null;
            }
        }

        private static Task EnsureStartedInCurrentProcessAsync(CodexLocalRelaySettings settings, CancellationToken ct)
        {
            settings = NormalizeSettings(settings);
            if (settings == null || !settings.Enabled)
            {
                return Task.CompletedTask;
            }

            var state = BuildRuntimeState(settings);
            lock (SyncRoot)
            {
                _runtimeState = state;
                if (_listener != null)
                {
                    return Task.CompletedTask;
                }

                _listenerCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                _listener = new TcpListener(IPAddress.Loopback, settings.Port);
                _listener.Start();
                _acceptLoopTask = Task.Run(() => AcceptLoopAsync(_listener, _listenerCts.Token));
            }

            return Task.CompletedTask;
        }

        private static async Task AcceptLoopAsync(TcpListener listener, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                TcpClient client = null;
                try
                {
                    client = await listener.AcceptTcpClientAsync().ConfigureAwait(false);
                    var captured = client;
                    client = null;
                    _ = Task.Run(() => HandleClientAsync(captured, ct), ct);
                }
                catch (ObjectDisposedException)
                {
                    return;
                }
                catch (SocketException) when (ct.IsCancellationRequested)
                {
                    return;
                }
                catch (Exception ex)
                {
                    AppLogService.Warning("Codex local relay accept failed: {ErrorType}", ex.GetType().Name);
                }
                finally
                {
                    try { client?.Close(); } catch { }
                }
            }
        }

        private static async Task HandleClientAsync(TcpClient client, CancellationToken ct)
        {
            using (client)
            {
                client.NoDelay = true;
                using (var stream = client.GetStream())
                {
                    CodexLocalRelayRequest request = null;
                    try
                    {
                        request = await ReadRequestAsync(stream, ct).ConfigureAwait(false);
                        if (request == null)
                        {
                            return;
                        }

                        if (string.Equals(request.Path, "/", StringComparison.Ordinal)
                            || string.Equals(request.Path, "/health", StringComparison.OrdinalIgnoreCase))
                        {
                            await WritePlainResponseAsync(stream, HttpStatusCode.OK, "MyTools Codex local relay is running.", ct).ConfigureAwait(false);
                            return;
                        }

                        await RefreshRuntimeStateFromSettingsAsync(ct).ConfigureAwait(false);
                        var state = GetRuntimeState();
                        if (state == null)
                        {
                            await WritePlainResponseAsync(stream, HttpStatusCode.ServiceUnavailable, "Codex local relay is not configured.", ct).ConfigureAwait(false);
                            return;
                        }

                        if (!IsAuthorized(request, state.LocalApiToken))
                        {
                            await WritePlainResponseAsync(stream, HttpStatusCode.Unauthorized, "Unauthorized.", ct).ConfigureAwait(false);
                            return;
                        }

                        var upstreamUri = BuildUpstreamUri(state.UpstreamBaseUri, state.LocalBasePath, request.Path, request.Query);
                        using (var upstreamRequest = BuildUpstreamRequest(request, upstreamUri, state.UpstreamToken))
                        using (var response = await UpstreamClient.SendAsync(upstreamRequest, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                        {
                            await WriteUpstreamResponseAsync(stream, response, ct).ConfigureAwait(false);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLogService.Warning("Codex local relay request failed: {ErrorType}", ex.GetType().Name);
                        try
                        {
                            await WritePlainResponseAsync(stream, HttpStatusCode.BadGateway, "Codex local relay request failed.", CancellationToken.None).ConfigureAwait(false);
                        }
                        catch
                        {
                        }
                    }
                    finally
                    {
                        try { stream.Close(); } catch { }
                    }
                }
            }
        }

        private static HttpRequestMessage BuildUpstreamRequest(CodexLocalRelayRequest request, Uri upstreamUri, string upstreamToken)
        {
            var message = new HttpRequestMessage(new HttpMethod(request.Method), upstreamUri);
            if (request.Body != null && request.Body.Length > 0)
            {
                message.Content = new ByteArrayContent(request.Body);
            }

            foreach (var header in request.Headers)
            {
                if (ShouldSkipRequestHeader(header.Key))
                {
                    continue;
                }

                if (!message.Headers.TryAddWithoutValidation(header.Key, header.Value)
                    && message.Content != null)
                {
                    message.Content.Headers.TryAddWithoutValidation(header.Key, header.Value);
                }
            }

            message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", NormalizeBearerToken(upstreamToken));
            return message;
        }

        private static async Task<CodexLocalRelayRequest> ReadRequestAsync(NetworkStream stream, CancellationToken ct)
        {
            var header = await ReadHeaderBlockAsync(stream, ct).ConfigureAwait(false);
            if (header == null)
            {
                return null;
            }

            var headerText = Encoding.ASCII.GetString(header.HeaderBytes);
            var lines = headerText.Replace("\r\n", "\n").Split('\n');
            if (lines.Length == 0 || string.IsNullOrWhiteSpace(lines[0]))
            {
                return null;
            }

            var parts = lines[0].Split(new[] { ' ' }, 3, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
            {
                return null;
            }

            var headers = new List<KeyValuePair<string, string>>();
            for (var i = 1; i < lines.Length; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                headers.Add(new KeyValuePair<string, string>(
                    line.Substring(0, colon).Trim(),
                    line.Substring(colon + 1).Trim()));
            }

            var pathAndQuery = parts[1];
            if (Uri.TryCreate(pathAndQuery, UriKind.Absolute, out var absoluteRequestUri))
            {
                pathAndQuery = absoluteRequestUri.PathAndQuery;
            }

            var queryIndex = pathAndQuery.IndexOf('?');
            var path = queryIndex >= 0 ? pathAndQuery.Substring(0, queryIndex) : pathAndQuery;
            var query = queryIndex >= 0 ? pathAndQuery.Substring(queryIndex) : string.Empty;
            var pendingReader = new PendingNetworkReader(stream, header.RemainderBytes);
            var body = await ReadBodyAsync(pendingReader, headers, ct).ConfigureAwait(false);

            return new CodexLocalRelayRequest
            {
                Method = parts[0],
                Path = string.IsNullOrWhiteSpace(path) ? "/" : path,
                Query = query,
                Headers = headers,
                Body = body
            };
        }

        private static async Task<byte[]> ReadBodyAsync(PendingNetworkReader reader, IList<KeyValuePair<string, string>> headers, CancellationToken ct)
        {
            var transferEncoding = GetHeader(headers, "Transfer-Encoding");
            if (!string.IsNullOrWhiteSpace(transferEncoding)
                && transferEncoding.IndexOf("chunked", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return await ReadChunkedBodyAsync(reader, ct).ConfigureAwait(false);
            }

            var contentLengthText = GetHeader(headers, "Content-Length");
            if (!long.TryParse(contentLengthText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var contentLength)
                || contentLength <= 0)
            {
                return new byte[0];
            }

            if (contentLength > 128 * 1024 * 1024)
            {
                throw new InvalidOperationException("Codex 本地中转请求体超过 128 MB，已拒绝。");
            }

            return await reader.ReadBytesAsync((int)contentLength, ct).ConfigureAwait(false);
        }

        private static async Task<byte[]> ReadChunkedBodyAsync(PendingNetworkReader reader, CancellationToken ct)
        {
            using (var body = new MemoryStream())
            {
                while (true)
                {
                    var line = await reader.ReadAsciiLineAsync(ct).ConfigureAwait(false);
                    var sizeText = (line.Split(';').FirstOrDefault() ?? string.Empty).Trim();
                    if (!int.TryParse(sizeText, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var size))
                    {
                        throw new InvalidOperationException("无法解析 chunked 请求体。");
                    }

                    if (size == 0)
                    {
                        while (!string.IsNullOrEmpty(await reader.ReadAsciiLineAsync(ct).ConfigureAwait(false)))
                        {
                        }

                        break;
                    }

                    var chunk = await reader.ReadBytesAsync(size, ct).ConfigureAwait(false);
                    body.Write(chunk, 0, chunk.Length);
                    await reader.ReadBytesAsync(2, ct).ConfigureAwait(false);
                }

                return body.ToArray();
            }
        }

        private static async Task<HeaderReadResult> ReadHeaderBlockAsync(NetworkStream stream, CancellationToken ct)
        {
            using (var buffer = new MemoryStream())
            {
                var temp = new byte[4096];
                while (buffer.Length <= MaxHeaderBytes)
                {
                    var read = await stream.ReadAsync(temp, 0, temp.Length, ct).ConfigureAwait(false);
                    if (read <= 0)
                    {
                        return null;
                    }

                    buffer.Write(temp, 0, read);
                    var data = buffer.ToArray();
                    var end = FindHeaderEnd(data);
                    if (end >= 0)
                    {
                        return new HeaderReadResult
                        {
                            HeaderBytes = data.Take(end).ToArray(),
                            RemainderBytes = data.Skip(end + 4).ToArray()
                        };
                    }
                }
            }

            throw new InvalidOperationException("HTTP 请求头过大。");
        }

        private static int FindHeaderEnd(byte[] data)
        {
            if (data == null || data.Length < 4)
            {
                return -1;
            }

            for (var i = 0; i <= data.Length - 4; i++)
            {
                if (data[i] == 13 && data[i + 1] == 10 && data[i + 2] == 13 && data[i + 3] == 10)
                {
                    return i;
                }
            }

            return -1;
        }

        private static async Task WriteUpstreamResponseAsync(NetworkStream stream, HttpResponseMessage response, CancellationToken ct)
        {
            var builder = new StringBuilder();
            builder.Append("HTTP/1.1 ")
                .Append((int)response.StatusCode)
                .Append(' ')
                .Append(string.IsNullOrWhiteSpace(response.ReasonPhrase) ? response.StatusCode.ToString() : response.ReasonPhrase)
                .Append("\r\n");

            foreach (var header in response.Headers)
            {
                AppendResponseHeader(builder, header.Key, header.Value);
            }

            foreach (var header in response.Content.Headers)
            {
                AppendResponseHeader(builder, header.Key, header.Value);
            }

            builder.Append("Connection: close\r\n\r\n");
            var headerBytes = Encoding.ASCII.GetBytes(builder.ToString());
            await stream.WriteAsync(headerBytes, 0, headerBytes.Length, ct).ConfigureAwait(false);

            using (var responseStream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
            {
                await responseStream.CopyToAsync(stream, 81920, ct).ConfigureAwait(false);
            }
        }

        private static void AppendResponseHeader(StringBuilder builder, string key, IEnumerable<string> values)
        {
            if (string.IsNullOrWhiteSpace(key) || ShouldSkipResponseHeader(key))
            {
                return;
            }

            foreach (var value in values ?? Enumerable.Empty<string>())
            {
                builder.Append(key).Append(": ").Append(value).Append("\r\n");
            }
        }

        private static async Task WritePlainResponseAsync(NetworkStream stream, HttpStatusCode statusCode, string message, CancellationToken ct)
        {
            var body = Encoding.UTF8.GetBytes(message ?? string.Empty);
            var header = "HTTP/1.1 " + (int)statusCode + " " + statusCode + "\r\n"
                         + "Content-Type: text/plain; charset=utf-8\r\n"
                         + "Content-Length: " + body.Length.ToString(CultureInfo.InvariantCulture) + "\r\n"
                         + "Connection: close\r\n\r\n";
            var headerBytes = Encoding.ASCII.GetBytes(header);
            await stream.WriteAsync(headerBytes, 0, headerBytes.Length, ct).ConfigureAwait(false);
            await stream.WriteAsync(body, 0, body.Length, ct).ConfigureAwait(false);
        }

        private static bool IsAuthorized(CodexLocalRelayRequest request, string localApiToken)
        {
            if (string.IsNullOrWhiteSpace(localApiToken))
            {
                return true;
            }

            var authorization = GetHeader(request.Headers, "Authorization");
            if (string.IsNullOrWhiteSpace(authorization))
            {
                return false;
            }

            return string.Equals(NormalizeBearerToken(authorization), NormalizeBearerToken(localApiToken), StringComparison.Ordinal);
        }

        private static Uri BuildUpstreamUri(Uri upstreamBaseUri, string localBasePath, string requestPath, string query)
        {
            var relativePath = requestPath ?? string.Empty;
            var normalizedLocalPath = NormalizePath(localBasePath);
            var normalizedRequestPath = NormalizePath(relativePath);
            if (normalizedRequestPath.Equals(normalizedLocalPath, StringComparison.OrdinalIgnoreCase))
            {
                relativePath = string.Empty;
            }
            else if (normalizedRequestPath.StartsWith(normalizedLocalPath + "/", StringComparison.OrdinalIgnoreCase))
            {
                relativePath = normalizedRequestPath.Substring(normalizedLocalPath.Length).TrimStart('/');
            }
            else
            {
                relativePath = normalizedRequestPath.TrimStart('/');
            }

            var builder = new UriBuilder(upstreamBaseUri);
            var basePath = (builder.Path ?? string.Empty).TrimEnd('/');
            builder.Path = string.IsNullOrWhiteSpace(relativePath)
                ? basePath
                : basePath + "/" + relativePath.TrimStart('/');
            builder.Query = (query ?? string.Empty).TrimStart('?');
            return builder.Uri;
        }

        private static string BuildRelayConfig(string configText, CodexLocalRelaySettings settings)
        {
            var lines = SplitLines(RemoveRelayProviderSections(configText, settings.ProviderId)).ToList();
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

            var effectiveWireApi = string.IsNullOrWhiteSpace(settings.WireApi) ? "responses" : settings.WireApi.Trim();
            builder.AppendLine("[model_providers." + settings.ProviderId + "]");
            builder.AppendLine("name = \"MyTools Local Relay\"");
            builder.AppendLine("base_url = " + TomlQuote(settings.LocalBaseUrl));
            builder.AppendLine("wire_api = " + TomlQuote(effectiveWireApi));
            builder.AppendLine();
            builder.AppendLine("[model_providers." + settings.ProviderId + ".auth]");
            builder.AppendLine("command = \"powershell.exe\"");
            builder.AppendLine("args = [\"-NoProfile\", \"-ExecutionPolicy\", \"Bypass\", \"-File\", "
                               + TomlQuote(settings.ScriptPath) + ", " + TomlQuote(settings.LocalTokenPath) + "]");
            builder.AppendLine("timeout_ms = 5000");
            builder.AppendLine("refresh_interval_ms = 60000");
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

        private static string RemoveRelayProviderSections(string configText, string providerId)
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

        private static async Task<CodexLocalRelaySettings> LoadSettingsAsync(CancellationToken ct)
        {
            if (!File.Exists(SettingsPath))
            {
                return null;
            }

            try
            {
                var json = await ReadAllTextAsync(SettingsPath, ct).ConfigureAwait(false);
                return NormalizeSettings(JsonConvert.DeserializeObject<CodexLocalRelaySettings>(json));
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Codex local relay settings load failed: {ErrorType}", ex.GetType().Name);
                return null;
            }
        }

        private static CodexLocalRelaySettings NormalizeSettings(CodexLocalRelaySettings settings)
        {
            if (settings == null)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(settings.ProviderId))
            {
                settings.ProviderId = ProviderId;
            }

            if (settings.Port <= 0)
            {
                settings.Port = DefaultPort;
            }

            if (string.IsNullOrWhiteSpace(settings.LocalBasePath))
            {
                settings.LocalBasePath = "/v1";
            }

            settings.LocalBasePath = NormalizePath(settings.LocalBasePath);
            if (string.IsNullOrWhiteSpace(settings.LocalBaseUrl))
            {
                settings.LocalBaseUrl = BuildLocalBaseUrl(settings.Port);
            }

            if (string.IsNullOrWhiteSpace(settings.LocalTokenPath))
            {
                settings.LocalTokenPath = LocalTokenPath;
            }

            if (string.IsNullOrWhiteSpace(settings.ScriptPath))
            {
                settings.ScriptPath = ScriptPath;
            }

            if (string.IsNullOrWhiteSpace(settings.WireApi))
            {
                settings.WireApi = "responses";
            }

            return settings;
        }

        private static async Task SaveSettingsAsync(CodexLocalRelaySettings settings, CancellationToken ct)
        {
            Directory.CreateDirectory(RelayFolderPath);
            var json = JsonConvert.SerializeObject(settings ?? new CodexLocalRelaySettings(), Formatting.Indented);
            await WriteAllTextAsync(SettingsPath, json, ct).ConfigureAwait(false);
        }

        private static CodexLocalRelayRuntimeState BuildRuntimeState(CodexLocalRelaySettings settings)
        {
            if (!Uri.TryCreate(settings.UpstreamBaseUrl, UriKind.Absolute, out var upstreamBaseUri)
                || (upstreamBaseUri.Scheme != Uri.UriSchemeHttp && upstreamBaseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("Codex 本地中转上游 base_url 不是有效的 http/https 地址。");
            }

            return new CodexLocalRelayRuntimeState
            {
                LocalBasePath = settings.LocalBasePath,
                LocalApiToken = UnprotectString(settings.ProtectedLocalApiTokenBase64),
                UpstreamBaseUri = upstreamBaseUri,
                UpstreamToken = UnprotectString(settings.ProtectedUpstreamTokenBase64)
            };
        }

        private static CodexLocalRelayRuntimeState GetRuntimeState()
        {
            lock (SyncRoot)
            {
                return _runtimeState;
            }
        }

        private static async Task RefreshRuntimeStateFromSettingsAsync(CancellationToken ct)
        {
            var settings = await LoadSettingsAsync(ct).ConfigureAwait(false);
            if (settings == null || !settings.Enabled)
            {
                lock (SyncRoot)
                {
                    _runtimeState = null;
                }

                return;
            }

            var state = BuildRuntimeState(settings);
            lock (SyncRoot)
            {
                _runtimeState = state;
            }
        }

        private static async Task WriteLocalTokenScriptAsync(CodexLocalRelaySettings settings, string localToken, CancellationToken ct)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(settings.LocalTokenPath) ?? RelayFolderPath);
            Directory.CreateDirectory(Path.GetDirectoryName(settings.ScriptPath) ?? RelayFolderPath);

            var protectedBytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(localToken ?? string.Empty), null, DataProtectionScope.CurrentUser);
            await WriteAllTextAsync(settings.LocalTokenPath, Convert.ToBase64String(protectedBytes), ct).ConfigureAwait(false);

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

        private static void ValidateRuntime(CodexRelayRuntimeInfo runtime)
        {
            if (runtime == null || string.IsNullOrWhiteSpace(runtime.BaseUrl))
            {
                throw new InvalidOperationException("当前 Codex 配置缺少 base_url，无法启动本地中转。");
            }

            if (!Uri.TryCreate(runtime.BaseUrl, UriKind.Absolute, out var uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("当前 Codex 配置的 base_url 不是有效的 http/https 地址。");
            }

            if (!runtime.HasToken)
            {
                throw new InvalidOperationException("当前 Codex 配置未解析到可用于本地中转的 API key 或 token。");
            }
        }

        private static int FindAvailablePort(int startPort)
        {
            for (var port = startPort; port < startPort + 20; port++)
            {
                try
                {
                    var probe = new TcpListener(IPAddress.Loopback, port);
                    try
                    {
                        probe.Start();
                        return port;
                    }
                    finally
                    {
                        probe.Stop();
                    }
                }
                catch
                {
                }
            }

            throw new InvalidOperationException("未找到可用于 Codex 本地中转的 127.0.0.1 端口。");
        }

        private static async Task<bool> IsLocalRelayHealthyAsync(CodexLocalRelaySettings settings, CancellationToken ct)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    var connectTask = client.ConnectAsync(IPAddress.Loopback, settings.Port);
                    if (await Task.WhenAny(connectTask, Task.Delay(1000, ct)).ConfigureAwait(false) != connectTask)
                    {
                        return false;
                    }

                    await connectTask.ConfigureAwait(false);
                    using (var stream = client.GetStream())
                    {
                        var request = Encoding.ASCII.GetBytes("GET /health HTTP/1.1\r\nHost: 127.0.0.1\r\nConnection: close\r\n\r\n");
                        await stream.WriteAsync(request, 0, request.Length, ct).ConfigureAwait(false);
                        var text = await ReadLocalRelayHealthResponseAsync(stream, ct).ConfigureAwait(false);
                        return text.IndexOf("MyTools Codex local relay", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

        private static async Task<string> ReadLocalRelayHealthResponseAsync(NetworkStream stream, CancellationToken ct)
        {
            var buffer = new byte[512];
            using (var memory = new MemoryStream())
            {
                var deadline = DateTime.UtcNow.AddSeconds(2);
                while (memory.Length < 4096 && DateTime.UtcNow < deadline)
                {
                    var remaining = Math.Max(1, (int)(deadline - DateTime.UtcNow).TotalMilliseconds);
                    var readTask = stream.ReadAsync(buffer, 0, buffer.Length, ct);
                    if (await Task.WhenAny(readTask, Task.Delay(Math.Min(remaining, 500), ct)).ConfigureAwait(false) != readTask)
                    {
                        break;
                    }

                    var read = await readTask.ConfigureAwait(false);
                    if (read <= 0)
                    {
                        break;
                    }

                    memory.Write(buffer, 0, read);
                    var text = Encoding.UTF8.GetString(memory.ToArray());
                    if (text.IndexOf("MyTools Codex local relay", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return text;
                    }
                }

                return Encoding.UTF8.GetString(memory.ToArray());
            }
        }

        private static string ResolveCurrentExecutablePath()
        {
            try
            {
                var processPath = Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(processPath) && File.Exists(processPath))
                {
                    return processPath;
                }
            }
            catch
            {
            }

            var assemblyPath = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            if (!string.IsNullOrWhiteSpace(assemblyPath) && File.Exists(assemblyPath))
            {
                return assemblyPath;
            }

            throw new InvalidOperationException("无法定位 MyTools.exe，不能启动 Codex 本地中转后台进程。");
        }

        private static HttpClient CreateHttpClient()
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch
            {
            }

            return new HttpClient(new HttpClientHandler
            {
                AllowAutoRedirect = false,
                UseProxy = true
            })
            {
                Timeout = Timeout.InfiniteTimeSpan
            };
        }

        private static bool ShouldSkipRequestHeader(string name)
        {
            return HeaderEquals(name, "Host")
                   || HeaderEquals(name, "Authorization")
                   || HeaderEquals(name, "Content-Length")
                   || HeaderEquals(name, "Transfer-Encoding")
                   || HeaderEquals(name, "Connection")
                   || HeaderEquals(name, "Keep-Alive")
                   || HeaderEquals(name, "Proxy-Connection")
                   || HeaderEquals(name, "Expect");
        }

        private static bool ShouldSkipResponseHeader(string name)
        {
            return HeaderEquals(name, "Transfer-Encoding")
                   || HeaderEquals(name, "Connection")
                   || HeaderEquals(name, "Keep-Alive")
                   || HeaderEquals(name, "Proxy-Connection");
        }

        private static bool HeaderEquals(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetHeader(IEnumerable<KeyValuePair<string, string>> headers, string name)
        {
            return headers?
                .FirstOrDefault(header => HeaderEquals(header.Key, name))
                .Value ?? string.Empty;
        }

        private static string NormalizePath(string value)
        {
            var path = (value ?? string.Empty).Trim();
            if (path.Length == 0)
            {
                return "/";
            }

            if (!path.StartsWith("/", StringComparison.Ordinal))
            {
                path = "/" + path;
            }

            return path.TrimEnd('/');
        }

        private static string BuildLocalBaseUrl(int port)
        {
            return "http://127.0.0.1:" + port.ToString(CultureInfo.InvariantCulture) + "/v1";
        }

        private static string NormalizeBearerToken(string token)
        {
            var value = (token ?? string.Empty).Trim();
            return value.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase)
                ? value.Substring("Bearer ".Length).Trim()
                : value;
        }

        private static string ProtectString(string value)
        {
            var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            return Convert.ToBase64String(ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser));
        }

        private static string UnprotectString(string protectedBase64)
        {
            if (string.IsNullOrWhiteSpace(protectedBase64))
            {
                return string.Empty;
            }

            var protectedBytes = Convert.FromBase64String(protectedBase64);
            var plainBytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plainBytes);
        }

        private static string GenerateToken()
        {
            var bytes = new byte[32];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            return Convert.ToBase64String(bytes);
        }

        private static string TomlQuote(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        private static IEnumerable<string> SplitLines(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
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

        private sealed class HeaderReadResult
        {
            public byte[] HeaderBytes { get; set; }
            public byte[] RemainderBytes { get; set; }
        }

        private sealed class CodexLocalRelayRequest
        {
            public string Method { get; set; }
            public string Path { get; set; }
            public string Query { get; set; }
            public List<KeyValuePair<string, string>> Headers { get; set; }
            public byte[] Body { get; set; }
        }

        private sealed class CodexLocalRelayRuntimeState
        {
            public string LocalBasePath { get; set; }
            public string LocalApiToken { get; set; }
            public Uri UpstreamBaseUri { get; set; }
            public string UpstreamToken { get; set; }
        }

        private sealed class PendingNetworkReader
        {
            private readonly NetworkStream _stream;
            private byte[] _buffer;
            private int _offset;
            private int _count;

            public PendingNetworkReader(NetworkStream stream, byte[] initial)
            {
                _stream = stream;
                _buffer = initial ?? new byte[0];
                _offset = 0;
                _count = _buffer.Length;
            }

            public async Task<byte[]> ReadBytesAsync(int count, CancellationToken ct)
            {
                var result = new byte[count];
                var written = 0;
                while (written < count)
                {
                    if (_count > 0)
                    {
                        var take = Math.Min(_count, count - written);
                        Buffer.BlockCopy(_buffer, _offset, result, written, take);
                        _offset += take;
                        _count -= take;
                        written += take;
                        continue;
                    }

                    _buffer = new byte[Math.Min(8192, count - written)];
                    _offset = 0;
                    _count = await _stream.ReadAsync(_buffer, 0, _buffer.Length, ct).ConfigureAwait(false);
                    if (_count <= 0)
                    {
                        throw new EndOfStreamException("读取 HTTP 请求体时连接已关闭。");
                    }
                }

                return result;
            }

            public async Task<string> ReadAsciiLineAsync(CancellationToken ct)
            {
                using (var line = new MemoryStream())
                {
                    while (true)
                    {
                        var b = await ReadBytesAsync(1, ct).ConfigureAwait(false);
                        if (b[0] == 10)
                        {
                            var bytes = line.ToArray();
                            if (bytes.Length > 0 && bytes[bytes.Length - 1] == 13)
                            {
                                Array.Resize(ref bytes, bytes.Length - 1);
                            }

                            return Encoding.ASCII.GetString(bytes);
                        }

                        line.WriteByte(b[0]);
                        if (line.Length > 8192)
                        {
                            throw new InvalidOperationException("HTTP chunk 行过长。");
                        }
                    }
                }
            }
        }
    }

    public sealed class CodexLocalRelaySettings
    {
        public bool Enabled { get; set; }
        public string ProviderId { get; set; }
        public int Port { get; set; }
        public string LocalBasePath { get; set; }
        public string LocalBaseUrl { get; set; }
        public string ScriptPath { get; set; }
        public string LocalTokenPath { get; set; }
        public string ProtectedLocalApiTokenBase64 { get; set; }
        public string UpstreamBaseUrl { get; set; }
        public string ProtectedUpstreamTokenBase64 { get; set; }
        public string WireApi { get; set; }
        public string ActiveDisplayName { get; set; }
        public DateTime EnabledAtUtc { get; set; }
        public DateTime UpdatedAtUtc { get; set; }
    }

    public sealed class CodexLocalRelaySetupResult
    {
        public bool Success { get; set; }
        public bool RequiresRestart { get; set; }
        public string LocalBaseUrl { get; set; }
        public string UpstreamBaseUrl { get; set; }
        public string BackupPath { get; set; }
        public string Message { get; set; }
    }

    public sealed class CodexLocalRelayApplyResult
    {
        public bool LocalRelayEnabled { get; set; }
        public bool Success { get; set; }
        public bool AllowFullConfigSwitch { get; set; }
        public bool RequiresCodexRestart { get; set; }
        public string LocalBaseUrl { get; set; }
        public string UpstreamBaseUrl { get; set; }
        public string Message { get; set; }
    }

    public sealed class CodexLocalRelayStartResult
    {
        public bool Enabled { get; set; }
        public bool Success { get; set; }
        public string LocalBaseUrl { get; set; }
        public string UpstreamBaseUrl { get; set; }
        public string Message { get; set; }
    }

    public sealed class CodexLocalRelayProbeResult
    {
        public bool Enabled { get; set; }
        public bool Success { get; set; }
        public bool ProbeSuccess { get; set; }
        public string LocalBaseUrl { get; set; }
        public string UpstreamBaseUrl { get; set; }
        public string ProbeMessage { get; set; }
        public string Message { get; set; }
    }

    internal sealed class CodexLocalRelayPinResult
    {
        public bool RepairedConfig { get; set; }
        public string BackupPath { get; set; }
        public string Message { get; set; }
    }
}
