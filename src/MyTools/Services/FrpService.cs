using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    public enum FrpState
    {
        Stopped,
        Starting,
        Running,
        Error
    }

    public sealed class FrpServerConfig
    {
        public string ServerAddress { get; set; } = string.Empty;
        public int ServerPort { get; set; } = 7000;
        public string EncryptedToken { get; set; } = string.Empty;
        public string ClientId { get; set; } = string.Empty;
    }

    public sealed class FrpTunnelRule : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _type = "tcp";
        private int _localPort;
        private int _remotePort;
        private string _description = string.Empty;
        private bool _isEnabled = true;

        public event PropertyChangedEventHandler PropertyChanged;

        public string Name
        {
            get => _name;
            set { if (_name == value) return; _name = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string Type
        {
            get => _type;
            set { if (_type == value) return; _type = string.IsNullOrWhiteSpace(value) ? "tcp" : value; OnPropertyChanged(); }
        }

        public int LocalPort
        {
            get => _localPort;
            set { if (_localPort == value) return; _localPort = value; OnPropertyChanged(); }
        }

        public int RemotePort
        {
            get => _remotePort;
            set { if (_remotePort == value) return; _remotePort = value; OnPropertyChanged(); }
        }

        public string Description
        {
            get => _description;
            set { if (_description == value) return; _description = value ?? string.Empty; OnPropertyChanged(); }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set { if (_isEnabled == value) return; _isEnabled = value; OnPropertyChanged(); }
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public static class FrpService
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        static FrpService()
        {
            ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.frpconfig.json");
            RulesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.frprules.json");
        }

        public static string ConfigPath { get; }
        public static string RulesPath { get; }

        public static async Task<string> EnsureFrpcExtractedAsync()
        {
            var tempDirectory = GetTempDirectory();
            var targetPath = Path.Combine(tempDirectory, "frpc.exe");

            try
            {
                Directory.CreateDirectory(tempDirectory);

                if (File.Exists(targetPath))
                {
                    var existing = new FileInfo(targetPath);
                    if (existing.Length > 1024 * 1024)
                    {
                        return targetPath;
                    }
                }

                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "MyTools.NativeBinaries.frpc.exe";
                var resourceStream = assembly.GetManifestResourceStream(resourceName);

                if (resourceStream == null)
                {
                    resourceName = assembly
                        .GetManifestResourceNames()
                        .FirstOrDefault(name => name.EndsWith(".frpc.exe", StringComparison.OrdinalIgnoreCase));

                    if (!string.IsNullOrEmpty(resourceName))
                    {
                        resourceStream = assembly.GetManifestResourceStream(resourceName);
                    }
                }

                if (resourceStream == null)
                {
                    throw new InvalidOperationException("frpc.exe 解压失败：未找到嵌入资源。");
                }

                using (resourceStream)
                using (var output = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    await resourceStream.CopyToAsync(output).ConfigureAwait(false);
                }

                return targetPath;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("frpc.exe 解压失败：" + ex.Message, ex);
            }
        }

        public static string BuildFrpcIni(FrpServerConfig config, string plainToken, IEnumerable<FrpTunnelRule> rules)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            var serverAddress = (config.ServerAddress ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(serverAddress))
            {
                throw new InvalidOperationException("请填写 frp 服务器地址。");
            }

            if (!IsValidPort(config.ServerPort))
            {
                throw new InvalidOperationException("frp 服务器端口必须在 1-65535。");
            }

            if (string.IsNullOrWhiteSpace(plainToken))
            {
                throw new InvalidOperationException("请填写 frp Token。");
            }

            var enabledRules = (rules ?? Enumerable.Empty<FrpTunnelRule>())
                .Where(rule => rule != null && rule.IsEnabled)
                .ToList();

            if (enabledRules.Count == 0)
            {
                throw new InvalidOperationException("请至少添加并启用一条隧道规则。");
            }

            var builder = new StringBuilder();
            builder.AppendLine("[common]");
            builder.Append("server_addr = ").AppendLine(serverAddress);
            builder.Append("server_port = ").AppendLine(config.ServerPort.ToString());
            builder.Append("token = ").AppendLine(plainToken ?? string.Empty);

            foreach (var rule in enabledRules)
            {
                if (!string.Equals(rule.Type, "tcp", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("当前版本仅支持 TCP 隧道规则。");
                }

                if (!IsValidPort(rule.LocalPort) || !IsValidPort(rule.RemotePort))
                {
                    throw new InvalidOperationException("隧道规则端口必须在 1-65535。");
                }

                builder.AppendLine();
                builder.Append('[').Append(BuildProxyName(config.ClientId, rule)).AppendLine("]");
                builder.AppendLine("type = tcp");
                builder.AppendLine("local_ip = 127.0.0.1");
                builder.Append("local_port = ").AppendLine(rule.LocalPort.ToString());
                builder.Append("remote_port = ").AppendLine(rule.RemotePort.ToString());
            }

            return builder.ToString();
        }

        public static string EncryptToken(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                return string.Empty;
            }

            var bytes = Encoding.UTF8.GetBytes(plainText);
            var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        public static string DecryptToken(string cipherText)
        {
            if (string.IsNullOrWhiteSpace(cipherText))
            {
                return string.Empty;
            }

            try
            {
                var protectedBytes = Convert.FromBase64String(cipherText);
                var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch
            {
                return string.Empty;
            }
        }

        public static async Task SaveConfigAsync(FrpServerConfig config)
        {
            var json = JsonConvert.SerializeObject(config ?? new FrpServerConfig(), Formatting.Indented);
            await WriteTextAsync(ConfigPath, json).ConfigureAwait(false);
        }

        public static async Task<FrpServerConfig> LoadConfigAsync()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    return new FrpServerConfig();
                }

                var json = await ReadTextAsync(ConfigPath).ConfigureAwait(false);
                return JsonConvert.DeserializeObject<FrpServerConfig>(json) ?? new FrpServerConfig();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("FRP config load failed: {Msg}", ex.Message);
                return new FrpServerConfig();
            }
        }

        public static async Task SaveRulesAsync(IEnumerable<FrpTunnelRule> rules)
        {
            var list = (rules ?? Enumerable.Empty<FrpTunnelRule>()).ToList();
            var json = JsonConvert.SerializeObject(list, Formatting.Indented);
            await WriteTextAsync(RulesPath, json).ConfigureAwait(false);
        }

        public static async Task<List<FrpTunnelRule>> LoadRulesAsync()
        {
            try
            {
                if (!File.Exists(RulesPath))
                {
                    return new List<FrpTunnelRule>();
                }

                var json = await ReadTextAsync(RulesPath).ConfigureAwait(false);
                var rules = JsonConvert.DeserializeObject<List<FrpTunnelRule>>(json) ?? new List<FrpTunnelRule>();
                foreach (var rule in rules)
                {
                    if (string.IsNullOrWhiteSpace(rule.Type))
                    {
                        rule.Type = "tcp";
                    }
                }

                return rules;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("FRP rules load failed: {Msg}", ex.Message);
                return new List<FrpTunnelRule>();
            }
        }

        public static bool IsValidPort(int port)
        {
            return port >= 1 && port <= 65535;
        }

        internal static string GetTempDirectory()
        {
            return Path.Combine(Path.GetTempPath(), "MyTools");
        }

        internal static string GetTempIniPath()
        {
            return Path.Combine(GetTempDirectory(), "frpc.ini");
        }

        private static string BuildProxyName(string clientId, FrpTunnelRule rule)
        {
            var machine = SanitizeName(Environment.MachineName);
            var id = string.IsNullOrWhiteSpace(clientId) ? "client" : clientId.Trim();
            if (id.Length > 8)
            {
                id = id.Substring(0, 8);
            }

            return SanitizeName($"mytools_{machine}_{id}_{rule.LocalPort}_{rule.RemotePort}");
        }

        private static string SanitizeName(string value)
        {
            var text = string.IsNullOrWhiteSpace(value) ? "pc" : value.Trim();
            return Regex.Replace(text, "[^A-Za-z0-9_]", "_");
        }

        private static async Task<string> ReadTextAsync(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
            using (var reader = new StreamReader(stream, Utf8NoBom, detectEncodingFromByteOrderMarks: true))
            {
                return await reader.ReadToEndAsync().ConfigureAwait(false);
            }
        }

        private static async Task WriteTextAsync(string path, string content)
        {
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, Utf8NoBom))
            {
                await writer.WriteAsync(content ?? string.Empty).ConfigureAwait(false);
            }
        }
    }

    public sealed class FrpProcessManager : IDisposable
    {
        private static readonly string[] SuccessMarkers =
        {
            "login to server success",
            "start proxy success",
            "proxy added"
        };

        private static readonly string[] FailureMarkers =
        {
            "login to server failed",
            "authorization failed",
            "port unavailable",
            "EOF"
        };

        private Process _process;
        private string _plainTokenForSanitize = string.Empty;
        private bool _disposed;

        public FrpState State { get; private set; } = FrpState.Stopped;
        public string StatusMessage { get; private set; } = "未运行";
        public event EventHandler StateChanged;

        public async Task StartAsync(string frpcExePath, string iniContent)
        {
            if (string.IsNullOrWhiteSpace(frpcExePath))
            {
                throw new ArgumentException("frpc.exe 路径为空。", nameof(frpcExePath));
            }

            Stop();
            SetState(FrpState.Starting, "正在连接...");
            _plainTokenForSanitize = ExtractToken(iniContent);

            Directory.CreateDirectory(FrpService.GetTempDirectory());
            var iniPath = FrpService.GetTempIniPath();
            await WriteIniAsync(iniPath, iniContent).ConfigureAwait(false);

            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = frpcExePath,
                    Arguments = "-c \"" + iniPath + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                },
                EnableRaisingEvents = true
            };

            process.OutputDataReceived += (sender, args) => HandleOutputLine(args.Data, completion);
            process.ErrorDataReceived += (sender, args) => HandleOutputLine(args.Data, completion);
            process.Exited += (sender, args) =>
            {
                if (State != FrpState.Stopped)
                {
                    SetState(FrpState.Error, "frpc 已退出");
                }

                completion.TrySetResult(false);
            };

            try
            {
                if (!process.Start())
                {
                    SetState(FrpState.Error, "frpc 启动失败");
                    completion.TrySetResult(false);
                    return;
                }

                _process = process;
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
            }
            catch (Exception ex)
            {
                process.Dispose();
                SetState(FrpState.Error, "frpc 启动失败：" + ex.Message);
                throw;
            }

            var fallbackAt = DateTime.UtcNow.AddSeconds(5);
            var timeoutAt = DateTime.UtcNow.AddSeconds(15);
            while (DateTime.UtcNow < timeoutAt)
            {
                var completed = await Task.WhenAny(completion.Task, Task.Delay(200)).ConfigureAwait(false);
                if (completed == completion.Task)
                {
                    return;
                }

                if (DateTime.UtcNow >= fallbackAt && IsProcessAlive(_process) && State == FrpState.Starting)
                {
                    SetState(FrpState.Running, "已启动，等待服务器确认");
                    completion.TrySetResult(true);
                    return;
                }
            }

            if (IsProcessAlive(_process) && State == FrpState.Starting)
            {
                SetState(FrpState.Running, "已启动，等待服务器确认");
            }
        }

        public void Stop()
        {
            var process = _process;
            _process = null;

            try
            {
                if (process != null && !process.HasExited)
                {
                    try { process.CloseMainWindow(); } catch { }

                    if (!process.WaitForExit(1500))
                    {
                        process.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("FRP stop failed: {Msg}", SanitizeLogLine(ex.Message));
            }
            finally
            {
                try { process?.Dispose(); } catch { }
                _plainTokenForSanitize = string.Empty;
                SetState(FrpState.Stopped, "未运行");
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            Stop();
        }

        private void HandleOutputLine(string line, TaskCompletionSource<bool> completion)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return;
            }

            var sanitized = SanitizeLogLine(line);
            AppLogService.Information("FRP: {Line}", sanitized);

            if (ContainsAny(line, SuccessMarkers))
            {
                SetState(FrpState.Running, "已连接");
                completion.TrySetResult(true);
                return;
            }

            if (ContainsAny(line, FailureMarkers))
            {
                SetState(FrpState.Error, "连接失败：" + sanitized);
                completion.TrySetResult(false);
            }
        }

        private void SetState(FrpState state, string statusMessage)
        {
            State = state;
            StatusMessage = statusMessage ?? string.Empty;
            StateChanged?.Invoke(this, EventArgs.Empty);
        }

        private string SanitizeLogLine(string line)
        {
            var sanitized = line ?? string.Empty;

            if (!string.IsNullOrEmpty(_plainTokenForSanitize))
            {
                sanitized = sanitized.Replace(_plainTokenForSanitize, "***");
            }

            sanitized = Regex.Replace(sanitized, "(?i)(token\\s*=\\s*)\\S+", "$1***");
            if (sanitized.Length > 300)
            {
                sanitized = sanitized.Substring(0, 300) + "...";
            }

            return sanitized;
        }

        private static bool ContainsAny(string line, IEnumerable<string> markers)
        {
            return markers.Any(marker => line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string ExtractToken(string iniContent)
        {
            if (string.IsNullOrWhiteSpace(iniContent))
            {
                return string.Empty;
            }

            var match = Regex.Match(iniContent, "^\\s*token\\s*=\\s*(.+?)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        private static bool IsProcessAlive(Process process)
        {
            try
            {
                return process != null && !process.HasExited;
            }
            catch
            {
                return false;
            }
        }

        private static async Task WriteIniAsync(string path, string content)
        {
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                await writer.WriteAsync(content ?? string.Empty).ConfigureAwait(false);
            }
        }
    }
}
