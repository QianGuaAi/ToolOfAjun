using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;
using MyTools.Services;

namespace MyTools.Services
{
    public class StartupItem
    {
        public string Name { get; set; }
        public string Command { get; set; }
        public string Location { get; set; }
        public string SourceCategory { get; set; }
        public string Publisher { get; set; }
        public string ExecutablePath { get; set; }
        public bool ExecutableExists { get; set; }
        public bool IsDigitallySigned { get; set; }
        public string SignatureSubject { get; set; }
        public bool IsSignatureChainTrusted { get; set; }
        public string SignatureTrustStatus { get; set; }
        public bool IsEnabled { get; set; }
        public bool IsUserLevel { get; set; }

        public string PublisherDisplay => string.IsNullOrWhiteSpace(Publisher) ? "未知发布者" : Publisher;
        public string ExecutablePathDisplay => string.IsNullOrWhiteSpace(ExecutablePath) ? "未识别路径" : ExecutablePath;
        public bool HasExecutablePath => !string.IsNullOrWhiteSpace(ExecutablePath);
        public string ExecutableStatusDisplay => !HasExecutablePath ? "未识别路径" : ExecutableExists ? "文件存在" : "文件不存在";
        public string SourceLocationDisplay => string.IsNullOrWhiteSpace(SourceCategory)
            ? Location
            : SourceCategory + " · " + Location;
        public string SignatureStatusDisplay => !ExecutableExists ? "未检测签名" : IsDigitallySigned ? "已签名" : "未签名";
        public string SignatureDetailDisplay
        {
            get
            {
                if (!ExecutableExists)
                {
                    return "未检测签名；" + SignatureTrustStatus;
                }

                if (!IsDigitallySigned)
                {
                    return SignatureStatusDisplay + "；" + SignatureTrustStatus;
                }

                var subject = string.IsNullOrWhiteSpace(SignatureSubject) ? "未知证书主体" : SignatureSubject;
                var trust = string.IsNullOrWhiteSpace(SignatureTrustStatus) ? "未校验证书链" : SignatureTrustStatus;
                return SignatureStatusDisplay + "：" + subject + "；" + trust;
            }
        }
    }

    public static class StartupService
    {
        private const string RunKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string StartupApprovedRunKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";
        private const string BackupKeyPath = @"SOFTWARE\AJunTools\DisabledRun";
        private const byte StartupApprovedEnabledState = 0x02;
        private const byte StartupApprovedDisabledState = 0x03;

        public static List<StartupItem> GetStartupItems()
        {
            var items = new List<StartupItem>();
            var seenItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            AddItemsFromKey(Registry.CurrentUser, RunKeyPath, items, true, true, seenItems);
            AddItemsFromKey(Registry.LocalMachine, RunKeyPath, items, true, false, seenItems);

            AddItemsFromKey(Registry.CurrentUser, BackupKeyPath, items, false, true, seenItems);
            AddItemsFromKey(Registry.LocalMachine, BackupKeyPath, items, false, false, seenItems);

            return items;
        }

        private static void AddItemsFromKey(RegistryKey root, string path, List<StartupItem> list, bool isEnabled, bool isUserLevel, HashSet<string> seenItems)
        {
            using (var key = root.OpenSubKey(path))
            {
                if (key != null)
                {
                    foreach (var name in key.GetValueNames())
                    {
                        if (!seenItems.Add(BuildItemIdentity(isUserLevel, name)))
                        {
                            continue;
                        }

                        var command = key.GetValue(name)?.ToString();
                        var executablePath = TryExtractExecutablePath(command);
                        var executableExists = !string.IsNullOrWhiteSpace(executablePath) && File.Exists(executablePath);
                        var signatureInfo = ReadSignatureInfo(executablePath);
                        var isStartupApprovedDisabled = isEnabled && IsStartupApprovedDisabled(root, name);
                        var effectiveIsEnabled = isEnabled && !isStartupApprovedDisabled;
                        list.Add(new StartupItem
                        {
                            Name = name,
                            Command = command,
                            Location = isUserLevel ? "当前用户" : "所有用户",
                            SourceCategory = isEnabled
                                ? isStartupApprovedDisabled ? "注册表 Run（Windows 已禁用）" : "注册表 Run"
                                : "MyTools 禁用备份",
                            Publisher = ReadPublisher(executablePath),
                            ExecutablePath = executablePath,
                            ExecutableExists = executableExists,
                            IsDigitallySigned = signatureInfo.IsSigned,
                            SignatureSubject = signatureInfo.Subject,
                            IsSignatureChainTrusted = signatureInfo.IsChainTrusted,
                            SignatureTrustStatus = signatureInfo.TrustStatus,
                            IsEnabled = effectiveIsEnabled,
                            IsUserLevel = isUserLevel
                        });
                    }
                }
            }
        }

        public static void ToggleStartupItem(StartupItem item)
        {
            try
            {
                if (item == null)
                {
                    return;
                }

                if (item.IsEnabled)
                {
                    DisableStartupItem(item);
                }
                else
                {
                    EnableStartupItem(item);
                }
            }
            catch (UnauthorizedAccessException)
            {
                System.Windows.MessageBox.Show(
                    "当前程序未以管理员身份运行，无法修改系统级（所有用户）启动项。\n请关闭程序后，右键以「管理员身份运行」阿君的工具，再执行此操作。",
                    "权限不足",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "ToggleStartupItem failed for {Name}", item.Name);
                System.Windows.MessageBox.Show(
                    "切换启动项状态时发生错误：" + ex.Message,
                    "操作失败",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        public static void DeleteStartupItem(StartupItem item)
        {
            try
            {
                if (item == null)
                {
                    return;
                }

                var root = item.IsUserLevel ? Registry.CurrentUser : Registry.LocalMachine;
                DeleteRegistryValue(root, RunKeyPath, item.Name);
                DeleteRegistryValue(root, BackupKeyPath, item.Name);
                DeleteRegistryValue(root, StartupApprovedRunKeyPath, item.Name);
            }
            catch (UnauthorizedAccessException)
            {
                System.Windows.MessageBox.Show(
                    "当前程序未以管理员身份运行，无法删除系统级（所有用户）启动项。\n请关闭程序后，右键以「管理员身份运行」阿君的工具，再执行此操作。",
                    "权限不足",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "DeleteStartupItem failed for {Name}", item.Name);
                System.Windows.MessageBox.Show(
                    "删除启动项时发生错误：" + ex.Message,
                    "操作失败",
                    System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            }
        }

        private static void DisableStartupItem(StartupItem item)
        {
            var root = item.IsUserLevel ? Registry.CurrentUser : Registry.LocalMachine;
            var valueKind = RegistryValueKind.String;
            object value = null;

            using (var runKey = root.OpenSubKey(RunKeyPath, true))
            {
                value = ReadRegistryValue(runKey, item.Name, out valueKind);
                if (value == null && !string.IsNullOrWhiteSpace(item.Command))
                {
                    value = item.Command;
                }

                if (value != null)
                {
                    using (var backupKey = root.CreateSubKey(BackupKeyPath))
                    {
                        backupKey?.SetValue(item.Name, value, valueKind);
                    }
                }

                SetStartupApprovedState(root, item.Name, false);

                if (value != null)
                {
                    runKey?.DeleteValue(item.Name, false);
                }
            }
        }

        private static void EnableStartupItem(StartupItem item)
        {
            var root = item.IsUserLevel ? Registry.CurrentUser : Registry.LocalMachine;
            var valueKind = RegistryValueKind.String;
            object value;

            using (var backupKey = root.OpenSubKey(BackupKeyPath, true))
            using (var runKey = root.CreateSubKey(RunKeyPath))
            {
                value = ReadRegistryValue(backupKey, item.Name, out valueKind);
                if (value == null)
                {
                    value = ReadRegistryValue(runKey, item.Name, out valueKind);
                }

                if (value == null && !string.IsNullOrWhiteSpace(item.Command))
                {
                    value = item.Command;
                }

                if (value == null)
                {
                    return;
                }

                SetStartupApprovedState(root, item.Name, true);
                runKey?.SetValue(item.Name, value, valueKind);
                backupKey?.DeleteValue(item.Name, false);
            }
        }

        private static object ReadRegistryValue(RegistryKey key, string name, out RegistryValueKind valueKind)
        {
            valueKind = RegistryValueKind.String;
            if (key == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            try
            {
                valueKind = key.GetValueKind(name);
                return key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            }
            catch (IOException)
            {
                return null;
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private static void DeleteRegistryValue(RegistryKey root, string path, string name)
        {
            using (var key = root.OpenSubKey(path, true))
            {
                key?.DeleteValue(name, false);
            }
        }

        private static bool IsStartupApprovedDisabled(RegistryKey root, string name)
        {
            using (var key = root.OpenSubKey(StartupApprovedRunKeyPath))
            {
                var value = key?.GetValue(name) as byte[];
                return value != null && value.Length > 0 && value[0] == StartupApprovedDisabledState;
            }
        }

        private static void SetStartupApprovedState(RegistryKey root, string name, bool enabled)
        {
            using (var key = root.CreateSubKey(StartupApprovedRunKeyPath))
            {
                key?.SetValue(name, CreateStartupApprovedValue(enabled), RegistryValueKind.Binary);
            }
        }

        private static byte[] CreateStartupApprovedValue(bool enabled)
        {
            var value = new byte[12];
            value[0] = enabled ? StartupApprovedEnabledState : StartupApprovedDisabledState;

            if (!enabled)
            {
                var fileTime = BitConverter.GetBytes(DateTime.UtcNow.ToFileTimeUtc());
                Buffer.BlockCopy(fileTime, 0, value, 4, fileTime.Length);
            }

            return value;
        }

        private static string BuildItemIdentity(bool isUserLevel, string name)
        {
            return (isUserLevel ? "HKCU\\" : "HKLM\\") + (name ?? string.Empty);
        }

        private static string TryExtractExecutablePath(string command)
        {
            if (string.IsNullOrWhiteSpace(command))
            {
                return string.Empty;
            }

            var value = Environment.ExpandEnvironmentVariables(command.Trim().TrimStart('@'));
            string candidate = string.Empty;
            if (value.StartsWith("\"", StringComparison.Ordinal))
            {
                var closingQuote = value.IndexOf('"', 1);
                if (closingQuote > 1)
                {
                    candidate = value.Substring(1, closingQuote - 1);
                }
            }

            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = ExtractPathByKnownExtension(value);
            }

            if (string.IsNullOrWhiteSpace(candidate))
            {
                var firstToken = value.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                candidate = firstToken ?? string.Empty;
            }

            candidate = candidate.Trim('"', ' ');
            if (File.Exists(candidate))
            {
                return candidate;
            }

            var resolved = ResolveFromPath(candidate);
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                return resolved;
            }

            return LooksLikeExecutablePath(candidate) ? candidate : string.Empty;
        }

        private static string ExtractPathByKnownExtension(string value)
        {
            var bestIndex = -1;
            var bestLength = 0;
            foreach (var extension in new[] { ".exe", ".bat", ".cmd", ".com", ".ps1" })
            {
                var index = value.IndexOf(extension, StringComparison.OrdinalIgnoreCase);
                if (index >= 0 && (bestIndex < 0 || index < bestIndex))
                {
                    bestIndex = index;
                    bestLength = extension.Length;
                }
            }

            return bestIndex < 0 ? string.Empty : value.Substring(0, bestIndex + bestLength);
        }

        private static string ResolveFromPath(string executableName)
        {
            if (string.IsNullOrWhiteSpace(executableName) || executableName.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            {
                return string.Empty;
            }

            if (Path.IsPathRooted(executableName))
            {
                return string.Empty;
            }

            foreach (var directory in new[]
            {
                Environment.SystemDirectory,
                Environment.GetFolderPath(Environment.SpecialFolder.Windows)
            })
            {
                try
                {
                    var candidate = Path.Combine(directory, executableName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                }
            }

            var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var directory in pathValue.Split(new[] { Path.PathSeparator }, StringSplitOptions.RemoveEmptyEntries))
            {
                try
                {
                    var candidate = Path.Combine(directory.Trim(), executableName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        private static bool LooksLikeExecutablePath(string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate) || candidate.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            {
                return false;
            }

            var extension = Path.GetExtension(candidate);
            if (!new[] { ".exe", ".bat", ".cmd", ".com", ".ps1" }.Any(x => string.Equals(x, extension, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            return Path.IsPathRooted(candidate)
                || candidate.IndexOf('\\') >= 0
                || candidate.IndexOf('/') >= 0
                || !string.IsNullOrWhiteSpace(extension);
        }

        private static string ReadPublisher(string executablePath)
        {
            if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            {
                return string.Empty;
            }

            try
            {
                var versionInfo = FileVersionInfo.GetVersionInfo(executablePath);
                return string.IsNullOrWhiteSpace(versionInfo.CompanyName)
                    ? string.Empty
                    : versionInfo.CompanyName.Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static SignatureInfo ReadSignatureInfo(string executablePath)
        {
            if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath))
            {
                return SignatureInfo.Empty;
            }

            try
            {
                using (var certificate = X509Certificate.CreateFromSignedFile(executablePath))
                using (var certificate2 = new X509Certificate2(certificate))
                {
                    var subject = certificate2.GetNameInfo(X509NameType.SimpleName, false);
                    if (string.IsNullOrWhiteSpace(subject))
                    {
                        subject = certificate2.Subject;
                    }

                    var isTrusted = IsCertificateChainTrusted(certificate2);
                    return new SignatureInfo
                    {
                        IsSigned = true,
                        Subject = subject ?? string.Empty,
                        IsChainTrusted = isTrusted,
                        TrustStatus = isTrusted ? "证书链可信" : "证书链不可信或已过期"
                    };
                }
            }
            catch
            {
                return SignatureInfo.Empty;
            }
        }

        private static bool IsCertificateChainTrusted(X509Certificate2 certificate)
        {
            try
            {
                using (var chain = new X509Chain())
                {
                    chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                    chain.ChainPolicy.VerificationFlags = X509VerificationFlags.NoFlag;
                    return chain.Build(certificate);
                }
            }
            catch
            {
                return false;
            }
        }

        private sealed class SignatureInfo
        {
            public static readonly SignatureInfo Empty = new SignatureInfo
            {
                TrustStatus = "未校验证书链"
            };

            public bool IsSigned { get; set; }
            public string Subject { get; set; } = string.Empty;
            public bool IsChainTrusted { get; set; }
            public string TrustStatus { get; set; } = string.Empty;
        }
    }
}
