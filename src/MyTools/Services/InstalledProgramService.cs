using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Win32;

namespace MyTools.Services
{
    public sealed class InstalledProgram : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string DisplayName { get; set; }
        public string DisplayVersion { get; set; }
        public string Publisher { get; set; }
        public string InstallLocation { get; set; }
        public DateTime? InstallDate { get; set; }
        public string InstallDateDisplay { get; set; }
        public int EstimatedSizeKb { get; set; }
        public string EstimatedSizeDisplay { get; set; }
        public string UninstallString { get; set; }
        public string QuietUninstallString { get; set; }
        public bool IsSilentUninstallCandidate { get; set; }
        public string SilentUninstallReason { get; set; }
        public bool RequiresAdmin { get; set; }
        public string Source { get; set; }

        public string PublisherDisplay => string.IsNullOrWhiteSpace(Publisher) ? "未知发布者" : Publisher;
        public string VersionDisplay => string.IsNullOrWhiteSpace(DisplayVersion) ? "-" : DisplayVersion;
        public string InstallLocationDisplay => string.IsNullOrWhiteSpace(InstallLocation) ? "-" : InstallLocation;
        public string RequiresAdminDisplay => RequiresAdmin ? "可能需要管理员权限" : "当前用户可卸载";
        public string SilentUninstallDisplay => IsSilentUninstallCandidate ? "可静默候选" : "需交互卸载";
        public string SilentUninstallDetail => IsSilentUninstallCandidate
            ? (string.IsNullOrWhiteSpace(SilentUninstallReason) ? "检测到可能可静默卸载的命令。" : SilentUninstallReason)
            : "未发现可靠静默卸载参数；MyTools 默认只启动交互卸载向导。";

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public static class InstalledProgramService
    {
        private const string UninstallKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
        private const string Wow6432UninstallKeyPath = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";

        public static List<InstalledProgram> GetUninstallablePrograms()
        {
            var items = new List<InstalledProgram>();

            AddProgramsFromKey(Registry.CurrentUser, UninstallKeyPath, "当前用户", false, items);
            AddProgramsFromKey(Registry.LocalMachine, UninstallKeyPath, "所有用户 64 位", true, items);
            AddProgramsFromKey(Registry.LocalMachine, Wow6432UninstallKeyPath, "所有用户 32 位", true, items);

            return items
                .GroupBy(BuildDedupeKey, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public static Process StartUninstall(InstalledProgram program)
        {
            if (program == null)
            {
                throw new ArgumentNullException(nameof(program));
            }

            if (string.IsNullOrWhiteSpace(program.UninstallString))
            {
                throw new InvalidOperationException("该程序没有提供卸载命令。");
            }

            var startInfo = BuildStartInfo(program.UninstallString);
            AppLogService.Information(
                "Starting uninstall for {ProgramName}, publisher {Publisher}, source {Source}",
                program.DisplayName ?? string.Empty,
                program.Publisher ?? string.Empty,
                program.Source ?? string.Empty);

            var process = Process.Start(startInfo);
            if (process == null)
            {
                throw new InvalidOperationException("无法启动卸载程序。");
            }
            return process;
        }

        /// <summary>
        /// 重新扫描注册表，判断该程序是否仍存在。比较 DisplayName + UninstallString 双键。
        /// </summary>
        public static bool IsStillInstalled(InstalledProgram program)
        {
            if (program == null) return false;
            var all = GetUninstallablePrograms();
            foreach (var p in all)
            {
                if (string.Equals(p.DisplayName, program.DisplayName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.UninstallString, program.UninstallString, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static void AddProgramsFromKey(
            RegistryKey root,
            string path,
            string source,
            bool requiresAdmin,
            ICollection<InstalledProgram> items)
        {
            try
            {
                using (var key = root.OpenSubKey(path))
                {
                    if (key == null)
                    {
                        return;
                    }

                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        using (var subKey = key.OpenSubKey(subKeyName))
                        {
                            var item = ReadProgram(subKey, source, requiresAdmin);
                            if (item != null)
                            {
                                items.Add(item);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading uninstall registry key failed for {Path}", path);
            }
        }

        private static InstalledProgram ReadProgram(RegistryKey key, string source, bool requiresAdmin)
        {
            if (key == null)
            {
                return null;
            }

            var displayName = ReadString(key, "DisplayName");
            var uninstallString = ReadString(key, "UninstallString");
            if (string.IsNullOrWhiteSpace(displayName) || string.IsNullOrWhiteSpace(uninstallString))
            {
                return null;
            }

            if (ReadDword(key, "SystemComponent") == 1
                || ReadDword(key, "NoRemove") == 1
                || ReadDword(key, "WindowsInstaller") == 1 && string.IsNullOrWhiteSpace(uninstallString))
            {
                return null;
            }

            var installDateRaw = ReadString(key, "InstallDate");
            var estimatedSizeKb = ReadDword(key, "EstimatedSize");
            var quietUninstallString = ReadString(key, "QuietUninstallString");
            var silentInfo = DetectSilentUninstallCandidate(uninstallString, quietUninstallString);

            return new InstalledProgram
            {
                DisplayName = displayName.Trim(),
                DisplayVersion = ReadString(key, "DisplayVersion"),
                Publisher = ReadString(key, "Publisher"),
                InstallLocation = ReadString(key, "InstallLocation"),
                InstallDate = ParseInstallDate(installDateRaw),
                InstallDateDisplay = FormatInstallDate(installDateRaw),
                EstimatedSizeKb = estimatedSizeKb,
                EstimatedSizeDisplay = FormatEstimatedSize(estimatedSizeKb),
                UninstallString = uninstallString.Trim(),
                QuietUninstallString = quietUninstallString.Trim(),
                IsSilentUninstallCandidate = silentInfo.IsCandidate,
                SilentUninstallReason = silentInfo.Reason,
                RequiresAdmin = requiresAdmin,
                Source = source
            };
        }

        private static string ReadString(RegistryKey key, string name)
        {
            return key.GetValue(name) as string ?? string.Empty;
        }

        private static int ReadDword(RegistryKey key, string name)
        {
            var value = key.GetValue(name);
            if (value is int intValue)
            {
                return intValue;
            }

            return 0;
        }

        private static DateTime? ParseInstallDate(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            if (DateTime.TryParseExact(
                rawValue,
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var exactDate))
            {
                return exactDate.Date;
            }

            if (DateTime.TryParse(rawValue, CultureInfo.CurrentCulture, DateTimeStyles.None, out var looseDate))
            {
                return looseDate.Date;
            }

            return null;
        }

        private static string FormatInstallDate(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return "-";
            }

            if (DateTime.TryParseExact(
                rawValue,
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date))
            {
                return date.ToString("yyyy-MM-dd", CultureInfo.CurrentCulture);
            }

            return rawValue;
        }

        private static string FormatEstimatedSize(int estimatedSizeKb)
        {
            if (estimatedSizeKb <= 0)
            {
                return "-";
            }

            var sizeMb = estimatedSizeKb / 1024d;
            return sizeMb >= 1024
                ? string.Format(CultureInfo.CurrentCulture, "{0:F1} GB", sizeMb / 1024d)
                : string.Format(CultureInfo.CurrentCulture, "{0:F0} MB", sizeMb);
        }

        private static string BuildDedupeKey(InstalledProgram item)
        {
            return string.Join(
                "|",
                item.DisplayName ?? string.Empty,
                item.DisplayVersion ?? string.Empty,
                item.Publisher ?? string.Empty,
                item.UninstallString ?? string.Empty);
        }

        private static SilentUninstallInfo DetectSilentUninstallCandidate(string uninstallString, string quietUninstallString)
        {
            if (!string.IsNullOrWhiteSpace(quietUninstallString))
            {
                return new SilentUninstallInfo
                {
                    IsCandidate = true,
                    Reason = "注册表提供 QuietUninstallString，适合人工确认后静默执行。"
                };
            }

            var command = uninstallString ?? string.Empty;
            if (command.IndexOf("msiexec", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new SilentUninstallInfo
                {
                    IsCandidate = true,
                    Reason = "MSI 卸载命令可追加 /qn /norestart 作为静默候选。"
                };
            }

            if (ContainsSilentToken(command))
            {
                return new SilentUninstallInfo
                {
                    IsCandidate = true,
                    Reason = "卸载命令已包含常见静默参数，仍需人工确认。"
                };
            }

            return SilentUninstallInfo.None;
        }

        private static bool ContainsSilentToken(string command)
        {
            if (string.IsNullOrWhiteSpace(command))
            {
                return false;
            }

            var normalized = " " + command.Replace('\t', ' ').Trim().ToLowerInvariant() + " ";
            return normalized.Contains(" /quiet ")
                || normalized.Contains(" /qn ")
                || normalized.Contains(" /s ")
                || normalized.Contains(" /silent ")
                || normalized.Contains(" /verysilent ")
                || normalized.Contains(" -quiet ")
                || normalized.Contains(" -silent ")
                || normalized.Contains(" --silent ");
        }

        private static ProcessStartInfo BuildStartInfo(string commandLine)
        {
            var trimmed = commandLine.Trim();
            if (TryBuildMsiStartInfo(trimmed, out var msiStartInfo))
            {
                return msiStartInfo;
            }

            if (TrySplitExecutableCommand(trimmed, out var fileName, out var arguments))
            {
                if (IsMsiExec(fileName))
                {
                    return new ProcessStartInfo
                    {
                        FileName = fileName,
                        Arguments = NormalizeMsiArguments(arguments),
                        UseShellExecute = true,
                        WorkingDirectory = GetWorkingDirectory(fileName),
                        WindowStyle = ProcessWindowStyle.Normal
                    };
                }

                return new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = true,
                    WorkingDirectory = GetWorkingDirectory(fileName),
                    WindowStyle = ProcessWindowStyle.Normal
                };
            }

            return new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c " + trimmed,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Normal
            };
        }

        private static bool TryBuildMsiStartInfo(string commandLine, out ProcessStartInfo startInfo)
        {
            startInfo = null;
            if (!commandLine.StartsWith("msiexec", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var firstSpace = commandLine.IndexOf(' ');
            var arguments = firstSpace >= 0 ? commandLine.Substring(firstSpace + 1).Trim() : string.Empty;
            startInfo = new ProcessStartInfo
            {
                FileName = "msiexec.exe",
                Arguments = NormalizeMsiArguments(arguments),
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Normal
            };
            return true;
        }

        private static bool IsMsiExec(string fileName)
        {
            return string.Equals(Path.GetFileName(fileName), "msiexec.exe", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeMsiArguments(string arguments)
        {
            var trimmed = (arguments ?? string.Empty).Trim();
            if (trimmed.StartsWith("/I", StringComparison.OrdinalIgnoreCase))
            {
                return "/X" + trimmed.Substring(2);
            }

            if (trimmed.StartsWith("-I", StringComparison.OrdinalIgnoreCase))
            {
                return "/X" + trimmed.Substring(2);
            }

            return trimmed;
        }

        private static bool TrySplitExecutableCommand(string commandLine, out string fileName, out string arguments)
        {
            fileName = string.Empty;
            arguments = string.Empty;

            if (string.IsNullOrWhiteSpace(commandLine))
            {
                return false;
            }

            if (commandLine[0] == '"')
            {
                var closingQuote = commandLine.IndexOf('"', 1);
                if (closingQuote <= 1)
                {
                    return false;
                }

                fileName = commandLine.Substring(1, closingQuote - 1);
                arguments = commandLine.Substring(closingQuote + 1).Trim();
                return true;
            }

            var exeIndex = commandLine.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            if (exeIndex < 0)
            {
                return false;
            }

            fileName = commandLine.Substring(0, exeIndex + 4).Trim();
            arguments = commandLine.Substring(exeIndex + 4).Trim();
            return true;
        }

        private static string GetWorkingDirectory(string fileName)
        {
            try
            {
                var directory = Path.GetDirectoryName(fileName);
                return Directory.Exists(directory) ? directory : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private sealed class SilentUninstallInfo
        {
            public static readonly SilentUninstallInfo None = new SilentUninstallInfo();

            public bool IsCandidate { get; set; }
            public string Reason { get; set; } = string.Empty;
        }
    }
}
