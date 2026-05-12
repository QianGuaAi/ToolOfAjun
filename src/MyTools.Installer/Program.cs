using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace MyTools.Installer
{
    internal static class Program
    {
        private const string ProductName = "阿君的工具";
        private const string ProductKey = "MyTools";
        private const string Publisher = "Ajun";
        private const string PayloadResourceName = "MyToolsPayload.zip";
        private const string UninstallerResourceName = "MyTools.Uninstaller.exe";
        private const string AppExeName = "MyTools.exe";
        private const string UninstallerExeName = "MyTools.Uninstaller.exe";
        private static readonly string[] InstalledFiles =
        {
            AppExeName,
            "MyTools.exe.config",
            "LockWin10_22H2.ps1",
            UninstallerExeName
        };

        private static readonly string[] InstalledDirectories =
        {
            "NativeBinaries",
            "x64",
            "x86"
        };

        [STAThread]
        private static int Main(string[] args)
        {
            var options = Options.Parse(args);

            try
            {
                if (options.ExtractDirectory != null)
                {
                    ExtractPayload(options.ExtractDirectory);
                    return 0;
                }

                var installDirectory = options.InstallDirectory ?? GetDefaultInstallDirectory();
                installDirectory = Path.GetFullPath(installDirectory);

                RunPreviousUninstaller(installDirectory);
                StopInstalledApplication(installDirectory);
                Directory.CreateDirectory(installDirectory);

                ExtractPayload(installDirectory);
                WriteUninstaller(installDirectory);
                CreateShortcuts(installDirectory);
                WriteUninstallRegistry(installDirectory);

                if (!options.Silent)
                {
                    MessageBox.Show(
                        $"安装完成。\r\n\r\n安装目录：{installDirectory}",
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                if (options.LaunchAfterInstall)
                {
                    Process.Start(Path.Combine(installDirectory, AppExeName));
                }

                return 0;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "安装失败：" + ex.Message,
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                return 1;
            }
        }

        private static string GetDefaultInstallDirectory()
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            return Path.Combine(programFiles, "AjunTools", "MyTools");
        }

        private static void RunPreviousUninstaller(string targetInstallDirectory)
        {
            var entries = FindInstalledProducts()
                .Concat(FindLocalInstalledProduct(targetInstallDirectory))
                .GroupBy(entry => entry.UninstallString ?? entry.InstallLocation ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

            var ranUninstaller = false;
            foreach (var entry in entries)
            {
                var command = BuildUninstallCommand(entry);
                if (command == null)
                {
                    continue;
                }

                var exitCode = RunAndWait(command.FileName, command.Arguments, 300000);
                if (exitCode != 0)
                {
                    throw new InvalidOperationException($"旧版本卸载失败，退出码：{exitCode}");
                }

                ranUninstaller = true;
            }

            if (!ranUninstaller && File.Exists(Path.Combine(targetInstallDirectory, AppExeName)))
            {
                StopInstalledApplication(targetInstallDirectory);
                RemoveKnownInstalledFiles(targetInstallDirectory);
            }

            WaitForDirectoryUnlock(targetInstallDirectory);
        }

        private static IEnumerable<InstalledProduct> FindLocalInstalledProduct(string targetInstallDirectory)
        {
            if (string.IsNullOrWhiteSpace(targetInstallDirectory))
            {
                yield break;
            }

            var uninstallerPath = Path.Combine(targetInstallDirectory, UninstallerExeName);
            if (!File.Exists(uninstallerPath))
            {
                yield break;
            }

            yield return new InstalledProduct
            {
                KeyName = ProductKey,
                DisplayName = ProductName,
                InstallLocation = targetInstallDirectory,
                UninstallString = Quote(uninstallerPath),
                QuietUninstallString = Quote(uninstallerPath) + " /silent"
            };
        }

        private static IEnumerable<InstalledProduct> FindInstalledProducts()
        {
            foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            {
                foreach (var view in GetRegistryViews())
                {
                    foreach (var product in FindInstalledProducts(hive, view))
                    {
                        yield return product;
                    }
                }
            }
        }

        private static IEnumerable<RegistryView> GetRegistryViews()
        {
            yield return RegistryView.Registry32;
            if (Environment.Is64BitOperatingSystem)
            {
                yield return RegistryView.Registry64;
            }
        }

        private static IEnumerable<InstalledProduct> FindInstalledProducts(RegistryHive hive, RegistryView view)
        {
            using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
            using (var uninstallKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", false))
            {
                if (uninstallKey == null)
                {
                    yield break;
                }

                foreach (var subKeyName in uninstallKey.GetSubKeyNames())
                {
                    using (var productKey = uninstallKey.OpenSubKey(subKeyName, false))
                    {
                        if (productKey == null)
                        {
                            continue;
                        }

                        var displayName = Convert.ToString(productKey.GetValue("DisplayName"));
                        if (!IsMyToolsProduct(subKeyName, displayName))
                        {
                            continue;
                        }

                        yield return new InstalledProduct
                        {
                            KeyName = subKeyName,
                            DisplayName = displayName,
                            InstallLocation = Convert.ToString(productKey.GetValue("InstallLocation")),
                            UninstallString = Convert.ToString(productKey.GetValue("UninstallString")),
                            QuietUninstallString = Convert.ToString(productKey.GetValue("QuietUninstallString"))
                        };
                    }
                }
            }
        }

        private static bool IsMyToolsProduct(string keyName, string displayName)
        {
            if (string.Equals(keyName, ProductKey, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            return displayName.IndexOf(ProductName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   displayName.IndexOf("MyTools", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static CommandLine BuildUninstallCommand(InstalledProduct entry)
        {
            if (entry == null)
            {
                return null;
            }

            var command = CommandLine.Parse(entry.QuietUninstallString ?? entry.UninstallString);
            if (command == null || string.IsNullOrWhiteSpace(command.FileName))
            {
                return null;
            }

            var fileName = ResolveExecutable(command.FileName);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            var arguments = command.Arguments ?? string.Empty;
            var executableName = Path.GetFileName(fileName);
            if (executableName.Equals(UninstallerExeName, StringComparison.OrdinalIgnoreCase))
            {
                arguments = EnsureArgument(arguments, "/from-upgrade");
                arguments = EnsureArgument(arguments, "/silent");
                if (!string.IsNullOrWhiteSpace(entry.InstallLocation))
                {
                    arguments = AppendArgument(arguments, "/install-dir");
                    arguments = AppendArgument(arguments, Quote(entry.InstallLocation));
                }
            }
            else if (executableName.Equals("msiexec.exe", StringComparison.OrdinalIgnoreCase) ||
                     executableName.Equals("msiexec", StringComparison.OrdinalIgnoreCase))
            {
                arguments = NormalizeMsiUninstallArguments(arguments);
            }

            return new CommandLine
            {
                FileName = fileName,
                Arguments = arguments
            };
        }

        private static string ResolveExecutable(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            if (Path.IsPathRooted(fileName))
            {
                return File.Exists(fileName) ? fileName : null;
            }

            if (File.Exists(fileName))
            {
                return Path.GetFullPath(fileName);
            }

            var pathEnvironment = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var directory in pathEnvironment.Split(new[] { Path.PathSeparator }, StringSplitOptions.RemoveEmptyEntries))
            {
                try
                {
                    var candidate = Path.Combine(directory.Trim(), fileName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }

                    if (Path.GetExtension(candidate).Length == 0)
                    {
                        var exeCandidate = candidate + ".exe";
                        if (File.Exists(exeCandidate))
                        {
                            return exeCandidate;
                        }
                    }
                }
                catch
                {
                    // Ignore malformed PATH segments.
                }
            }

            return fileName;
        }

        private static string NormalizeMsiUninstallArguments(string arguments)
        {
            arguments = arguments ?? string.Empty;
            arguments = ReplaceMsiInstallVerb(arguments);
            arguments = EnsureArgument(arguments, "/qn");
            arguments = EnsureArgument(arguments, "/norestart");
            return arguments.Trim();
        }

        private static string ReplaceMsiInstallVerb(string arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments))
            {
                return arguments;
            }

            var parts = arguments.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (var i = 0; i < parts.Length; i++)
            {
                if (parts[i].StartsWith("/I", StringComparison.OrdinalIgnoreCase))
                {
                    parts[i] = "/X" + parts[i].Substring(2);
                    return string.Join(" ", parts);
                }
            }

            return arguments;
        }

        private static int RunAndWait(string fileName, string arguments, int timeoutMilliseconds)
        {
            using (var process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments ?? string.Empty,
                UseShellExecute = true,
                Verb = "runas",
                WindowStyle = ProcessWindowStyle.Normal
            }))
            {
                if (process == null)
                {
                    throw new InvalidOperationException("无法启动旧版本卸载程序。");
                }

                if (!process.WaitForExit(timeoutMilliseconds))
                {
                    throw new TimeoutException("旧版本卸载超时。");
                }

                return process.ExitCode;
            }
        }

        private static void WaitForDirectoryUnlock(string installDirectory)
        {
            if (!Directory.Exists(installDirectory))
            {
                return;
            }

            var probePath = Path.Combine(installDirectory, ".install-probe");
            for (var i = 0; i < 20; i++)
            {
                try
                {
                    File.WriteAllText(probePath, string.Empty);
                    File.Delete(probePath);
                    return;
                }
                catch
                {
                    Thread.Sleep(250);
                }
            }
        }

        private static void StopInstalledApplication(string installDirectory)
        {
            var normalizedInstallDirectory = EnsureTrailingSeparator(Path.GetFullPath(installDirectory));
            foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(AppExeName)))
            {
                using (process)
                {
                    try
                    {
                        var processPath = process.MainModule == null ? string.Empty : process.MainModule.FileName;
                        if (string.IsNullOrWhiteSpace(processPath) ||
                            !Path.GetFullPath(processPath).StartsWith(normalizedInstallDirectory, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        process.CloseMainWindow();
                        if (!process.WaitForExit(5000))
                        {
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                    }
                    catch
                    {
                        // Ignore protected or already-exited processes.
                    }
                }
            }
        }

        private static void RemoveKnownInstalledFiles(string installDirectory)
        {
            foreach (var fileName in InstalledFiles)
            {
                TryDeleteFile(Path.Combine(installDirectory, fileName));
            }

            foreach (var directoryName in InstalledDirectories)
            {
                TryDeleteDirectory(Path.Combine(installDirectory, directoryName), true);
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.SetAttributes(path, FileAttributes.Normal);
                    File.Delete(path);
                }
            }
            catch
            {
                // Leave locked files in place; the subsequent extract step may still overwrite newer payload files.
            }
        }

        private static void TryDeleteDirectory(string path, bool recursive)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, recursive);
                }
            }
            catch
            {
                // User data and locked files are intentionally preserved.
            }
        }

        private static void ExtractPayload(string installDirectory)
        {
            Directory.CreateDirectory(installDirectory);
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(PayloadResourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("安装包中缺少程序文件载荷，请重新构建安装程序。");
                }

                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    foreach (var entry in archive.Entries)
                    {
                        var destinationPath = Path.GetFullPath(Path.Combine(installDirectory, entry.FullName));
                        if (!destinationPath.StartsWith(EnsureTrailingSeparator(Path.GetFullPath(installDirectory)), StringComparison.OrdinalIgnoreCase))
                        {
                            throw new InvalidOperationException("安装包载荷包含非法路径。");
                        }

                        if (string.IsNullOrEmpty(entry.Name))
                        {
                            Directory.CreateDirectory(destinationPath);
                            continue;
                        }

                        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                        entry.ExtractToFile(destinationPath, true);
                    }
                }
            }
        }

        private static void WriteUninstaller(string installDirectory)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(UninstallerResourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("安装包中缺少卸载程序，请重新构建安装程序。");
                }

                var destinationPath = Path.Combine(installDirectory, UninstallerExeName);
                using (var file = File.Create(destinationPath))
                {
                    stream.CopyTo(file);
                }
            }
        }

        private static void CreateShortcuts(string installDirectory)
        {
            var appPath = Path.Combine(installDirectory, AppExeName);
            var uninstallerPath = Path.Combine(installDirectory, UninstallerExeName);
            var programsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms), ProductName);
            Directory.CreateDirectory(programsDirectory);

            CreateShortcut(
                Path.Combine(programsDirectory, ProductName + ".lnk"),
                appPath,
                string.Empty,
                installDirectory,
                "打开阿君的工具");

            CreateShortcut(
                Path.Combine(programsDirectory, "卸载 " + ProductName + ".lnk"),
                uninstallerPath,
                string.Empty,
                installDirectory,
                "卸载阿君的工具");
        }

        private static void CreateShortcut(string shortcutPath, string targetPath, string arguments, string workingDirectory, string description)
        {
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                return;
            }

            dynamic shell = Activator.CreateInstance(shellType);
            dynamic shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = targetPath;
            shortcut.Arguments = arguments;
            shortcut.WorkingDirectory = workingDirectory;
            shortcut.Description = description;
            shortcut.IconLocation = targetPath + ",0";
            shortcut.Save();
        }

        private static void WriteUninstallRegistry(string installDirectory)
        {
            var appPath = Path.Combine(installDirectory, AppExeName);
            var uninstallerPath = Path.Combine(installDirectory, UninstallerExeName);
            var displayVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, GetWritableRegistryView()))
            using (var uninstallKey = baseKey.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + ProductKey))
            {
                if (uninstallKey == null)
                {
                    throw new InvalidOperationException("无法写入卸载注册表项。");
                }

                uninstallKey.SetValue("DisplayName", ProductName, RegistryValueKind.String);
                uninstallKey.SetValue("DisplayVersion", displayVersion, RegistryValueKind.String);
                uninstallKey.SetValue("Publisher", Publisher, RegistryValueKind.String);
                uninstallKey.SetValue("InstallLocation", installDirectory, RegistryValueKind.String);
                uninstallKey.SetValue("DisplayIcon", appPath, RegistryValueKind.String);
                uninstallKey.SetValue("UninstallString", Quote(uninstallerPath), RegistryValueKind.String);
                uninstallKey.SetValue("QuietUninstallString", Quote(uninstallerPath) + " /silent", RegistryValueKind.String);
                uninstallKey.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"), RegistryValueKind.String);
                uninstallKey.SetValue("NoModify", 1, RegistryValueKind.DWord);
                uninstallKey.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                uninstallKey.SetValue("EstimatedSize", CalculateEstimatedSizeKb(installDirectory), RegistryValueKind.DWord);
            }
        }

        private static RegistryView GetWritableRegistryView()
        {
            return Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
        }

        private static int CalculateEstimatedSizeKb(string installDirectory)
        {
            try
            {
                var totalBytes = Directory.GetFiles(installDirectory, "*", SearchOption.AllDirectories)
                    .Sum(path => new FileInfo(path).Length);
                return (int)Math.Min(int.MaxValue, Math.Max(1, totalBytes / 1024));
            }
            catch
            {
                return 1;
            }
        }

        private static string AppendArgument(string arguments, string argument)
        {
            return string.IsNullOrWhiteSpace(arguments)
                ? argument
                : arguments + " " + argument;
        }

        private static string EnsureArgument(string arguments, string argument)
        {
            if (ContainsArgument(arguments, argument))
            {
                return arguments;
            }

            return AppendArgument(arguments, argument);
        }

        private static bool ContainsArgument(string arguments, string argument)
        {
            if (string.IsNullOrWhiteSpace(arguments))
            {
                return false;
            }

            return arguments.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Any(part => string.Equals(part, argument, StringComparison.OrdinalIgnoreCase));
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }

        private static string EnsureTrailingSeparator(string path)
        {
            if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        private sealed class InstalledProduct
        {
            public string KeyName { get; set; }
            public string DisplayName { get; set; }
            public string InstallLocation { get; set; }
            public string UninstallString { get; set; }
            public string QuietUninstallString { get; set; }
        }

        private sealed class CommandLine
        {
            public string FileName { get; set; }
            public string Arguments { get; set; }

            public static CommandLine Parse(string commandLine)
            {
                if (string.IsNullOrWhiteSpace(commandLine))
                {
                    return null;
                }

                commandLine = commandLine.Trim();
                if (commandLine.StartsWith("\"", StringComparison.Ordinal))
                {
                    var endQuote = commandLine.IndexOf('"', 1);
                    if (endQuote <= 1)
                    {
                        return null;
                    }

                    return new CommandLine
                    {
                        FileName = commandLine.Substring(1, endQuote - 1),
                        Arguments = commandLine.Substring(endQuote + 1).Trim()
                    };
                }

                var exeIndex = commandLine.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
                if (exeIndex >= 0)
                {
                    var fileNameEnd = exeIndex + 4;
                    return new CommandLine
                    {
                        FileName = commandLine.Substring(0, fileNameEnd),
                        Arguments = commandLine.Substring(fileNameEnd).Trim()
                    };
                }

                var firstSpace = commandLine.IndexOf(' ');
                if (firstSpace < 0)
                {
                    return new CommandLine { FileName = commandLine, Arguments = string.Empty };
                }

                return new CommandLine
                {
                    FileName = commandLine.Substring(0, firstSpace),
                    Arguments = commandLine.Substring(firstSpace + 1).Trim()
                };
            }
        }

        private sealed class Options
        {
            public bool Silent { get; private set; }
            public bool LaunchAfterInstall { get; private set; }
            public string InstallDirectory { get; private set; }
            public string ExtractDirectory { get; private set; }

            public static Options Parse(string[] args)
            {
                var options = new Options();

                for (var i = 0; i < args.Length; i++)
                {
                    var arg = args[i];
                    if (EqualsArg(arg, "/silent") || EqualsArg(arg, "-silent") || EqualsArg(arg, "--silent"))
                    {
                        options.Silent = true;
                    }
                    else if (EqualsArg(arg, "/launch") || EqualsArg(arg, "--launch"))
                    {
                        options.LaunchAfterInstall = true;
                    }
                    else if (StartsWithArg(arg, "/install-dir:") || StartsWithArg(arg, "--install-dir="))
                    {
                        options.InstallDirectory = arg.Substring(arg.IndexOfAny(new[] { ':', '=' }) + 1).Trim('"');
                    }
                    else if ((EqualsArg(arg, "/install-dir") || EqualsArg(arg, "--install-dir")) && i + 1 < args.Length)
                    {
                        options.InstallDirectory = args[++i].Trim('"');
                    }
                    else if (StartsWithArg(arg, "/extract:") || StartsWithArg(arg, "--extract="))
                    {
                        options.ExtractDirectory = arg.Substring(arg.IndexOfAny(new[] { ':', '=' }) + 1).Trim('"');
                    }
                    else if ((EqualsArg(arg, "/extract") || EqualsArg(arg, "--extract")) && i + 1 < args.Length)
                    {
                        options.ExtractDirectory = args[++i].Trim('"');
                    }
                }

                return options;
            }

            private static bool EqualsArg(string value, string expected)
            {
                return string.Equals(value, expected, StringComparison.OrdinalIgnoreCase);
            }

            private static bool StartsWithArg(string value, string expectedPrefix)
            {
                return value.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
