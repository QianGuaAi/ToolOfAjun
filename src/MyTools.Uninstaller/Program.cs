using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace MyTools.Uninstaller
{
    internal static class Program
    {
        private const string ProductName = "阿君的工具";
        private const string ProductKey = "MyTools";
        private const string AppExeName = "MyTools.exe";
        private const string UninstallerExeName = "MyTools.Uninstaller.exe";

        private static readonly string[] InstalledFiles =
        {
            AppExeName,
            "MyTools.exe.config",
            "LockWin10_22H2.ps1"
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
                var installDirectory = ResolveInstallDirectory(options.InstallDirectory);
                if (string.IsNullOrWhiteSpace(installDirectory) || !Directory.Exists(installDirectory))
                {
                    RemoveUninstallRegistry();
                    return 0;
                }

                StopInstalledApplication(installDirectory);
                RemoveShortcuts();
                RemoveUninstallRegistry();
                RemoveInstalledFiles(installDirectory, options.PurgeData);

                if (!options.FromUpgrade)
                {
                    ScheduleSelfRemoval(installDirectory, options.PurgeData);
                }

                if (!options.Silent)
                {
                    MessageBox.Show(
                        options.PurgeData
                            ? "阿君的工具已卸载。"
                            : "阿君的工具已卸载。配置、日志等用户数据会保留在安装目录中。",
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return 0;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "卸载失败：" + ex.Message,
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                return 1;
            }
        }

        private static string ResolveInstallDirectory(string requestedDirectory)
        {
            if (!string.IsNullOrWhiteSpace(requestedDirectory))
            {
                return Path.GetFullPath(requestedDirectory);
            }

            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            return Path.GetFullPath(baseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
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
                        // Ignore protected processes and continue with uninstalling files that are not locked.
                    }
                }
            }
        }

        private static void RemoveInstalledFiles(string installDirectory, bool purgeData)
        {
            if (purgeData)
            {
                return;
            }

            foreach (var fileName in InstalledFiles)
            {
                TryDeleteFile(Path.Combine(installDirectory, fileName));
            }

            foreach (var directoryName in InstalledDirectories)
            {
                TryDeleteDirectory(Path.Combine(installDirectory, directoryName), true);
            }

            RemoveEmptyDirectories(installDirectory);
        }

        private static void RemoveEmptyDirectories(string installDirectory)
        {
            if (!Directory.Exists(installDirectory))
            {
                return;
            }

            foreach (var directory in Directory.GetDirectories(installDirectory, "*", SearchOption.AllDirectories)
                         .OrderByDescending(path => path.Length))
            {
                TryDeleteDirectory(directory, false);
            }

            TryDeleteDirectory(installDirectory, false);
        }

        private static void RemoveShortcuts()
        {
            var commonPrograms = Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms);
            var commonDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
            var productProgramsDirectory = Path.Combine(commonPrograms, ProductName);

            TryDeleteFile(Path.Combine(productProgramsDirectory, ProductName + ".lnk"));
            TryDeleteFile(Path.Combine(productProgramsDirectory, "卸载 " + ProductName + ".lnk"));
            TryDeleteDirectory(productProgramsDirectory, false);
            TryDeleteFile(Path.Combine(commonDesktop, ProductName + ".lnk"));
        }

        private static void RemoveUninstallRegistry()
        {
            DeleteUninstallKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            DeleteUninstallKey(RegistryHive.CurrentUser, RegistryView.Registry32);

            if (Environment.Is64BitOperatingSystem)
            {
                DeleteUninstallKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                DeleteUninstallKey(RegistryHive.CurrentUser, RegistryView.Registry64);
            }
        }

        private static void DeleteUninstallKey(RegistryHive hive, RegistryView view)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (var uninstallKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", true))
                {
                    uninstallKey?.DeleteSubKeyTree(ProductKey, false);
                }
            }
            catch
            {
                // Best effort: an absent key or denied alternate registry view should not block file cleanup.
            }
        }

        private static void ScheduleSelfRemoval(string installDirectory, bool purgeData)
        {
            var selfPath = Path.Combine(installDirectory, UninstallerExeName);
            if (!File.Exists(selfPath))
            {
                return;
            }

            var command = purgeData
                ? $"ping 127.0.0.1 -n 3 > nul & rmdir /s /q \"{installDirectory}\""
                : $"ping 127.0.0.1 -n 3 > nul & del /f /q \"{selfPath}\" & rmdir /q \"{installDirectory}\"";

            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c " + command,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false
            });
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
                // Leave locked user files in place.
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
                // Non-empty or locked directories are intentionally preserved.
            }
        }

        private static string EnsureTrailingSeparator(string path)
        {
            if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        private sealed class Options
        {
            public bool Silent { get; private set; }
            public bool FromUpgrade { get; private set; }
            public bool PurgeData { get; private set; }
            public string InstallDirectory { get; private set; }

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
                    else if (EqualsArg(arg, "/from-upgrade") || EqualsArg(arg, "--from-upgrade"))
                    {
                        options.FromUpgrade = true;
                        options.Silent = true;
                    }
                    else if (EqualsArg(arg, "/purge-data") || EqualsArg(arg, "--purge-data"))
                    {
                        options.PurgeData = true;
                    }
                    else if (StartsWithArg(arg, "/install-dir:") || StartsWithArg(arg, "--install-dir="))
                    {
                        options.InstallDirectory = arg.Substring(arg.IndexOfAny(new[] { ':', '=' }) + 1).Trim('"');
                    }
                    else if ((EqualsArg(arg, "/install-dir") || EqualsArg(arg, "--install-dir")) && i + 1 < args.Length)
                    {
                        options.InstallDirectory = args[++i].Trim('"');
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
