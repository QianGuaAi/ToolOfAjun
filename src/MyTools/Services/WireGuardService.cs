using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public class WireGuardStatus
    {
        public bool IsConnected { get; set; }
        public string InterfaceName { get; set; }
        public string IpAddress { get; set; }
        public string ErrorMessage { get; set; }
    }

    public static class WireGuardService
    {
        private static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string WgExePath = GetWireGuardPath();
        private static readonly string ConfigDir = Path.Combine(AppDir, "Configs");

        public static bool IsExeAvailable => WgExePath != "wireguard.exe";

        public static string GetSavedConfig(string interfaceName)
        {
            try
            {
                string configPath = Path.Combine(ConfigDir, $"{interfaceName}.conf");
                return File.Exists(configPath) ? File.ReadAllText(configPath) : null;
            }
            catch { return null; }
        }

        public static List<WireGuardTunnelInfo> GetSavedTunnels()
        {
            try
            {
                if (!Directory.Exists(ConfigDir))
                {
                    return new List<WireGuardTunnelInfo>();
                }

                return Directory.GetFiles(ConfigDir, "*.conf")
                    .Select(path => new FileInfo(path))
                    .OrderByDescending(info => info.LastWriteTime)
                    .Select(info => new WireGuardTunnelInfo
                    {
                        InterfaceName = Path.GetFileNameWithoutExtension(info.Name),
                        FilePath = info.FullName,
                        LastModified = info.LastWriteTime,
                        SizeBytes = info.Length
                    })
                    .ToList();
            }
            catch
            {
                return new List<WireGuardTunnelInfo>();
            }
        }

        public static void RenameSavedTunnel(string interfaceName, string newInterfaceName)
        {
            var sourceName = NormalizeInterfaceName(interfaceName);
            var targetName = NormalizeInterfaceName(newInterfaceName);
            if (string.IsNullOrWhiteSpace(sourceName))
            {
                throw new InvalidOperationException("原隧道名称不能为空。");
            }

            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new InvalidOperationException("新隧道名称不能为空。");
            }

            if (string.Equals(sourceName, targetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var sourcePath = Path.Combine(ConfigDir, sourceName + ".conf");
            var targetPath = Path.Combine(ConfigDir, targetName + ".conf");
            if (!File.Exists(sourcePath))
            {
                throw new FileNotFoundException("原隧道配置不存在。", sourcePath);
            }

            if (File.Exists(targetPath))
            {
                throw new IOException("已存在同名隧道配置。");
            }

            File.Move(sourcePath, targetPath);
        }

        public static void DeleteSavedTunnel(string interfaceName)
        {
            var name = NormalizeInterfaceName(interfaceName);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException("隧道名称不能为空。");
            }

            var path = Path.Combine(ConfigDir, name + ".conf");
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private static string NormalizeInterfaceName(string interfaceName)
        {
            var value = (interfaceName ?? string.Empty).Trim();
            if (value.EndsWith(".conf", StringComparison.OrdinalIgnoreCase))
            {
                value = Path.GetFileNameWithoutExtension(value);
            }

            foreach (var ch in Path.GetInvalidFileNameChars())
            {
                value = value.Replace(ch, '_');
            }

            return value;
        }

        private static string GetWireGuardPath()
        {
            // 1. Check local app bin directory (Integrated mode)
            string localPath = Path.Combine(AppDir, "NativeBinaries", "wireguard.exe");
            if (File.Exists(localPath)) return localPath;

            localPath = Path.Combine(AppDir, "bin", "wireguard.exe");
            if (File.Exists(localPath)) return localPath;

            // 2. Check app root
            localPath = Path.Combine(AppDir, "wireguard.exe");
            if (File.Exists(localPath)) return localPath;

            // 3. Check common installation paths (Fallback)
            string[] paths = {
                @"C:\Program Files\WireGuard\wireguard.exe",
                @"C:\Program Files (x86)\WireGuard\wireguard.exe"
            };

            foreach (var path in paths)
            {
                if (File.Exists(path)) return path;
            }
            return "wireguard.exe"; // Fallback to PATH
        }

        public static async Task<WireGuardStatus> ConnectAsync(string interfaceName, string configContent)
        {
            try
            {
                if (!Directory.Exists(ConfigDir)) Directory.CreateDirectory(ConfigDir);
                string configPath = Path.Combine(ConfigDir, $"{interfaceName}.conf");
                using (var stream = new FileStream(configPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    await writer.WriteAsync(configContent ?? string.Empty).ConfigureAwait(false);
                }

                // Command: wireguard.exe /installtunnelservice config_path
                var startInfo = new ProcessStartInfo
                {
                    FileName = WgExePath,
                    Arguments = $"/installtunnelservice \"{configPath}\"",
                    UseShellExecute = true,
                    Verb = "runas", // Ensure admin rights
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = Process.Start(startInfo))
                {
                    await Task.Run(() => process?.WaitForExit());
                }

                // Wait a bit for the interface to come up
                await Task.Delay(2000);
                return GetCurrentStatus(interfaceName);
            }
            catch (Exception ex)
            {
                return new WireGuardStatus { IsConnected = false, ErrorMessage = ex.Message };
            }
        }

        public static async Task<bool> DisconnectAsync(string interfaceName)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = WgExePath,
                    Arguments = $"/uninstalltunnelservice \"{interfaceName}\"",
                    UseShellExecute = true,
                    Verb = "runas",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = Process.Start(startInfo))
                {
                    await Task.Run(() => process?.WaitForExit());
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static WireGuardStatus GetCurrentStatus(string interfaceName)
        {
            var ni = NetworkInterface.GetAllNetworkInterfaces()
                .FirstOrDefault(i => i.Name.Equals(interfaceName, StringComparison.OrdinalIgnoreCase));

            if (ni != null && ni.OperationalStatus == OperationalStatus.Up)
            {
                var props = ni.GetIPProperties();
                var ip = props.UnicastAddresses
                    .FirstOrDefault(a => a.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)?
                    .Address.ToString();

                return new WireGuardStatus
                {
                    IsConnected = true,
                    InterfaceName = ni.Name,
                    IpAddress = ip
                };
            }

            return new WireGuardStatus { IsConnected = false };
        }
    }

    public sealed class WireGuardTunnelInfo
    {
        public string InterfaceName { get; set; }
        public string FilePath { get; set; }
        public DateTime LastModified { get; set; }
        public long SizeBytes { get; set; }
    }
}
