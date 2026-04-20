using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
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
                File.WriteAllText(configPath, configContent);

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
}
