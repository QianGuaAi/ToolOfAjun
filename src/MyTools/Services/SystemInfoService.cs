using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;

namespace MyTools.Services
{
    public class SystemInfoSnapshot
    {
        public string OsName { get; set; }
        public string CpuName { get; set; }
        public string TotalRam { get; set; }
        public string AvailableRam { get; set; }
        public string SystemDisk { get; set; }
        public string Uptime { get; set; }
        public string DotNetVersion { get; set; }
        public bool Is64BitOs { get; set; }
    }

    public static class SystemInfoService
    {
        [DllImport("kernel32.dll")]
        private static extern ulong GetTickCount64();

        public static System.Threading.Tasks.Task<SystemInfoSnapshot> GetSnapshotAsync()
        {
            return System.Threading.Tasks.Task.Run(() => GetSnapshot());
        }

        public static SystemInfoSnapshot GetSnapshot()
        {
            ulong tickMs;
            try { tickMs = GetTickCount64(); }
            catch { tickMs = (ulong)(Environment.TickCount & int.MaxValue); }

            var snapshot = new SystemInfoSnapshot
            {
                OsName = OsVersionService.DisplayName,
                Is64BitOs = Environment.Is64BitOperatingSystem,
                DotNetVersion = Environment.Version.ToString(),
                Uptime = FormatUptime(TimeSpan.FromMilliseconds(tickMs))
            };

            try
            {
                snapshot.CpuName = GetCpuName();
            }
            catch
            {
                snapshot.CpuName = "未知";
            }

            try
            {
                GetMemoryInfo(out var totalMb, out var availableMb);
                snapshot.TotalRam = FormatMb(totalMb);
                snapshot.AvailableRam = FormatMb(availableMb);
            }
            catch
            {
                snapshot.TotalRam = "未知";
                snapshot.AvailableRam = "未知";
            }

            try
            {
                snapshot.SystemDisk = GetSystemDiskInfo();
            }
            catch
            {
                snapshot.SystemDisk = "未知";
            }

            return snapshot;
        }

        private static string GetCpuName()
        {
            using (var searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor"))
            using (var results = searcher.Get())
            {
                foreach (ManagementObject obj in results)
                {
                    try
                    {
                        var name = obj["Name"]?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            return name;
                        }
                    }
                    finally { obj.Dispose(); }
                }
            }

            return "未知";
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx(ref MEMORYSTATUSEX lpBuffer);

        private static void GetMemoryInfo(out long totalMb, out long availableMb)
        {
            var mem = new MEMORYSTATUSEX { dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX)) };
            if (GlobalMemoryStatusEx(ref mem))
            {
                totalMb = (long)(mem.ullTotalPhys / (1024 * 1024));
                availableMb = (long)(mem.ullAvailPhys / (1024 * 1024));
            }
            else
            {
                totalMb = 0;
                availableMb = 0;
            }
        }

        private static string GetSystemDiskInfo()
        {
            var systemDrive = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
            if (string.IsNullOrWhiteSpace(systemDrive))
            {
                return "未知";
            }

            var drive = DriveInfo.GetDrives().FirstOrDefault(d =>
                d.IsReady && d.RootDirectory.FullName.Equals(systemDrive, StringComparison.OrdinalIgnoreCase));

            if (drive == null)
            {
                return "未知";
            }

            var totalGb = drive.TotalSize / (1024.0 * 1024 * 1024);
            var freeGb = drive.AvailableFreeSpace / (1024.0 * 1024 * 1024);
            return $"{drive.Name.TrimEnd('\\')} {freeGb:0.#} GB 可用 / {totalGb:0.#} GB";
        }

        private static string FormatMb(long mb)
        {
            if (mb >= 1024)
            {
                return $"{mb / 1024.0:0.#} GB";
            }

            return $"{mb} MB";
        }

        private static string FormatUptime(TimeSpan ts)
        {
            if (ts.TotalDays >= 1)
            {
                return $"{(int)ts.TotalDays} 天 {ts.Hours} 小时 {ts.Minutes} 分钟";
            }

            if (ts.TotalHours >= 1)
            {
                return $"{(int)ts.TotalHours} 小时 {ts.Minutes} 分钟";
            }

            return $"{(int)ts.TotalMinutes} 分钟";
        }
    }
}
