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
        public string SystemDiskName { get; set; }
        public string Uptime { get; set; }
        public string DotNetVersion { get; set; }
        public bool Is64BitOs { get; set; }
        public double CpuUsagePercent { get; set; }
        public string CpuUsageText { get; set; }
        public string CpuSummary { get; set; }
        public double MemoryUsagePercent { get; set; }
        public string MemoryUsageText { get; set; }
        public double DiskUsagePercent { get; set; }
        public string DiskUsageText { get; set; }
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
                snapshot.CpuUsagePercent = GetCpuLoadPercentage();
                snapshot.CpuSummary = $"{Environment.ProcessorCount} 逻辑处理器";
                snapshot.CpuUsageText = $"{snapshot.CpuUsagePercent:0}% 当前负载 · {snapshot.CpuSummary}";
            }
            catch
            {
                snapshot.CpuName = "未知";
                snapshot.CpuSummary = $"{Environment.ProcessorCount} 逻辑处理器";
                snapshot.CpuUsageText = snapshot.CpuSummary;
            }

            try
            {
                GetMemoryInfo(out var totalMb, out var availableMb, out var memoryLoad);
                snapshot.TotalRam = FormatMb(totalMb);
                snapshot.AvailableRam = FormatMb(availableMb);
                snapshot.MemoryUsagePercent = ClampPercent(memoryLoad);
                snapshot.MemoryUsageText = $"{snapshot.MemoryUsagePercent:0}% 已用 · {snapshot.AvailableRam} 可用";
            }
            catch
            {
                snapshot.TotalRam = "未知";
                snapshot.AvailableRam = "未知";
                snapshot.MemoryUsageText = "未知";
            }

            try
            {
                var disk = GetSystemDiskInfo();
                snapshot.SystemDisk = disk.Text;
                snapshot.SystemDiskName = disk.Name;
                snapshot.DiskUsagePercent = disk.UsedPercent;
                snapshot.DiskUsageText = disk.UsageText;
            }
            catch
            {
                snapshot.SystemDisk = "未知";
                snapshot.SystemDiskName = "系统盘";
                snapshot.DiskUsageText = "未知";
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

        private static double GetCpuLoadPercentage()
        {
            using (var searcher = new ManagementObjectSearcher("SELECT LoadPercentage FROM Win32_Processor"))
            using (var results = searcher.Get())
            {
                var values = results.Cast<ManagementObject>()
                    .Select(obj =>
                    {
                        try
                        {
                            var raw = obj["LoadPercentage"];
                            return raw == null ? (double?)null : Convert.ToDouble(raw);
                        }
                        finally { obj.Dispose(); }
                    })
                    .Where(value => value.HasValue)
                    .Select(value => value.Value)
                    .ToList();

                return values.Count == 0 ? 0 : ClampPercent(values.Average());
            }
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

        private static void GetMemoryInfo(out long totalMb, out long availableMb, out uint memoryLoad)
        {
            var mem = new MEMORYSTATUSEX { dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX)) };
            if (GlobalMemoryStatusEx(ref mem))
            {
                totalMb = (long)(mem.ullTotalPhys / (1024 * 1024));
                availableMb = (long)(mem.ullAvailPhys / (1024 * 1024));
                memoryLoad = mem.dwMemoryLoad;
            }
            else
            {
                totalMb = 0;
                availableMb = 0;
                memoryLoad = 0;
            }
        }

        private static DiskInfoSummary GetSystemDiskInfo()
        {
            var systemDrive = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
            if (string.IsNullOrWhiteSpace(systemDrive))
            {
                return DiskInfoSummary.Unknown();
            }

            var drive = DriveInfo.GetDrives().FirstOrDefault(d =>
                d.IsReady && d.RootDirectory.FullName.Equals(systemDrive, StringComparison.OrdinalIgnoreCase));

            if (drive == null)
            {
                return DiskInfoSummary.Unknown();
            }

            var totalGb = drive.TotalSize / (1024.0 * 1024 * 1024);
            var freeGb = drive.AvailableFreeSpace / (1024.0 * 1024 * 1024);
            var usedPercent = drive.TotalSize <= 0
                ? 0
                : ClampPercent((drive.TotalSize - drive.AvailableFreeSpace) * 100.0 / drive.TotalSize);
            var name = drive.Name.TrimEnd('\\');
            return new DiskInfoSummary
            {
                Name = name,
                Text = $"{name} {freeGb:0.#} GB 可用 / {totalGb:0.#} GB",
                UsedPercent = usedPercent,
                UsageText = $"{usedPercent:0}% 已用 · {freeGb:0.#} GB 可用"
            };
        }

        private static double ClampPercent(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return 0;
            }

            return Math.Max(0, Math.Min(100, value));
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

        private sealed class DiskInfoSummary
        {
            public string Name { get; set; }
            public string Text { get; set; }
            public double UsedPercent { get; set; }
            public string UsageText { get; set; }

            public static DiskInfoSummary Unknown()
            {
                return new DiskInfoSummary
                {
                    Name = "系统盘",
                    Text = "未知",
                    UsageText = "未知"
                };
            }
        }
    }
}
