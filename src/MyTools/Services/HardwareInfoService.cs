using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public class HardwareSummary
    {
        public string ComputerName { get; set; }
        public string OsName { get; set; }
        public string OsVersion { get; set; }
        public string OsArchitecture { get; set; }
        public string Manufacturer { get; set; }
        public string Model { get; set; }
        public string BiosVersion { get; set; }
        public string BiosDate { get; set; }
        public string MotherboardManufacturer { get; set; }
        public string MotherboardProduct { get; set; }
        public string SystemUpTime { get; set; }

        public List<CpuInfo> Cpus { get; set; } = new List<CpuInfo>();
        public List<GpuInfo> Gpus { get; set; } = new List<GpuInfo>();
        public List<MemoryModuleInfo> MemoryModules { get; set; } = new List<MemoryModuleInfo>();
        public string TotalMemoryGb { get; set; }
        public List<DiskInfo> Disks { get; set; } = new List<DiskInfo>();
        public List<NetworkAdapterInfo> NetworkAdapters { get; set; } = new List<NetworkAdapterInfo>();
        public List<MonitorInfo> Monitors { get; set; } = new List<MonitorInfo>();
    }

    public class CpuInfo
    {
        public string Name { get; set; }
        public string Manufacturer { get; set; }
        public int Cores { get; set; }
        public int LogicalProcessors { get; set; }
        public string MaxClockMhz { get; set; }
        public string CurrentClockMhz { get; set; }
        public string Socket { get; set; }
        public string L2CacheKb { get; set; }
        public string L3CacheKb { get; set; }
    }

    public class GpuInfo
    {
        public string Name { get; set; }
        public string Manufacturer { get; set; }
        public string DriverVersion { get; set; }
        public string DriverDate { get; set; }
        public string AdapterRamGb { get; set; }
        public string VideoMode { get; set; }
    }

    public class MemoryModuleInfo
    {
        public string BankLabel { get; set; }
        public string CapacityGb { get; set; }
        public string SpeedMhz { get; set; }
        public string Manufacturer { get; set; }
        public string PartNumber { get; set; }
        public string MemoryType { get; set; }
    }

    public class DiskInfo
    {
        public string Model { get; set; }
        public string InterfaceType { get; set; }
        public string SizeGb { get; set; }
        public string MediaType { get; set; }
        public string SerialNumber { get; set; }
    }

    public class NetworkAdapterInfo
    {
        public string Name { get; set; }
        public string Manufacturer { get; set; }
        public string MacAddress { get; set; }
        public string AdapterType { get; set; }
        public string LinkSpeed { get; set; }
    }

    public class MonitorInfo
    {
        public string Name { get; set; }
        public string Manufacturer { get; set; }
        public string Resolution { get; set; }
    }

    public static class HardwareInfoService
    {
        // Cached options to skip ACL / system properties for faster, lighter WMI queries.
        private static readonly EnumerationOptions FastEnumOptions = new EnumerationOptions
        {
            ReturnImmediately = true,
            Rewindable = false,
            UseAmendedQualifiers = false,
            DirectRead = true
        };

        public static Task<HardwareSummary> GetSummaryAsync()
        {
            return Task.Run(() => GetSummary());
        }

        // Helper: query WMI and dispose every ManagementObject + the collection + the searcher.
        private static void Query(string wql, Action<ManagementObject> handler)
        {
            using (var searcher = new ManagementObjectSearcher("root\\CIMV2", wql, FastEnumOptions))
            using (var results = searcher.Get())
            {
                foreach (ManagementObject mo in results)
                {
                    try { handler(mo); }
                    finally { mo.Dispose(); }
                }
            }
        }

        private static HardwareSummary GetSummary()
        {
            var summary = new HardwareSummary
            {
                ComputerName = Environment.MachineName
            };

            TryFill(() => Query("SELECT Caption, Version, OSArchitecture, LastBootUpTime FROM Win32_OperatingSystem", mo =>
            {
                summary.OsName = mo["Caption"]?.ToString();
                summary.OsVersion = mo["Version"]?.ToString();
                summary.OsArchitecture = mo["OSArchitecture"]?.ToString();
                var lastBoot = mo["LastBootUpTime"]?.ToString();
                if (!string.IsNullOrEmpty(lastBoot))
                {
                    try
                    {
                        var dt = ManagementDateTimeConverter.ToDateTime(lastBoot);
                        var up = DateTime.Now - dt;
                        summary.SystemUpTime = $"{(int)up.TotalDays} 天 {up.Hours} 小时 {up.Minutes} 分钟";
                    }
                    catch { }
                }
            }));

            TryFill(() => Query("SELECT Manufacturer, Model FROM Win32_ComputerSystem", mo =>
            {
                summary.Manufacturer = mo["Manufacturer"]?.ToString();
                summary.Model = mo["Model"]?.ToString();
            }));

            TryFill(() => Query("SELECT SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS", mo =>
            {
                summary.BiosVersion = mo["SMBIOSBIOSVersion"]?.ToString();
                var releaseDate = mo["ReleaseDate"]?.ToString();
                if (!string.IsNullOrEmpty(releaseDate))
                {
                    try { summary.BiosDate = ManagementDateTimeConverter.ToDateTime(releaseDate).ToString("yyyy-MM-dd"); } catch { }
                }
            }));

            TryFill(() => Query("SELECT Manufacturer, Product FROM Win32_BaseBoard", mo =>
            {
                summary.MotherboardManufacturer = mo["Manufacturer"]?.ToString();
                summary.MotherboardProduct = mo["Product"]?.ToString();
            }));

            TryFill(() => Query("SELECT Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, CurrentClockSpeed, SocketDesignation, L2CacheSize, L3CacheSize FROM Win32_Processor", mo =>
            {
                summary.Cpus.Add(new CpuInfo
                {
                    Name = mo["Name"]?.ToString()?.Trim(),
                    Manufacturer = mo["Manufacturer"]?.ToString(),
                    Cores = ToInt(mo["NumberOfCores"]),
                    LogicalProcessors = ToInt(mo["NumberOfLogicalProcessors"]),
                    MaxClockMhz = mo["MaxClockSpeed"]?.ToString(),
                    CurrentClockMhz = mo["CurrentClockSpeed"]?.ToString(),
                    Socket = mo["SocketDesignation"]?.ToString(),
                    L2CacheKb = mo["L2CacheSize"]?.ToString(),
                    L3CacheKb = mo["L3CacheSize"]?.ToString()
                });
            }));

            TryFill(() => Query("SELECT Name, AdapterCompatibility, DriverVersion, DriverDate, AdapterRAM, VideoModeDescription FROM Win32_VideoController", mo =>
            {
                    var ramRaw = mo["AdapterRAM"];
                    string ramGb = null;
                    if (ramRaw != null && ulong.TryParse(ramRaw.ToString(), out var bytes))
                    {
                        ramGb = (bytes / 1024.0 / 1024.0 / 1024.0).ToString("0.##") + " GB";
                    }
                    var driverDateRaw = mo["DriverDate"]?.ToString();
                    string driverDate = null;
                    if (!string.IsNullOrEmpty(driverDateRaw))
                    {
                        try { driverDate = ManagementDateTimeConverter.ToDateTime(driverDateRaw).ToString("yyyy-MM-dd"); } catch { }
                    }
                summary.Gpus.Add(new GpuInfo
                {
                    Name = mo["Name"]?.ToString(),
                    Manufacturer = mo["AdapterCompatibility"]?.ToString(),
                    DriverVersion = mo["DriverVersion"]?.ToString(),
                    DriverDate = driverDate,
                    AdapterRamGb = ramGb,
                    VideoMode = mo["VideoModeDescription"]?.ToString()
                });
            }));

            TryFill(() =>
            {
                ulong totalBytes = 0;
                Query("SELECT BankLabel, Capacity, Speed, Manufacturer, PartNumber, SMBIOSMemoryType, MemoryType FROM Win32_PhysicalMemory", mo =>
                {
                    var capRaw = mo["Capacity"];
                    string capGb = null;
                    if (capRaw != null && ulong.TryParse(capRaw.ToString(), out var bytes))
                    {
                        totalBytes += bytes;
                        capGb = (bytes / 1024.0 / 1024.0 / 1024.0).ToString("0.##") + " GB";
                    }
                    summary.MemoryModules.Add(new MemoryModuleInfo
                    {
                        BankLabel = mo["BankLabel"]?.ToString(),
                        CapacityGb = capGb,
                        SpeedMhz = mo["Speed"]?.ToString(),
                        Manufacturer = mo["Manufacturer"]?.ToString()?.Trim(),
                        PartNumber = mo["PartNumber"]?.ToString()?.Trim(),
                        MemoryType = TranslateMemoryType(mo["SMBIOSMemoryType"], mo["MemoryType"])
                    });
                });
                if (totalBytes > 0)
                {
                    summary.TotalMemoryGb = (totalBytes / 1024.0 / 1024.0 / 1024.0).ToString("0.##") + " GB";
                }
            });

            TryFill(() => Query("SELECT Model, InterfaceType, Size, MediaType, SerialNumber FROM Win32_DiskDrive", mo =>
            {
                var sizeRaw = mo["Size"];
                string sizeGb = null;
                if (sizeRaw != null && ulong.TryParse(sizeRaw.ToString(), out var bytes))
                {
                    sizeGb = (bytes / 1024.0 / 1024.0 / 1024.0).ToString("0.##") + " GB";
                }
                summary.Disks.Add(new DiskInfo
                {
                    Model = mo["Model"]?.ToString()?.Trim(),
                    InterfaceType = mo["InterfaceType"]?.ToString(),
                    SizeGb = sizeGb,
                    MediaType = mo["MediaType"]?.ToString(),
                    SerialNumber = mo["SerialNumber"]?.ToString()?.Trim()
                });
            }));

            TryFill(() => Query("SELECT Name, Manufacturer, MACAddress, AdapterType, Speed FROM Win32_NetworkAdapter WHERE PhysicalAdapter = TRUE AND MACAddress IS NOT NULL", mo =>
            {
                var speedRaw = mo["Speed"];
                string linkSpeed = null;
                if (speedRaw != null && ulong.TryParse(speedRaw.ToString(), out var bps) && bps > 0)
                {
                    if (bps >= 1_000_000_000) linkSpeed = (bps / 1_000_000_000.0).ToString("0.#") + " Gbps";
                    else if (bps >= 1_000_000) linkSpeed = (bps / 1_000_000.0).ToString("0.#") + " Mbps";
                    else linkSpeed = (bps / 1_000.0).ToString("0.#") + " Kbps";
                }
                summary.NetworkAdapters.Add(new NetworkAdapterInfo
                {
                    Name = mo["Name"]?.ToString(),
                    Manufacturer = mo["Manufacturer"]?.ToString(),
                    MacAddress = mo["MACAddress"]?.ToString(),
                    AdapterType = mo["AdapterType"]?.ToString(),
                    LinkSpeed = linkSpeed
                });
            }));

            TryFill(() => Query("SELECT Name, MonitorManufacturer, ScreenWidth, ScreenHeight FROM Win32_DesktopMonitor", mo =>
            {
                var w = mo["ScreenWidth"];
                var h = mo["ScreenHeight"];
                string res = null;
                if (w != null && h != null) res = $"{w}×{h}";
                summary.Monitors.Add(new MonitorInfo
                {
                    Name = mo["Name"]?.ToString(),
                    Manufacturer = mo["MonitorManufacturer"]?.ToString(),
                    Resolution = res
                });
            }));

            return summary;
        }

        private static int ToInt(object value)
        {
            if (value == null) return 0;
            return int.TryParse(value.ToString(), out var i) ? i : 0;
        }

        private static string TranslateMemoryType(object smbios, object legacy)
        {
            // SMBIOSMemoryType is more accurate on modern Windows
            int code = 0;
            if (smbios != null && int.TryParse(smbios.ToString(), out var s) && s > 0) code = s;
            else if (legacy != null && int.TryParse(legacy.ToString(), out var l)) code = l;

            switch (code)
            {
                case 20: return "DDR";
                case 21: return "DDR2";
                case 24: return "DDR3";
                case 26: return "DDR4";
                case 30: return "LPDDR4";
                case 34: return "DDR5";
                case 35: return "LPDDR5";
                default: return code > 0 ? $"类型 {code}" : null;
            }
        }

        private static void TryFill(Action action)
        {
            try { action(); }
            catch (Exception ex) { AppLogService.Warning("HardwareInfo query failed: {Msg}", ex.Message); }
        }
    }
}
