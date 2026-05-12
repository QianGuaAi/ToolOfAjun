using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using LibreHardwareMonitor.Hardware;

namespace MyTools.Services
{
    public class SensorReading
    {
        public string HardwareName { get; set; }
        public string HardwareKind { get; set; }
        public string SensorName { get; set; }
        public string SensorKind { get; set; }
        public string Value { get; set; }
        public string Unit { get; set; }
    }

    /// <summary>
    /// 包装 LibreHardwareMonitorLib，按需启动并轮询采集 CPU / GPU / 主板 温度、风扇、电压、负载等。
    /// </summary>
    public sealed class HardwareSensorService : IDisposable
    {
        private Computer _computer;
        private bool _disposed;
        private bool _firstRead = true;

        public bool IsAvailable => _computer != null;
        public string LastError { get; private set; }

        public static bool IsRunningAsAdmin
        {
            get
            {
                try
                {
                    using (var identity = WindowsIdentity.GetCurrent())
                    {
                        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
                    }
                }
                catch { return false; }
            }
        }

        public bool TryStart()
        {
            try
            {
                _computer = new Computer
                {
                    IsCpuEnabled = true,
                    IsGpuEnabled = true,
                    IsMotherboardEnabled = true,
                    IsControllerEnabled = true,
                    IsMemoryEnabled = true,
                    IsStorageEnabled = true
                };
                _computer.Open();
                _firstRead = true;
                LastError = null;
                return true;
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                AppLogService.Warning("HardwareSensorService start failed: {Msg}", ex.Message);
                _computer = null;
                return false;
            }
        }

        public List<SensorReading> ReadAll()
        {
            var list = new List<SensorReading>();
            if (_computer == null) return list;

            // LibreHardwareMonitor: first Update() may return null sensor values;
            // running Update() twice on the first pass primes the readings.
            int passes = _firstRead ? 2 : 1;
            for (int pass = 0; pass < passes; pass++)
            {
                foreach (var hw in _computer.Hardware)
                {
                    try { hw.Update(); } catch { }
                    foreach (var sub in hw.SubHardware)
                    {
                        try { sub.Update(); } catch { }
                    }
                }
            }
            _firstRead = false;

            foreach (var hw in _computer.Hardware)
            {
                foreach (var sub in hw.SubHardware)
                {
                    foreach (var s in sub.Sensors)
                    {
                        if (!s.Value.HasValue) continue;
                        list.Add(BuildReading(hw, s));
                    }
                }
                foreach (var s in hw.Sensors)
                {
                    if (!s.Value.HasValue) continue;
                    list.Add(BuildReading(hw, s));
                }
            }
            return list;
        }

        private static SensorReading BuildReading(IHardware hw, ISensor s)
        {
            string unit;
            string formatted;
            switch (s.SensorType)
            {
                case SensorType.Temperature:
                    unit = "°C";
                    formatted = s.Value.Value.ToString("0.#");
                    break;
                case SensorType.Fan:
                    unit = "RPM";
                    formatted = s.Value.Value.ToString("0");
                    break;
                case SensorType.Voltage:
                    unit = "V";
                    formatted = s.Value.Value.ToString("0.000");
                    break;
                case SensorType.Clock:
                    unit = "MHz";
                    formatted = s.Value.Value.ToString("0");
                    break;
                case SensorType.Load:
                    unit = "%";
                    formatted = s.Value.Value.ToString("0.#");
                    break;
                case SensorType.Power:
                    unit = "W";
                    formatted = s.Value.Value.ToString("0.#");
                    break;
                case SensorType.Data:
                    unit = "GB";
                    formatted = s.Value.Value.ToString("0.##");
                    break;
                case SensorType.Throughput:
                    unit = "B/s";
                    formatted = s.Value.Value.ToString("0");
                    break;
                default:
                    unit = "";
                    formatted = s.Value.Value.ToString("0.##");
                    break;
            }

            return new SensorReading
            {
                HardwareName = hw.Name,
                HardwareKind = hw.HardwareType.ToString(),
                SensorName = s.Name,
                SensorKind = s.SensorType.ToString(),
                Value = formatted,
                Unit = unit
            };
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { _computer?.Close(); }
            catch (Exception ex) { AppLogService.Warning("HardwareSensorService dispose failed: {Msg}", ex.Message); }
            _computer = null;
        }
    }
}
