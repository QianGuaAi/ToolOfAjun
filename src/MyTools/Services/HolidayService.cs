using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace MyTools.Services
{
    /// <summary>
    /// 节假日：默认仅周末 (周六 / 周日)；用户可在 holidays.json 自定义法定节假日。
    /// 文件位置：%LOCALAPPDATA%\MyTools\holidays.json
    /// 格式：[ "2026-01-01", "2026-02-17", ... ]
    /// </summary>
    public static class HolidayService
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MyTools", "holidays.json");

        private static HashSet<DateTime> _custom;

        private static HashSet<DateTime> Load()
        {
            if (_custom != null) return _custom;
            _custom = new HashSet<DateTime>();
            try
            {
                if (File.Exists(FilePath))
                {
                    var text = File.ReadAllText(FilePath);
                    var arr = JsonConvert.DeserializeObject<List<string>>(text);
                    if (arr != null)
                    {
                        foreach (var s in arr)
                        {
                            if (DateTime.TryParse(s, out var d)) _custom.Add(d.Date);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Holiday list load failed: {Msg}", ex.Message);
            }
            return _custom;
        }

        public static bool IsHoliday(DateTime date)
        {
            if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday) return true;
            return Load().Contains(date.Date);
        }

        public static bool IsWeekend(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }

        public static void Reload() { _custom = null; }
    }
}
