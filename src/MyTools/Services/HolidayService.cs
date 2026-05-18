using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace MyTools.Services
{
    /// <summary>
    /// 节假日：默认包含周末 (周六 / 周日)；法定节假日来自用户配置 holidays.json，
    /// 若文件缺失则使用内置广西法定节假日 fallback（当前覆盖 2026 年）。
    /// 文件位置：%LOCALAPPDATA%\MyTools\holidays.json
    /// 格式：[ "2026-01-01", "2026-02-17", ... ]
    /// </summary>
    public static class HolidayService
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MyTools", "holidays.json");

        /// <summary>内置广西法定节假日 fallback（用户未提供 holidays.json 时使用）。</summary>
        private static readonly Dictionary<int, string[]> BuiltinHolidays = new Dictionary<int, string[]>
        {
            // 广西 2026 法定节假日（按国务院办公厅 2025-2026 节假日通知整理；调休工作日不视为节假日）
            [2026] = new[]
            {
                "2026-01-01",                                       // 元旦
                "2026-02-15", "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19", "2026-02-20", "2026-02-21", // 春节
                "2026-04-04", "2026-04-05", "2026-04-06",           // 清明
                "2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05", // 劳动节
                "2026-06-19",                                       // 端午
                "2026-09-25", "2026-09-26", "2026-09-27",           // 中秋
                "2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04", "2026-10-05", "2026-10-06", "2026-10-07", // 国庆
            }
        };

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

        /// <summary>当前年份是否已配置法定节假日。仅用户 holidays.json 与内置 fallback 命中其一即认为已配置。</summary>
        public static bool HasYearConfigured(int year)
        {
            // 用户配置中包含该年份任意一天？
            foreach (var d in Load())
            {
                if (d.Year == year) return true;
            }
            // fallback 命中？
            return BuiltinHolidays.ContainsKey(year);
        }

        public static bool IsHoliday(DateTime date)
        {
            if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday) return true;
            if (Load().Contains(date.Date)) return true;
            // fallback：用户未配置该年份，启用内置广西列表
            if (!UserHasYear(date.Year) && BuiltinHolidays.TryGetValue(date.Year, out var arr))
            {
                foreach (var s in arr)
                {
                    if (DateTime.TryParse(s, out var d) && d.Date == date.Date) return true;
                }
            }
            return false;
        }

        private static bool UserHasYear(int year)
        {
            foreach (var d in Load())
            {
                if (d.Year == year) return true;
            }
            return false;
        }

        public static bool IsWeekend(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }

        public static void Reload() { _custom = null; }
    }
}
