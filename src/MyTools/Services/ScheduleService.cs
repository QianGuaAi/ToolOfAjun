using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MyTools.Services
{
    /// <summary>
    /// 排班版本本地 JSON 持久化。结构：
    /// %LOCALAPPDATA%\MyTools\Schedules\{YYYY-MM}\{versionName}.json
    /// </summary>
    public static class ScheduleService
    {
        public static readonly string Root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MyTools", "Schedules");

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore
        };

        private static string MonthDir(int year, int month)
        {
            return Path.Combine(Root, $"{year}-{month:00}");
        }

        private static string SafeFileName(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "v1";
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new System.Text.StringBuilder(raw.Length);
            foreach (var c in raw) sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            return sb.ToString();
        }

        public static List<ScheduleVersionInfo> ListVersions()
        {
            var list = new List<ScheduleVersionInfo>();
            if (!Directory.Exists(Root)) return list;

            foreach (var monthDir in Directory.EnumerateDirectories(Root))
            {
                var name = Path.GetFileName(monthDir);
                if (string.IsNullOrEmpty(name)) continue;
                var parts = name.Split('-');
                if (parts.Length != 2) continue;
                if (!int.TryParse(parts[0], out var year) || !int.TryParse(parts[1], out var month)) continue;

                foreach (var f in Directory.EnumerateFiles(monthDir, "*.json"))
                {
                    list.Add(new ScheduleVersionInfo
                    {
                        Year = year,
                        Month = month,
                        VersionName = Path.GetFileNameWithoutExtension(f),
                        UpdatedAt = File.GetLastWriteTime(f),
                        FilePath = f
                    });
                }
            }

            return list.OrderByDescending(v => v.Year)
                       .ThenByDescending(v => v.Month)
                       .ThenByDescending(v => v.UpdatedAt)
                       .ToList();
        }

        public static async Task<ScheduleVersion> LoadAsync(string filePath)
        {
            if (!File.Exists(filePath)) return null;
            using (var r = new StreamReader(filePath, System.Text.Encoding.UTF8))
            {
                var json = await r.ReadToEndAsync().ConfigureAwait(false);
                return JsonConvert.DeserializeObject<ScheduleVersion>(json, JsonSettings);
            }
        }

        public static async Task<string> SaveAsync(ScheduleVersion sched)
        {
            if (sched == null) throw new ArgumentNullException(nameof(sched));
            var dir = MonthDir(sched.Year, sched.Month);
            Directory.CreateDirectory(dir);
            var fileName = SafeFileName(sched.VersionName) + ".json";
            var filePath = Path.Combine(dir, fileName);
            sched.UpdatedAt = DateTime.Now;

            var json = JsonConvert.SerializeObject(sched, JsonSettings);
            using (var w = new StreamWriter(filePath, false, new System.Text.UTF8Encoding(false)))
            {
                await w.WriteAsync(json).ConfigureAwait(false);
            }
            return filePath;
        }

        public static void Delete(string filePath)
        {
            try { if (File.Exists(filePath)) File.Delete(filePath); } catch { }
        }

        /// <summary>取最新版本（用于"新建"时复制人员名单）。</summary>
        public static async Task<ScheduleVersion> GetLatestAsync()
        {
            var list = ListVersions();
            if (list.Count == 0) return null;
            return await LoadAsync(list[0].FilePath).ConfigureAwait(false);
        }

        /// <summary>
        /// 创建空白排班表：默认每日休息人数 周一-周五=6，周六=8，周日=9；并复制上一版本人员名单。
        /// </summary>
        public static async Task<ScheduleVersion> CreateNewAsync(int year, int month, string versionName)
        {
            var days = DateTime.DaysInMonth(year, month);
            var sched = new ScheduleVersion
            {
                Year = year,
                Month = month,
                VersionName = string.IsNullOrWhiteSpace(versionName) ? "v1" : versionName,
                DailyRestQuotas = new List<double>(days)
            };

            for (int d = 1; d <= days; d++)
            {
                var dow = new DateTime(year, month, d).DayOfWeek;
                double q = 6;
                if (dow == DayOfWeek.Saturday) q = 8;
                else if (dow == DayOfWeek.Sunday) q = 9;
                sched.DailyRestQuotas.Add(q);
            }

            // 复制最近一个版本的人员名单
            var latest = await GetLatestAsync().ConfigureAwait(false);
            if (latest != null)
            {
                foreach (var emp in latest.Employees)
                {
                    var row = new EmployeeRow { Name = emp.Name };
                    for (int i = 0; i < days; i++) row.Cells.Add(new ShiftCell());
                    sched.Employees.Add(row);
                }
            }

            return sched;
        }

        /// <summary>判断同名版本是否已存在。</summary>
        public static bool VersionExists(int year, int month, string versionName)
        {
            var dir = MonthDir(year, month);
            if (!Directory.Exists(dir)) return false;
            var f = Path.Combine(dir, SafeFileName(versionName) + ".json");
            return File.Exists(f);
        }
    }
}
