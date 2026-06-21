using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace MyTools.Services
{
    /// <summary>
    /// 单格数据：班次代码 + 是否手动指定（手动具有最高优先级，自动优化不可覆盖）。
    /// </summary>
    public class ShiftCell
    {
        /// <summary>白/卡/副/感/大/小/休/公/产/午；空字符串表示未指定。旧“夜”输入会归一化为“大”。</summary>
        public string Code { get; set; } = string.Empty;

        public bool IsManual { get; set; }
    }

    public class EmployeeRow
    {
        public string Name { get; set; } = string.Empty;
        public List<ShiftCell> Cells { get; set; } = new List<ShiftCell>();
    }

    /// <summary>
    /// 一个月份下的一个排班版本。
    /// </summary>
    public class ScheduleVersion
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string VersionName { get; set; } = "v1";

        /// <summary>第 3 行：每日总休息人数硬性要求。长度 = 当月天数，支持 0.5 天精度。</summary>
        public List<double> DailyRestQuotas { get; set; } = new List<double>();

        public List<EmployeeRow> Employees { get; set; } = new List<EmployeeRow>();

        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;

        /// <summary>
        /// 自动设置休息日或 Excel 导入成功后写入的时间戳，仅作完成时间记录；Excel 导出允许随时导出当前表格。
        /// </summary>
        public DateTime? GeneratedAt { get; set; }

        /// <summary>当月天数（28-31）。</summary>
        [JsonIgnore]
        public int DayCount => DateTime.DaysInMonth(Year, Month);

        [JsonIgnore]
        public bool HasGenerated => GeneratedAt.HasValue;

        public DateTime DateOf(int dayIndex0Based) => new DateTime(Year, Month, dayIndex0Based + 1);
    }

    /// <summary>
    /// 版本列表项（仅用于左侧列表展示，不含完整数据）。
    /// </summary>
    public class ScheduleVersionInfo
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string VersionName { get; set; } = string.Empty;
        public DateTime UpdatedAt { get; set; }
        public string FilePath { get; set; } = string.Empty;

        public string DisplayLabel => $"{Year}-{Month:00} · {VersionName}";
        public string DisplaySub => UpdatedAt.ToString("MM-dd HH:mm");
    }

    /// <summary>
    /// 班次代码常量与显示规则。
    /// </summary>
    public static class ShiftCodes
    {
        public const string Empty = "";
        public const string Day = "白";       // 白班
        public const string Card = "卡";      // 卡班
        public const string Deputy = "副";    // 副小
        public const string Infect = "感";    // 感染科
        public const string Big = "大";       // 夜班（旧数据仍用“大”存储）
        public const string Small = "小";     // 小夜
        public const string Rest = "休";      // 休息
        public const string Public = "公";    // 公休
        public const string Maternity = "产"; // 产假
        public const string Half = "午";      // 下午休 0.5

        public static readonly string[] All = { Day, Card, Deputy, Infect, Big, Small, Rest, Public, Maternity, Half };

        public static string Normalize(string code)
        {
            code = (code ?? string.Empty).Trim();
            return code == "夜" ? Big : code;
        }

        /// <summary>是否计入"上班"。</summary>
        public static bool IsWork(string code)
        {
            code = Normalize(code);
            return code == Day || code == Card || code == Deputy || code == Infect || code == Big || code == Small;
        }

        /// <summary>休息天数（0 / 0.5 / 1）。</summary>
        public static double RestDays(string code)
        {
            code = Normalize(code);
            if (code == Rest || code == Public || code == Maternity) return 1.0;
            if (code == Half) return 0.5;
            return 0.0;
        }

        public static double WorkDays(string code) => IsWork(code) ? 1.0 : 0.0;

        public static string Description(string code)
        {
            code = Normalize(code);
            switch (code)
            {
                case Day: return "白班";
                case Card: return "卡班";
                case Deputy: return "副小";
                case Infect: return "感染科";
                case Big: return "夜班";
                case Small: return "小夜";
                case Rest: return "休息";
                case Public: return "公休";
                case Maternity: return "产假";
                case Half: return "下午休 0.5";
                default: return "未指定";
            }
        }
    }
}
