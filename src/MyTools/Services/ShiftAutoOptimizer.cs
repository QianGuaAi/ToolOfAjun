using System;
using System.Collections.Generic;
using System.Linq;

namespace MyTools.Services
{
    /// <summary>
    /// 自动设置休息日的启发式优化器：
    /// - 仅写入 白 / 休 两种值，不会自动写入特殊班次。
    /// - 完全尊重 IsManual=true 的单元格。
    /// - 硬约束：每人连续上班不得超过 5 天。
    /// - 软目标（按权重）：周末配对休息、节假日休息、节假日休 ≥4 天、总休 ≥8 天，并尽量均衡。
    /// </summary>
    public static class ShiftAutoOptimizer
    {
        public class OptimizeResult
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public int FilledCells { get; set; }
            public List<string> Warnings { get; } = new List<string>();
        }

        public static OptimizeResult Optimize(ScheduleVersion sched)
        {
            var result = new OptimizeResult();
            if (sched == null || sched.Employees.Count == 0)
            {
                result.Success = false;
                result.Message = "无人员可排。";
                return result;
            }

            int days = sched.DayCount;
            int empCount = sched.Employees.Count;

            // 确保每人 cells 长度对齐
            foreach (var e in sched.Employees)
            {
                while (e.Cells.Count < days) e.Cells.Add(new ShiftCell());
                if (e.Cells.Count > days) e.Cells.RemoveRange(days, e.Cells.Count - days);
            }

            // 预计算每天是否节假日
            var isHoliday = new bool[days];
            for (int d = 0; d < days; d++) isHoliday[d] = HolidayService.IsHoliday(sched.DateOf(d));

            int filled = 0;

            // ============== 主循环：逐天分配 ==============
            for (int d = 0; d < days; d++)
            {
                int quota = (d < sched.DailyRestQuotas.Count) ? sched.DailyRestQuotas[d] : 0;

                // 当天已有的手动休息计数
                int manualRest = 0;
                int manualWork = 0;
                foreach (var e in sched.Employees)
                {
                    var c = e.Cells[d];
                    if (c.IsManual && !string.IsNullOrEmpty(c.Code))
                    {
                        if (ShiftCodes.RestDays(c.Code) >= 1.0) manualRest++;
                        else if (ShiftCodes.IsWork(c.Code)) manualWork++;
                    }
                }

                // 候选：未手动指定的员工
                var candidates = new List<int>();
                for (int i = 0; i < empCount; i++)
                {
                    if (!sched.Employees[i].Cells[d].IsManual) candidates.Add(i);
                }

                int needRest = quota - manualRest;
                if (needRest < 0) needRest = 0; // 手动已超额，不再加

                // ---- Step 1: 强制休息（连续上班即将达 6 天）----
                var mustRest = new HashSet<int>();
                foreach (var i in candidates)
                {
                    if (CountConsecutiveWorkBefore(sched, i, d) >= 5)
                    {
                        mustRest.Add(i);
                    }
                }

                // ---- Step 2: 评分剩余候选 ----
                var remaining = candidates.Where(i => !mustRest.Contains(i)).ToList();
                int restToPick = Math.Max(0, needRest - mustRest.Count);

                // 总休天数（含手动 + 已分配的自动）
                var totalRest = new double[empCount];
                var totalHolidayRest = new double[empCount];
                for (int i = 0; i < empCount; i++)
                {
                    for (int k = 0; k < d; k++)
                    {
                        var c = sched.Employees[i].Cells[k];
                        var rd = ShiftCodes.RestDays(c.Code);
                        totalRest[i] += rd;
                        if (rd > 0 && isHoliday[k]) totalHolidayRest[i] += rd;
                    }
                }

                var dow = sched.DateOf(d).DayOfWeek;

                // 评分越高 → 越优先安排休息
                double Score(int empIdx)
                {
                    double s = 0;

                    // 总休越少越优先
                    s += Math.Max(0, 8.0 - totalRest[empIdx]) * 4.0;

                    // 节假日总休 < 4 时优先安排休（如果今天是节假日）
                    if (isHoliday[d])
                    {
                        s += Math.Max(0, 4.0 - totalHolidayRest[empIdx]) * 6.0;
                        s += 5.0; // 整体偏向把休排在节假日
                    }

                    // 周末配对：周日 & 周六已休 → 周日也休
                    if (dow == DayOfWeek.Sunday && d > 0)
                    {
                        var prev = sched.Employees[empIdx].Cells[d - 1];
                        if (ShiftCodes.RestDays(prev.Code) >= 1.0) s += 12.0;
                    }
                    // 周六 & 第二天周日 quota 仍有空间 → 鼓励周六休（启发：若该员工至今周末休得少）
                    if (dow == DayOfWeek.Saturday)
                    {
                        int weekendsRested = CountWeekendRestSoFar(sched, empIdx, d);
                        s += Math.Max(0, 2 - weekendsRested) * 3.0;
                    }

                    // 避免连续休 3 天以上：若前两天都休 → 降分
                    int prevRest = CountConsecutiveRestBefore(sched, empIdx, d);
                    if (prevRest >= 2) s -= 8.0;

                    return s;
                }

                var pickList = new HashSet<int>(
                    remaining
                        .Select(i => new { i, score = Score(i) })
                        .OrderByDescending(x => x.score)
                        .ThenBy(x => totalRest[x.i])
                        .Take(restToPick)
                        .Select(x => x.i));

                // ---- Step 3: 写入 ----
                foreach (var i in candidates)
                {
                    var cell = sched.Employees[i].Cells[d];
                    if (mustRest.Contains(i) || pickList.Contains(i))
                    {
                        cell.Code = ShiftCodes.Rest;
                        // 不标 Manual——保持自动可被再次优化
                        filled++;
                    }
                    else
                    {
                        cell.Code = ShiftCodes.Day;
                        filled++;
                    }
                }

                // 配额未达成警告
                int finalRest = manualRest + mustRest.Count + pickList.Count;
                if (finalRest < quota)
                {
                    result.Warnings.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 休息人数不足：要求 {quota}，仅排到 {finalRest}");
                }
            }

            // ============== 最终校验 ==============
            for (int i = 0; i < empCount; i++)
            {
                int maxRun = 0, run = 0;
                for (int d = 0; d < days; d++)
                {
                    if (ShiftCodes.IsWork(sched.Employees[i].Cells[d].Code)) { run++; if (run > maxRun) maxRun = run; }
                    else run = 0;
                }
                if (maxRun > 5)
                {
                    result.Warnings.Add($"{sched.Employees[i].Name} 连续上班 {maxRun} 天（应≤5）。");
                }
            }

            result.Success = true;
            result.FilledCells = filled;
            result.Message = $"已优化 {filled} 个单元格"
                + (result.Warnings.Count == 0 ? "。" : $"，{result.Warnings.Count} 项告警。");
            return result;
        }

        private static int CountConsecutiveWorkBefore(ScheduleVersion sched, int empIdx, int dayIdx)
        {
            int run = 0;
            for (int d = dayIdx - 1; d >= 0; d--)
            {
                if (ShiftCodes.IsWork(sched.Employees[empIdx].Cells[d].Code)) run++;
                else break;
            }
            return run;
        }

        private static int CountConsecutiveRestBefore(ScheduleVersion sched, int empIdx, int dayIdx)
        {
            int run = 0;
            for (int d = dayIdx - 1; d >= 0; d--)
            {
                if (ShiftCodes.RestDays(sched.Employees[empIdx].Cells[d].Code) >= 1.0) run++;
                else break;
            }
            return run;
        }

        private static int CountWeekendRestSoFar(ScheduleVersion sched, int empIdx, int dayIdx)
        {
            int count = 0;
            for (int d = 0; d < dayIdx; d++)
            {
                var dt = sched.DateOf(d);
                if (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday)
                {
                    if (ShiftCodes.RestDays(sched.Employees[empIdx].Cells[d].Code) >= 1.0) count++;
                }
            }
            return count;
        }
    }
}
