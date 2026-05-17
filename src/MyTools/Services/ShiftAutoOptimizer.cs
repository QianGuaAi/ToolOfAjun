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
        private const double Epsilon = 0.001;

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
                double quota = (d < sched.DailyRestQuotas.Count) ? sched.DailyRestQuotas[d] : 0;

                // 当天已有的手动休息计数，精确到 0.5。
                double manualRest = 0;
                foreach (var e in sched.Employees)
                {
                    var c = e.Cells[d];
                    if (c.IsManual && !string.IsNullOrEmpty(c.Code))
                    {
                        manualRest += ShiftCodes.RestDays(c.Code);
                    }
                }

                // 候选：未手动指定的员工
                var candidates = new List<int>();
                for (int i = 0; i < empCount; i++)
                {
                    if (!sched.Employees[i].Cells[d].IsManual) candidates.Add(i);
                }

                double restNeed = quota - manualRest;
                if (restNeed < -Epsilon)
                {
                    result.Warnings.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 手动休息已超额：已休 {FormatNumber(manualRest)}，目标 {FormatNumber(quota)}。");
                }

                int restToPick;
                if (restNeed <= Epsilon)
                {
                    restToPick = 0;
                }
                else
                {
                    var rounded = Math.Round(restNeed);
                    if (Math.Abs(restNeed - rounded) > Epsilon)
                    {
                        result.Warnings.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 需要再安排 {FormatNumber(restNeed)} 天休息；自动优化只填整天“休”，已按不超额原则安排 {Math.Floor(restNeed)} 天。");
                        restToPick = (int)Math.Floor(restNeed);
                    }
                    else
                    {
                        restToPick = (int)rounded;
                    }
                }

                if (restToPick > candidates.Count)
                {
                    result.Warnings.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 可自动安排人员不足：需要 {restToPick} 人休息，仅有 {candidates.Count} 个可改格子。");
                    restToPick = candidates.Count;
                }

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
                if (mustRest.Count > restToPick)
                {
                    result.Warnings.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 为避免连续上班需要 {mustRest.Count} 人休息，但当日配额只允许 {restToPick} 人自动休息。");
                }

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

                    if (mustRest.Contains(empIdx)) s += 10000.0;

                    return s;
                }

                var pickList = new HashSet<int>(
                    candidates
                        .Select(i => new { i, score = Score(i) })
                        .OrderByDescending(x => x.score)
                        .ThenBy(x => totalRest[x.i])
                        .Take(restToPick)
                        .Select(x => x.i));

                // ---- Step 3: 写入 ----
                foreach (var i in candidates)
                {
                    var cell = sched.Employees[i].Cells[d];
                    var nextCode = pickList.Contains(i) ? ShiftCodes.Rest : ShiftCodes.Day;
                    if (cell.Code != nextCode)
                    {
                        filled++;
                    }
                    cell.Code = nextCode;
                    // 不标 Manual——保持自动可被再次优化
                }
            }

            filled += RepairConsecutiveRuns(sched, result.Warnings);

            // ============== 最终校验 ==============
            var hardIssues = CollectHardConstraintIssues(sched).ToList();
            foreach (var issue in hardIssues)
            {
                result.Warnings.Add(issue);
            }

            result.Success = hardIssues.Count == 0;
            result.FilledCells = filled;
            result.Message = result.Success
                ? $"已优化 {filled} 个单元格。"
                : $"已优化 {filled} 个单元格，但仍有 {hardIssues.Count} 项硬约束未满足。";
            return result;
        }

        private static int RepairConsecutiveRuns(ScheduleVersion sched, List<string> warnings)
        {
            int repairs = 0;
            int guard = Math.Max(1, sched.DayCount * Math.Max(1, sched.Employees.Count) * 4);
            while (guard-- > 0)
            {
                if (!TryFindOverlongRun(sched, out var empIdx, out var startDay, out var endDay, out var length))
                {
                    return repairs;
                }

                var repaired = false;
                foreach (var day in CandidateBreakDays(startDay, endDay))
                {
                    var cell = sched.Employees[empIdx].Cells[day];
                    if (cell.IsManual || !ShiftCodes.IsWork(cell.Code))
                    {
                        continue;
                    }

                    if (TryBreakRunOnDay(sched, empIdx, day))
                    {
                        repairs++;
                        repaired = true;
                        break;
                    }
                }

                if (!repaired)
                {
                    warnings.Add($"{sched.Employees[empIdx].Name} 连续上班 {length} 天，且连续区间内没有可自动调整的格子。");
                    return repairs;
                }
            }

            warnings.Add("连续上班修复达到迭代上限，请检查冲突侧栏。");
            return repairs;
        }

        private static IEnumerable<int> CandidateBreakDays(int startDay, int endDay)
        {
            var middle = (startDay + endDay) / 2;
            return Enumerable.Range(startDay, endDay - startDay + 1)
                .OrderBy(day => Math.Abs(day - middle));
        }

        private static bool TryBreakRunOnDay(ScheduleVersion sched, int empIdx, int day)
        {
            var quota = day < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[day] : 0;
            var actual = ComputeColumnRestCount(sched, day);
            var offender = sched.Employees[empIdx].Cells[day];

            if (actual + 1.0 <= quota + Epsilon)
            {
                offender.Code = ShiftCodes.Rest;
                return true;
            }

            if (Math.Abs(actual - quota) > Epsilon)
            {
                return false;
            }

            var donorIndexes = Enumerable.Range(0, sched.Employees.Count)
                .Where(i => i != empIdx)
                .Where(i =>
                {
                    var cell = sched.Employees[i].Cells[day];
                    return !cell.IsManual && cell.Code == ShiftCodes.Rest;
                })
                .OrderByDescending(i => ComputeTotalRest(sched.Employees[i]))
                .ToList();

            foreach (var donorIdx in donorIndexes)
            {
                var donor = sched.Employees[donorIdx].Cells[day];
                var oldOffender = offender.Code;
                var oldDonor = donor.Code;
                offender.Code = ShiftCodes.Rest;
                donor.Code = ShiftCodes.Day;

                if (ComputeMaxConsecutiveWork(sched.Employees[empIdx]) <= 5 &&
                    ComputeMaxConsecutiveWork(sched.Employees[donorIdx]) <= 5)
                {
                    return true;
                }

                offender.Code = oldOffender;
                donor.Code = oldDonor;
            }

            return false;
        }

        private static bool TryFindOverlongRun(ScheduleVersion sched, out int employeeIndex, out int startDay, out int endDay, out int length)
        {
            for (int i = 0; i < sched.Employees.Count; i++)
            {
                var run = 0;
                for (int d = 0; d < sched.DayCount; d++)
                {
                    if (ShiftCodes.IsWork(sched.Employees[i].Cells[d].Code))
                    {
                        run++;
                        if (run > 5)
                        {
                            employeeIndex = i;
                            startDay = d - run + 1;
                            endDay = d;
                            length = run;
                            return true;
                        }
                    }
                    else
                    {
                        run = 0;
                    }
                }
            }

            employeeIndex = -1;
            startDay = -1;
            endDay = -1;
            length = 0;
            return false;
        }

        private static IEnumerable<string> CollectHardConstraintIssues(ScheduleVersion sched)
        {
            for (int d = 0; d < sched.DayCount; d++)
            {
                var actual = ComputeColumnRestCount(sched, d);
                var quota = d < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[d] : 0;
                var delta = actual - quota;
                if (Math.Abs(delta) > Epsilon)
                {
                    yield return $"{sched.Year}-{sched.Month:00}-{d + 1:00} 实际休息 {FormatNumber(actual)} / 总休目标 {FormatNumber(quota)}，{(delta > 0 ? "多" : "差")} {FormatNumber(Math.Abs(delta))}。";
                }
            }

            for (int i = 0; i < sched.Employees.Count; i++)
            {
                var maxRun = ComputeMaxConsecutiveWork(sched.Employees[i]);
                if (maxRun > 5)
                {
                    yield return $"{sched.Employees[i].Name} 连续上班 {maxRun} 天（应≤5）。";
                }
            }
        }

        private static double ComputeColumnRestCount(ScheduleVersion sched, int dayIdx)
        {
            double sum = 0;
            foreach (var emp in sched.Employees)
            {
                if (emp.Cells != null && dayIdx < emp.Cells.Count)
                {
                    sum += ShiftCodes.RestDays(emp.Cells[dayIdx].Code);
                }
            }

            return sum;
        }

        private static double ComputeTotalRest(EmployeeRow employee)
        {
            return employee.Cells == null ? 0 : employee.Cells.Sum(cell => ShiftCodes.RestDays(cell.Code));
        }

        private static int ComputeMaxConsecutiveWork(EmployeeRow employee)
        {
            var maxRun = 0;
            var run = 0;
            foreach (var cell in employee.Cells)
            {
                if (ShiftCodes.IsWork(cell.Code))
                {
                    run++;
                    if (run > maxRun) maxRun = run;
                }
                else
                {
                    run = 0;
                }
            }

            return maxRun;
        }

        private static string FormatNumber(double value)
        {
            return Math.Abs(value % 1) < Epsilon ? ((int)Math.Round(value)).ToString() : value.ToString("0.#");
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
