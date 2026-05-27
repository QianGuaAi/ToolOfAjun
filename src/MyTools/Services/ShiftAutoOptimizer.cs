using System;
using System.Collections.Generic;
using System.Linq;

namespace MyTools.Services
{
    /// <summary>
    /// 自动设置休息日：严格按规范的 5 级优先级生成，仅写入 白 / 休 两种值。
    /// 优先级（高 → 低）：
    ///   1. 任何人连续上班天数 ≤ 5（硬约束）
    ///   2. 周末（周六+周日）连续休息 2 天，并尽量均衡
    ///   3. 休息日尽量安排在节假日（含周末+法定节假日），并尽量均衡
    ///   4. 节假日休 ≥ 4 天 / 人，并尽量均衡
    ///   5. 总休 ≥ 9 天 / 人，并尽量均衡
    /// 每日总休数按 DailyRestQuotas 的硬性配额执行；用户手动修改 (IsManual=true) 不会被改写。
    /// </summary>
    public static class ShiftAutoOptimizer
    {
        private const double Epsilon = 0.001;
        private const int MaxConsecutiveWork = 5;
        private const int TargetTotalRest = 9;
        private const int TargetHolidayRest = 4;

        public class OptimizeResult
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public int FilledCells { get; set; }
            public List<string> Warnings { get; } = new List<string>();
        }

        public static OptimizeResult Optimize(ScheduleVersion sched)
        {
            return Optimize(sched, null);
        }

        /// <summary>
        /// 自动设置休息日。<paramref name="preservedCells"/> 中显式标记的 (empIdx, dayIdx) 在 sanitize 阶段不会被清空，
        /// 与手工锁定一同作为"已分配休"基线参与 §4.2.3 的前置拒绝判定（规范 §4.2.1/§7）。
        /// </summary>
        public static OptimizeResult Optimize(ScheduleVersion sched, IReadOnlyCollection<(int emp, int day)> preservedCells)
        {
            var result = new OptimizeResult();
            if (sched == null || sched.Employees.Count == 0)
            {
                result.Success = false;
                result.Message = "无人员可排。";
                return result;
            }

            // 每次点击都用新的随机种子，让多次执行产出不同的合规方案
            var rnd = new Random(unchecked(Environment.TickCount * 397) ^ Guid.NewGuid().GetHashCode());

            int days = sched.DayCount;
            int empCount = sched.Employees.Count;

            // --------- 0a. 法定节假日配置存在校验（规范 §4.2.3 #1） ---------
            if (!HolidayService.HasYearConfigured(sched.Year))
            {
                result.Success = false;
                result.Message = $"未配置 {sched.Year} 年法定节假日，请先在 holidays.json 中补充该年配置。";
                return result;
            }

            // --------- 0b. Sanitize：补齐/截断列长度，清空"非手动且未保留"的旧自动结果 ---------
            foreach (var e in sched.Employees)
            {
                while (e.Cells.Count < days) e.Cells.Add(new ShiftCell());
                if (e.Cells.Count > days) e.Cells.RemoveRange(days, e.Cells.Count - days);
            }

            var activeIndexes = Enumerable.Range(0, empCount)
                .Where(i => !string.IsNullOrWhiteSpace(sched.Employees[i].Name))
                .ToList();
            if (activeIndexes.Count == 0)
            {
                result.Success = false;
                result.Message = "无有效人员可排。";
                return result;
            }

            foreach (var i in Enumerable.Range(0, empCount).Except(activeIndexes))
            {
                foreach (var cell in sched.Employees[i].Cells)
                {
                    cell.Code = string.Empty;
                    cell.IsManual = false;
                }
            }

            var preserved = preservedCells != null
                ? new HashSet<(int emp, int day)>(preservedCells)
                : new HashSet<(int emp, int day)>();

            // --------- 0c. §4.2.3 #2 + #3 前置拒绝 ---------
            //   #2 sanitize 后每日 (总休 - 已锁休) 必须保持 0.5 步进，允许自动生成后欠 0.5
            //   #3 H2 强制休 > 当日剩余休息名额：当日因 5 连班必须强制休的人数超过 quota 余量
            var rejections = new List<string>();
            for (int d = 0; d < days; d++)
            {
                double quota = d < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[d] : 0;
                double lockedRest = 0;
                foreach (var i in activeIndexes)
                {
                    var c = sched.Employees[i].Cells[d];
                    if (c.IsManual || preserved.Contains((i, d)))
                    {
                        lockedRest += ShiftCodes.RestDays(c.Code);
                    }
                }

                double remaining = quota - lockedRest;
                if (remaining < -Epsilon)
                {
                    rejections.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 锁定休息 {FormatNumber(lockedRest)} 已超出总休 {FormatNumber(quota)}。");
                    continue;
                }

                double frac = Math.Abs(remaining - Math.Round(remaining));
                if (frac > Epsilon && Math.Abs(frac - 0.5) > Epsilon)
                {
                    rejections.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 总休减锁定休后剩余 {FormatNumber(remaining)} 天，无法用整天补齐。");
                }

                // H2 强制休：当前考察 d 之前是否有 ≥5 连续上班的人（基于已锁定单元格），且 d 自身可被自动占用
                int forcedCount = 0;
                foreach (var i in activeIndexes)
                {
                    var c = sched.Employees[i].Cells[d];
                    if (c.IsManual || preserved.Contains((i, d))) continue; // d 已锁定（含休或上班），不需要算强制
                    int run = 0;
                    for (int p = d - 1; p >= 0; p--)
                    {
                        var pc = sched.Employees[i].Cells[p];
                        if (!ShiftCodes.IsWork(pc.Code)) break;
                        if (!(pc.IsManual || preserved.Contains((i, p)))) break; // 仅锁定上班才算"必然连续"
                        run++;
                        if (run >= MaxConsecutiveWork) break;
                    }
                    if (run >= MaxConsecutiveWork) forcedCount++;
                }
                var autoRestCapacity = Math.Max(0, (int)Math.Floor(remaining + Epsilon));
                if (forcedCount > autoRestCapacity)
                {
                    rejections.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 因 5 连班必须强制休的人数 {forcedCount} 超过当日可自动安排休息名额 {autoRestCapacity}。");
                }
            }

            if (rejections.Count > 0)
            {
                result.Success = false;
                result.Message = "排班无法生成（违反硬性条件）：";
                foreach (var r in rejections) result.Warnings.Add(r);
                return result;
            }

            int filled = 0;
            for (int d = 0; d < days; d++)
            {
                for (int i = 0; i < empCount; i++)
                {
                    var cell = sched.Employees[i].Cells[d];
                    if (!cell.IsManual && !preserved.Contains((i, d)) && !string.IsNullOrEmpty(cell.Code))
                    {
                        cell.Code = string.Empty;
                    }
                }
            }

            // --------- 1. 预计算节假日 / 周末 / 周末对 ---------
            var isHoliday = new bool[days];
            var isWeekend = new bool[days];
            for (int d = 0; d < days; d++)
            {
                var dt = sched.DateOf(d);
                isWeekend[d] = dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday;
                isHoliday[d] = HolidayService.IsHoliday(dt);
            }

            var weekendPairs = new List<(int sat, int sun)>();
            for (int d = 0; d < days - 1; d++)
            {
                if (sched.DateOf(d).DayOfWeek == DayOfWeek.Saturday
                    && sched.DateOf(d + 1).DayOfWeek == DayOfWeek.Sunday)
                {
                    weekendPairs.Add((d, d + 1));
                }
            }

            // 每日剩余可分配休息（quota - 当前已 rest，0.5 精度）
            double RemainingSlots(int d)
            {
                double cur = 0;
                foreach (var i in activeIndexes) cur += ShiftCodes.RestDays(sched.Employees[i].Cells[d].Code);
                double quota = d < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[d] : 0;
                return quota - cur;
            }

            int CountPairWeekendsRested(int empIdx)
            {
                int cnt = 0;
                foreach (var (s, u) in weekendPairs)
                {
                    if (ShiftCodes.RestDays(sched.Employees[empIdx].Cells[s].Code) >= 1.0
                        && ShiftCodes.RestDays(sched.Employees[empIdx].Cells[u].Code) >= 1.0)
                    {
                        cnt++;
                    }
                }
                return cnt;
            }

            int CountHolidayRested(int empIdx)
            {
                int cnt = 0;
                for (int d = 0; d < days; d++)
                {
                    if (isHoliday[d] && ShiftCodes.RestDays(sched.Employees[empIdx].Cells[d].Code) >= 1.0)
                        cnt++;
                }
                return cnt;
            }

            double TotalRest(int empIdx)
            {
                double r = 0;
                foreach (var c in sched.Employees[empIdx].Cells) r += ShiftCodes.RestDays(c.Code);
                return r;
            }

            // 试将 (empIdx, dayIdx) 设为休：要求非手动、当前未分配、不会让"前 5 天 + 当天"形成 6 连休
            //（连续休本身没硬约束，但避免 3 连休是软目标，这里仅检查是否会让"上班 5+休"无效）
            bool TryAssignRest(int empIdx, int dayIdx)
            {
                var cell = sched.Employees[empIdx].Cells[dayIdx];
                if (cell.IsManual || preserved.Contains((empIdx, dayIdx))) return false;
                if (!string.IsNullOrEmpty(cell.Code)) return false; // 已被其它阶段分配
                cell.Code = ShiftCodes.Rest;
                return true;
            }

            // ============== 阶段 A：周末配对（Sat+Sun 整对休） ==============
            // 优先安排给"已配对周末数最少 + 总休最少"的人。
            foreach (var (sat, sun) in weekendPairs)
            {
                double slots = Math.Min(RemainingSlots(sat), RemainingSlots(sun));
                if (slots < 1.0 - Epsilon) continue;

                var candidates = activeIndexes
                    .Where(i =>
                    {
                        var cs = sched.Employees[i].Cells[sat];
                        var cu = sched.Employees[i].Cells[sun];
                        return !cs.IsManual && !cu.IsManual
                            && !preserved.Contains((i, sat)) && !preserved.Contains((i, sun))
                            && string.IsNullOrEmpty(cs.Code) && string.IsNullOrEmpty(cu.Code);
                    })
                    .OrderBy(CountPairWeekendsRested)
                    .ThenBy(TotalRest)
                    .ThenBy(_ => rnd.Next())
                    .ToList();

                int take = (int)Math.Floor(slots + Epsilon);
                foreach (var i in candidates.Take(take))
                {
                    if (TryAssignRest(i, sat) && TryAssignRest(i, sun)) { /* assigned */ }
                }
            }

            // ============== 阶段 B：法定节假日（非周末）预分配 ==============
            // 把节假日休息分给"节假日休最少"的人，提升优先级 4 的达成度。
            for (int d = 0; d < days; d++)
            {
                if (!isHoliday[d] || isWeekend[d]) continue;
                double rem = RemainingSlots(d);
                if (rem < 1.0 - Epsilon) continue;

                var candidates = activeIndexes
                    .Where(i => !sched.Employees[i].Cells[d].IsManual
                                && !preserved.Contains((i, d))
                                && string.IsNullOrEmpty(sched.Employees[i].Cells[d].Code))
                    .OrderBy(CountHolidayRested)
                    .ThenBy(TotalRest)
                    .ThenBy(_ => rnd.Next())
                    .ToList();

                int take = (int)Math.Floor(rem + Epsilon);
                foreach (var i in candidates.Take(take)) TryAssignRest(i, d);
            }

            // ============== 阶段 C：逐日填充剩余配额（白/休） ==============
            // 对每天：
            //   1) 找出"前 5 天连续都已上班、若再上就第 6 连"的人，强制休；
            //   2) 其余按 (节假日休不足度 → 总休不足度 → 连休抑制) 评分挑剩余配额；
            //   3) 没被挑中的剩余非手动空格 → 设为白班。
            for (int d = 0; d < days; d++)
            {
                double rem = RemainingSlots(d);
                if (rem < -Epsilon)
                {
                    result.Warnings.Add($"{sched.Year}-{sched.Month:00}-{d + 1:00} 已分配休息 {FormatNumber(-rem)} 天超出配额。");
                    rem = 0;
                }

                var unassigned = activeIndexes
                    .Where(i => !sched.Employees[i].Cells[d].IsManual
                                && !preserved.Contains((i, d))
                                && string.IsNullOrEmpty(sched.Employees[i].Cells[d].Code))
                    .ToList();

                // 强制休：前面已经 5 连班（再上就 6 连）
                // 规范 §4.2.2 H1：实休 ≤ 总休。强制休也必须服从配额上限，超出部分留给阶段 D 做 donor-swap。
                int maxAdd = Math.Max(0, (int)Math.Floor(rem + Epsilon));
                var forcedSorted = unassigned
                    .Where(i => CountConsecutiveWorkBefore(sched, i, d) >= MaxConsecutiveWork)
                    .OrderByDescending(i => CountConsecutiveWorkBefore(sched, i, d))
                    .ThenBy(_ => rnd.Next())
                    .ToList();
                var picks = new HashSet<int>(forcedSorted.Take(maxAdd));

                int restToAdd = maxAdd - picks.Count;
                if (restToAdd > 0)
                {
                    // 评分：严格按规范优先级 1>2>3>4>5（数值越大越应该把今天的休息分给他）
                    var scored = unassigned
                        .Where(i => !picks.Contains(i))
                        .Select(i =>
                        {
                            double score = 0;
                            // P2 周末配对：周日且周六已休 → 强力（大于 P4/P5 一切修正项）
                            if (sched.DateOf(d).DayOfWeek == DayOfWeek.Sunday && d > 0
                                && ShiftCodes.RestDays(sched.Employees[i].Cells[d - 1].Code) >= 1.0)
                                score += 1000;
                            // P3 把休息放在节假日：今天是节假日则整体加权（让"被动凑数"也优先选今天的人）
                            if (isHoliday[d]) score += 80;
                            // P4 节假日休 ≥ 4：高于 P5
                            if (isHoliday[d])
                                score += Math.Max(0, TargetHolidayRest - CountHolidayRested(i)) * 200;
                            // P5 总休 ≥ 9：基线
                            score += Math.Max(0, TargetTotalRest - TotalRest(i)) * 50;
                            var previousFullRest = d > 0 && ShiftCodes.RestDays(sched.Employees[i].Cells[d - 1].Code) >= 1.0;
                            var nextFullRest = d + 1 < days && ShiftCodes.RestDays(sched.Employees[i].Cells[d + 1].Code) >= 1.0;
                            if (previousFullRest || nextFullRest) score += 12;
                            if (!previousFullRest && !nextFullRest) score -= 12;
                            if (d >= 2
                                && ShiftCodes.RestDays(sched.Employees[i].Cells[d - 2].Code) >= 1.0
                                && ShiftCodes.IsWork(sched.Employees[i].Cells[d - 1].Code))
                                score -= 12;
                            // 抑制 3 连休
                            int prevRest = CountConsecutiveRestBefore(sched, i, d);
                            if (prevRest >= 2) score -= 300;
                            // 随机扰动（±10），在不影响优先级大方向的前提下让重复点击有不同结果
                            score += rnd.NextDouble() * 20.0 - 10.0;
                            return new { i, score };
                        })
                        .OrderByDescending(x => x.score)
                        .ThenBy(x => TotalRest(x.i))                  // 同分时偏向总休最少的
                        .ThenBy(x => CountHolidayRested(x.i))         // 再次同分则偏向节假日休最少的
                        .ThenBy(x => rnd.Next())                      // 最终并列再随机
                        .Take(restToAdd);

                    foreach (var x in scored) picks.Add(x.i);
                }

                foreach (var i in unassigned)
                {
                    var cell = sched.Employees[i].Cells[d];
                    var newCode = picks.Contains(i) ? ShiftCodes.Rest : ShiftCodes.Day;
                    if (cell.Code != newCode) filled++;
                    cell.Code = newCode;
                }
            }

            // ============== 阶段 D：连续 >5 修复（保留原实现） ==============
            filled += RepairConsecutiveRuns(sched, result.Warnings, activeIndexes, preserved);

            // ============== 阶段 E：均衡 swap（让总休、节假日休分布更均匀） ==============
            BalanceRestPass(sched, isHoliday, activeIndexes, preserved);

            // ============== 阶段 F：连续上班 ≤ 5 兜底（H2 硬约束） ==============
            // 均衡 swap 已自带 ≤5 检查不会引入新违规；这里兜底处理阶段 D 因 quota=actual 卡死而残留的旧违规。
            filled += RepairConsecutiveRuns(sched, result.Warnings, activeIndexes, preserved);

            // ============== 校验与汇总 ==============
            var hardIssues = CollectHardConstraintIssues(sched, activeIndexes).ToList();
            foreach (var issue in hardIssues) result.Warnings.Add(issue);
            result.Success = hardIssues.Count == 0;
            result.FilledCells = filled;
            result.Message = result.Success
                ? $"已优化 {filled} 个单元格。"
                : $"已优化 {filled} 个单元格，但仍有 {hardIssues.Count} 项硬约束未满足。";
            if (result.Success)
            {
                sched.GeneratedAt = DateTime.Now;
            }
            return result;
        }

        // ============== 均衡 swap：在不改变每日总休量、不破坏 5 连班、不动手动格的前提下做交换 ==============
        // 双目标：先压缩总休方差，再压缩节假日休方差。每轮挑当前差距最大者。
        private static void BalanceRestPass(ScheduleVersion sched, bool[] isHoliday, List<int> activeIndexes, HashSet<(int emp, int day)> preserved)
        {
            int days = sched.DayCount;
            int empCount = sched.Employees.Count;
            if (activeIndexes == null || activeIndexes.Count < 2) return;

            double[] Total()
            {
                var t = new double[empCount];
                for (int i = 0; i < empCount; i++)
                    foreach (var c in sched.Employees[i].Cells) t[i] += ShiftCodes.RestDays(c.Code);
                return t;
            }
            double[] HolidayTotal()
            {
                var t = new double[empCount];
                for (int i = 0; i < empCount; i++)
                    for (int d = 0; d < days; d++)
                        if (isHoliday[d]) t[i] += ShiftCodes.RestDays(sched.Employees[i].Cells[d].Code);
                return t;
            }

            // ---- Round 1: 总休均衡（差距 > 1 才交换） ----
            int guard = empCount * empCount * 4 + 64;
            while (guard-- > 0)
            {
                var tot = Total();
                int hi = activeIndexes[0], lo = activeIndexes[0];
                foreach (var i in activeIndexes)
                {
                    if (tot[i] > tot[hi]) hi = i;
                    if (tot[i] < tot[lo]) lo = i;
                }
                if (tot[hi] - tot[lo] <= 1.0 + Epsilon) break;

                if (!TrySwapRestDay(sched, hi, lo, days, preserved)) break;
            }

            // ---- Round 2: 节假日休均衡（差距 > 1 才交换） ----
            guard = empCount * empCount * 4 + 64;
            while (guard-- > 0)
            {
                var hRest = HolidayTotal();
                int hi = activeIndexes[0], lo = activeIndexes[0];
                foreach (var i in activeIndexes)
                {
                    if (hRest[i] > hRest[hi]) hi = i;
                    if (hRest[i] < hRest[lo]) lo = i;
                }
                if (hRest[hi] - hRest[lo] <= 1.0 + Epsilon) break;

                // 只在节假日上做交换，避免破坏总休均衡
                if (!TrySwapRestDayOnDays(sched, hi, lo, isHoliday, requireHoliday: true, days, preserved)) break;
            }
        }

        /// <summary>把 hi 在某天的 自动休 和 lo 在同一天的 自动白 交换；交换后两人都不超 5 连班才算成功。</summary>
        private static bool TrySwapRestDay(ScheduleVersion sched, int hi, int lo, int days, HashSet<(int emp, int day)> preserved)
        {
            return TrySwapRestDayOnDays(sched, hi, lo, null, requireHoliday: false, days, preserved);
        }

        private static bool TrySwapRestDayOnDays(ScheduleVersion sched, int hi, int lo, bool[] isHoliday, bool requireHoliday, int days, HashSet<(int emp, int day)> preserved)
        {
            // 随机化访问顺序，避免每次都从 d=0 开始固定挑同一天
            var order = Enumerable.Range(0, days).OrderBy(_ => Guid.NewGuid()).ToList();
            foreach (var d in order)
            {
                if (requireHoliday && (isHoliday == null || !isHoliday[d])) continue;
                if (preserved.Contains((hi, d)) || preserved.Contains((lo, d))) continue;
                var ch = sched.Employees[hi].Cells[d];
                var cl = sched.Employees[lo].Cells[d];
                if (ch.IsManual || cl.IsManual) continue;
                if (ch.Code != ShiftCodes.Rest || cl.Code != ShiftCodes.Day) continue;

                ch.Code = ShiftCodes.Day;
                cl.Code = ShiftCodes.Rest;

                if (ComputeMaxConsecutiveWork(sched.Employees[hi]) <= MaxConsecutiveWork
                    && ComputeMaxConsecutiveWork(sched.Employees[lo]) <= MaxConsecutiveWork)
                {
                    return true;
                }

                ch.Code = ShiftCodes.Rest;
                cl.Code = ShiftCodes.Day;
            }
            return false;
        }

        // 兼容用：把原来 for(int d=0; d<days; d++) 的尾大括号改成新逻辑后这里收口

        // ============== 以下为原有辅助方法 ==============

        private static int RepairConsecutiveRuns(ScheduleVersion sched, List<string> warnings, List<int> activeIndexes, HashSet<(int emp, int day)> preserved)
        {
            int repairs = 0;
            int guard = Math.Max(1, sched.DayCount * Math.Max(1, activeIndexes.Count) * 4);
            while (guard-- > 0)
            {
                if (!TryFindOverlongRun(sched, activeIndexes, out var empIdx, out var startDay, out var endDay, out var length))
                {
                    return repairs;
                }

                var repaired = false;
                foreach (var day in CandidateBreakDays(startDay, endDay))
                {
                    var cell = sched.Employees[empIdx].Cells[day];
                    if (preserved.Contains((empIdx, day))) continue;
                    if (cell.IsManual || !ShiftCodes.IsWork(cell.Code)) continue;

                    if (TryBreakRunOnDay(sched, empIdx, day, activeIndexes, preserved))
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

        private static bool TryBreakRunOnDay(ScheduleVersion sched, int empIdx, int day, List<int> activeIndexes, HashSet<(int emp, int day)> preserved)
        {
            var quota = day < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[day] : 0;
            var actual = ComputeColumnRestCount(sched, day, activeIndexes);
            var offender = sched.Employees[empIdx].Cells[day];

            // 已超员：不能再加休，但仍可尝试 donor-swap（列中性，能改善连班）
            // 未超员且直加不破上限：直接加最简单
            if (actual + 1.0 <= quota + Epsilon)
            {
                offender.Code = ShiftCodes.Rest;
                return true;
            }

            // 直加会超 quota：用 donor-swap（offender 转休 +1，donor 转白 -1，列总数不变）
            var donorIndexes = activeIndexes
                .Where(i => i != empIdx)
                .Where(i =>
                {
                    if (preserved.Contains((i, day))) return false;
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

                if (ComputeMaxConsecutiveWork(sched.Employees[empIdx]) <= MaxConsecutiveWork &&
                    ComputeMaxConsecutiveWork(sched.Employees[donorIdx]) <= MaxConsecutiveWork)
                {
                    return true;
                }

                offender.Code = oldOffender;
                donor.Code = oldDonor;
            }

            return false;
        }

        private static bool TryFindOverlongRun(ScheduleVersion sched, List<int> activeIndexes, out int employeeIndex, out int startDay, out int endDay, out int length)
        {
            foreach (var i in activeIndexes)
            {
                var run = 0;
                for (int d = 0; d < sched.DayCount; d++)
                {
                    if (ShiftCodes.IsWork(sched.Employees[i].Cells[d].Code))
                    {
                        run++;
                        if (run > MaxConsecutiveWork)
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

        private static IEnumerable<string> CollectHardConstraintIssues(ScheduleVersion sched, List<int> activeIndexes)
        {
            // 规范 §二：每日实休必须落在 [总休-0.5, 总休]。
            for (int d = 0; d < sched.DayCount; d++)
            {
                var actual = ComputeColumnRestCount(sched, d, activeIndexes);
                var quota = d < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[d] : 0;
                var minAllowed = Math.Max(0, quota - 0.5);
                if (actual > quota + Epsilon)
                {
                    yield return $"{sched.Year}-{sched.Month:00}-{d + 1:00} 实际休息 {FormatNumber(actual)} / 总休 {FormatNumber(quota)}，多排 {FormatNumber(actual - quota)}（超员）。";
                }
                else if (actual < minAllowed - Epsilon)
                {
                    yield return $"{sched.Year}-{sched.Month:00}-{d + 1:00} 实际休息 {FormatNumber(actual)} / 总休 {FormatNumber(quota)}，低于允许下限 {FormatNumber(minAllowed)}。";
                }
            }

            foreach (var i in activeIndexes)
            {
                var maxRun = ComputeMaxConsecutiveWork(sched.Employees[i]);
                if (maxRun > MaxConsecutiveWork)
                {
                    yield return $"{sched.Employees[i].Name} 连续上班 {maxRun} 天（应≤{MaxConsecutiveWork}）。";
                }

                var totalRest = ComputeTotalRest(sched.Employees[i]);
                if (totalRest < TargetTotalRest - Epsilon)
                {
                    yield return $"{sched.Employees[i].Name} 本月休息 {FormatNumber(totalRest)} 天，低于硬性要求 {TargetTotalRest} 天。";
                }
            }
        }

        private static double ComputeColumnRestCount(ScheduleVersion sched, int dayIdx, List<int> activeIndexes)
        {
            double sum = 0;
            foreach (var i in activeIndexes)
            {
                var emp = sched.Employees[i];
                if (emp.Cells != null && dayIdx < emp.Cells.Count)
                    sum += ShiftCodes.RestDays(emp.Cells[dayIdx].Code);
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
                else run = 0;
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
    }
}
