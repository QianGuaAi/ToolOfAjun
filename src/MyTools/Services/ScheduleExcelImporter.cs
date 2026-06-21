using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MyTools.Services
{
    public sealed class ScheduleExcelImportResult
    {
        public ScheduleExcelImportResult(ScheduleVersion schedule, IList<string> warnings)
        {
            Schedule = schedule ?? throw new ArgumentNullException(nameof(schedule));
            Warnings = (warnings ?? new List<string>()).ToList().AsReadOnly();
        }

        public ScheduleVersion Schedule { get; }
        public IReadOnlyList<string> Warnings { get; }
    }

    /// <summary>
    /// Reads the lightweight OOXML schedule format written by ScheduleExcelExporter.
    /// </summary>
    public static class ScheduleExcelImporter
    {
        private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace PackageRelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace DocumentRelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly Regex YearMonthInFileName = new Regex(
            @"(?<!\d)(?<year>19\d{2}|20\d{2}|21\d{2})[-_.年](?<month>0?[1-9]|1[0-2])(?!\d)",
            RegexOptions.Compiled);

        public static Task<ScheduleExcelImportResult> ImportAsync(string filePath)
        {
            return Task.Run(() => ImportCore(filePath, null, null, null));
        }

        public static Task<ScheduleExcelImportResult> ImportAsync(string filePath, int fallbackYear, int fallbackMonth, string versionName)
        {
            return Task.Run(() => ImportCore(filePath, fallbackYear, fallbackMonth, versionName));
        }

        private static ScheduleExcelImportResult ImportCore(string filePath, int? fallbackYear, int? fallbackMonth, string versionName)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("导入文件路径不能为空。", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到要导入的 Excel 文件。", filePath);
            }

            if (!string.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("仅支持导入 .xlsx 排班文件。");
            }

            var warnings = new List<string>();
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, false))
            {
                if (zip.GetEntry("xl/workbook.xml") == null)
                {
                    throw new InvalidDataException("Excel 文件结构不完整：缺少 xl/workbook.xml。");
                }

                var sharedStrings = ReadSharedStrings(zip);
                var cells = ReadWorksheetCells(zip, sharedStrings, out var maxRow);

                RequireHeader(cells, 1, 1, "姓名");
                RequireHeader(cells, 2, 1, "日期");
                RequireHeader(cells, 4, 1, "总休");

                var dayCount = ReadDayCount(cells);
                ResolveYearMonth(filePath, dayCount, fallbackYear, fallbackMonth, warnings, out var year, out var month);

                var schedule = new ScheduleVersion
                {
                    Year = year,
                    Month = month,
                    VersionName = BuildVersionName(versionName, filePath),
                    DailyRestQuotas = ReadDailyRestQuotas(cells, dayCount),
                    Employees = ReadEmployees(cells, maxRow, dayCount, warnings),
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now,
                    GeneratedAt = DateTime.Now
                };

                if (schedule.Employees.Count == 0)
                {
                    throw new InvalidDataException("Excel 中没有可导入的人员行。");
                }

                return new ScheduleExcelImportResult(schedule, warnings);
            }
        }

        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            var result = new List<string>();
            if (entry == null)
            {
                return result;
            }

            using (var stream = entry.Open())
            {
                var doc = XDocument.Load(stream);
                foreach (var item in doc.Descendants(SpreadsheetNs + "si"))
                {
                    result.Add(string.Concat(item.Descendants(SpreadsheetNs + "t").Select(node => node.Value)));
                }
            }

            return result;
        }

        private static Dictionary<Tuple<int, int>, string> ReadWorksheetCells(ZipArchive zip, IList<string> sharedStrings, out int maxRow)
        {
            var worksheetPath = ResolveFirstWorksheetPath(zip);
            var entry = zip.GetEntry(worksheetPath);
            if (entry == null)
            {
                throw new InvalidDataException($"Excel 文件缺少第一个工作表 {worksheetPath}。");
            }

            var cells = new Dictionary<Tuple<int, int>, string>();
            maxRow = 0;
            using (var stream = entry.Open())
            {
                var doc = XDocument.Load(stream);
                foreach (var cell in doc.Descendants(SpreadsheetNs + "c"))
                {
                    var reference = (string)cell.Attribute("r");
                    if (!TryParseCellReference(reference, out var row, out var column))
                    {
                        continue;
                    }

                    cells[Tuple.Create(row, column)] = ReadCellValue(cell, sharedStrings);
                    if (row > maxRow)
                    {
                        maxRow = row;
                    }
                }
            }

            return cells;
        }

        private static string ResolveFirstWorksheetPath(ZipArchive zip)
        {
            const string fallback = "xl/worksheets/sheet1.xml";
            var workbookEntry = zip.GetEntry("xl/workbook.xml");
            var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (workbookEntry == null || relsEntry == null)
            {
                return fallback;
            }

            try
            {
                string relationshipId;
                using (var workbookStream = workbookEntry.Open())
                {
                    var workbook = XDocument.Load(workbookStream);
                    var firstVisibleSheet = workbook
                        .Descendants(SpreadsheetNs + "sheet")
                        .FirstOrDefault(sheet =>
                        {
                            var state = ((string)sheet.Attribute("state") ?? string.Empty).Trim();
                            return !string.Equals(state, "hidden", StringComparison.OrdinalIgnoreCase) &&
                                   !string.Equals(state, "veryHidden", StringComparison.OrdinalIgnoreCase);
                        })
                        ?? workbook.Descendants(SpreadsheetNs + "sheet").FirstOrDefault();

                    relationshipId = (string)firstVisibleSheet?.Attribute(DocumentRelationshipNs + "id");
                }

                if (string.IsNullOrWhiteSpace(relationshipId))
                {
                    return fallback;
                }

                using (var relsStream = relsEntry.Open())
                {
                    var rels = XDocument.Load(relsStream);
                    var relationship = rels
                        .Descendants(PackageRelationshipNs + "Relationship")
                        .FirstOrDefault(rel => string.Equals((string)rel.Attribute("Id"), relationshipId, StringComparison.OrdinalIgnoreCase));
                    var target = (string)relationship?.Attribute("Target");
                    return NormalizeWorksheetTarget(target) ?? fallback;
                }
            }
            catch
            {
                return fallback;
            }
        }

        private static string NormalizeWorksheetTarget(string target)
        {
            if (string.IsNullOrWhiteSpace(target))
            {
                return null;
            }

            var normalized = target.Trim().Replace('\\', '/');
            if (normalized.StartsWith("/", StringComparison.Ordinal))
            {
                normalized = normalized.TrimStart('/');
                return normalized;
            }

            return normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
                ? normalized
                : "xl/" + normalized;
        }

        private static string ReadCellValue(XElement cell, IList<string> sharedStrings)
        {
            var type = ((string)cell.Attribute("t") ?? string.Empty).Trim();
            if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
            {
                return string.Concat(cell.Descendants(SpreadsheetNs + "t").Select(node => node.Value));
            }

            var value = (string)cell.Element(SpreadsheetNs + "v") ?? string.Empty;
            if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) &&
                    index >= 0 &&
                    index < sharedStrings.Count)
                {
                    return sharedStrings[index] ?? string.Empty;
                }

                return string.Empty;
            }

            return value;
        }

        private static void RequireHeader(Dictionary<Tuple<int, int>, string> cells, int row, int column, string expected)
        {
            var actual = GetCell(cells, row, column);
            if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException($"Excel 表头不匹配：{CellName(column, row)} 应为“{expected}”，实际为“{actual}”。");
            }
        }

        private static int ReadDayCount(Dictionary<Tuple<int, int>, string> cells)
        {
            var dayCount = 0;
            for (var column = 2; column <= 40; column++)
            {
                var raw = GetCell(cells, 2, column);
                if (string.IsNullOrWhiteSpace(raw))
                {
                    break;
                }

                if (!TryParseWholeNumber(raw, out var day) || day != dayCount + 1)
                {
                    throw new InvalidDataException($"日期列格式错误：{CellName(column, 2)} 应为 {dayCount + 1}，实际为“{raw}”。");
                }

                dayCount++;
            }

            if (dayCount < 28 || dayCount > 31)
            {
                throw new InvalidDataException($"日期列数量 {dayCount} 无效，请导入包含完整月份日期列的排班 Excel。");
            }

            return dayCount;
        }

        private static void ResolveYearMonth(
            string filePath,
            int dayCount,
            int? fallbackYear,
            int? fallbackMonth,
            List<string> warnings,
            out int year,
            out int month)
        {
            var match = YearMonthInFileName.Match(Path.GetFileNameWithoutExtension(filePath) ?? string.Empty);
            if (match.Success &&
                int.TryParse(match.Groups["year"].Value, out var fileYear) &&
                int.TryParse(match.Groups["month"].Value, out var fileMonth) &&
                IsValidYearMonth(fileYear, fileMonth))
            {
                if (DateTime.DaysInMonth(fileYear, fileMonth) == dayCount)
                {
                    year = fileYear;
                    month = fileMonth;
                    return;
                }

                if (HasMatchingFallback(dayCount, fallbackYear, fallbackMonth))
                {
                    year = fallbackYear.Value;
                    month = fallbackMonth.Value;
                    warnings.Add($"文件名中的年月 {fileYear}-{fileMonth:00} 与日期列数量不匹配，已按当前排班 {year}-{month:00} 导入。");
                    return;
                }

                throw new InvalidDataException($"文件名中的年月 {fileYear}-{fileMonth:00} 与 Excel 日期列数量 {dayCount} 不匹配。");
            }

            if (HasMatchingFallback(dayCount, fallbackYear, fallbackMonth))
            {
                year = fallbackYear.Value;
                month = fallbackMonth.Value;
                return;
            }

            throw new InvalidDataException("无法识别导入年月。请先新建或加载对应月份排班后再导入，或将文件名改为包含 YYYY-MM。");
        }

        private static bool HasMatchingFallback(int dayCount, int? fallbackYear, int? fallbackMonth)
        {
            return fallbackYear.HasValue &&
                   fallbackMonth.HasValue &&
                   IsValidYearMonth(fallbackYear.Value, fallbackMonth.Value) &&
                   DateTime.DaysInMonth(fallbackYear.Value, fallbackMonth.Value) == dayCount;
        }

        private static bool IsValidYearMonth(int year, int month)
        {
            return year >= 1900 && year <= 2100 && month >= 1 && month <= 12;
        }

        private static List<double> ReadDailyRestQuotas(Dictionary<Tuple<int, int>, string> cells, int dayCount)
        {
            var quotas = new List<double>(dayCount);
            for (var day = 0; day < dayCount; day++)
            {
                var raw = GetCell(cells, 4, 2 + day);
                var quota = string.IsNullOrWhiteSpace(raw) ? 0 : ParseNumber(raw, $"第 {day + 1} 日总休");
                if (quota < 0 || double.IsNaN(quota) || double.IsInfinity(quota))
                {
                    throw new InvalidDataException($"第 {day + 1} 日总休值无效：{raw}");
                }

                if (Math.Abs(quota * 2 - Math.Round(quota * 2)) > 0.001)
                {
                    throw new InvalidDataException($"第 {day + 1} 日总休 {FormatNumber(quota)} 不是 0.5 步进。");
                }

                quotas.Add(quota);
            }

            return quotas;
        }

        private static List<EmployeeRow> ReadEmployees(
            Dictionary<Tuple<int, int>, string> cells,
            int maxRow,
            int dayCount,
            List<string> warnings)
        {
            var employees = new List<EmployeeRow>();
            var workStatColumn = 2 + dayCount;
            var canRestoreWhiteShifts = string.Equals(GetCell(cells, 1, workStatColumn), "上班", StringComparison.OrdinalIgnoreCase);
            if (!canRestoreWhiteShifts)
            {
                warnings.Add("未找到“上班”统计列，空白日期格按空白导入，不自动还原为白班。");
            }

            for (var row = 5; row <= Math.Max(maxRow, 5); row++)
            {
                var name = GetCell(cells, row, 1).Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var employee = new EmployeeRow { Name = name, Cells = new List<ShiftCell>(dayCount) };
                for (var day = 0; day < dayCount; day++)
                {
                    var raw = GetCell(cells, row, 2 + day);
                    var code = ParseShiftCode(raw, row, day + 1);
                    employee.Cells.Add(new ShiftCell
                    {
                        Code = code,
                        IsManual = !string.IsNullOrEmpty(code)
                    });
                }

                if (canRestoreWhiteShifts)
                {
                    RestoreWhiteShiftsFromWorkStat(employee, GetCell(cells, row, workStatColumn), dayCount, row, warnings);
                }

                employees.Add(employee);
            }

            return employees;
        }

        private static void RestoreWhiteShiftsFromWorkStat(EmployeeRow employee, string workStatRaw, int dayCount, int row, List<string> warnings)
        {
            if (employee == null || string.IsNullOrWhiteSpace(workStatRaw))
            {
                return;
            }

            if (!TryParseNumber(workStatRaw, out var workTarget))
            {
                warnings.Add($"第 {row} 行“上班”统计无法读取，未按统计还原白班。");
                return;
            }

            var roundedTarget = Math.Round(workTarget);
            if (Math.Abs(workTarget - roundedTarget) > 0.001 || roundedTarget < 0)
            {
                warnings.Add($"第 {row} 行“上班”统计 {workStatRaw} 不是整数，未按统计还原白班。");
                return;
            }

            var explicitWork = employee.Cells.Count(cell => ShiftCodes.IsWork(cell.Code));
            var missingWhite = (int)roundedTarget - explicitWork;
            if (missingWhite < 0)
            {
                warnings.Add($"{employee.Name} 的班次上班数已超过 Excel 统计列，请核对该行。");
                return;
            }

            if (missingWhite == 0)
            {
                return;
            }

            var restored = 0;
            for (var day = 0; day < dayCount && restored < missingWhite; day++)
            {
                var cell = employee.Cells[day];
                if (cell != null && string.IsNullOrEmpty(cell.Code))
                {
                    cell.Code = ShiftCodes.Day;
                    cell.IsManual = true;
                    restored++;
                }
            }

            if (restored > 0)
            {
                warnings.Add($"已按“上班”统计为 {employee.Name} 还原 {restored} 个白班空格。");
            }

            if (restored < missingWhite)
            {
                warnings.Add($"{employee.Name} 仍有 {missingWhite - restored} 个白班空格无法还原，请核对该行。");
            }
        }

        private static string ParseShiftCode(string raw, int row, int day)
        {
            var text = (raw ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            switch (text)
            {
                case "白班":
                    return ShiftCodes.Day;
                case "卡班":
                    return ShiftCodes.Card;
                case "副小":
                    return ShiftCodes.Deputy;
                case "感染科":
                    return ShiftCodes.Infect;
                case "夜":
                case "夜班":
                case "大夜":
                    return ShiftCodes.Big;
                case "小夜":
                    return ShiftCodes.Small;
                case "休息":
                    return ShiftCodes.Rest;
                case "公休":
                    return ShiftCodes.Public;
                case "下午休":
                case "下午休0.5":
                    return ShiftCodes.Half;
            }

            var normalized = ShiftCodes.Normalize(text);
            if (ShiftCodes.All.Contains(normalized))
            {
                return normalized;
            }

            throw new InvalidDataException($"第 {row} 行 {day} 日班次“{text}”不在支持范围：白、卡、副、感、大、小、休、公、午。");
        }

        private static string BuildVersionName(string requested, string filePath)
        {
            var value = string.IsNullOrWhiteSpace(requested)
                ? Path.GetFileNameWithoutExtension(filePath)
                : requested;

            value = SanitizeFileName(value);
            if (string.IsNullOrWhiteSpace(value))
            {
                value = "导入_" + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            }

            return value.Length <= 80 ? value : value.Substring(0, 80).Trim();
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var invalid = Path.GetInvalidFileNameChars();
            var chars = value.Trim().Select(ch => invalid.Contains(ch) || char.IsControl(ch) ? '_' : ch).ToArray();
            return new string(chars).Trim().Trim('.');
        }

        private static string GetCell(Dictionary<Tuple<int, int>, string> cells, int row, int column)
        {
            return cells.TryGetValue(Tuple.Create(row, column), out var value) ? (value ?? string.Empty).Trim() : string.Empty;
        }

        private static bool TryParseNumber(string raw, out double value)
        {
            return double.TryParse((raw ?? string.Empty).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                   double.TryParse((raw ?? string.Empty).Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static double ParseNumber(string raw, string fieldName)
        {
            if (TryParseNumber(raw, out var value))
            {
                return value;
            }

            throw new InvalidDataException($"{fieldName} 不是有效数字：{raw}");
        }

        private static bool TryParseWholeNumber(string raw, out int value)
        {
            value = 0;
            if (!TryParseNumber(raw, out var number))
            {
                return false;
            }

            var rounded = Math.Round(number);
            if (Math.Abs(number - rounded) > 0.001)
            {
                return false;
            }

            value = (int)rounded;
            return true;
        }

        private static bool TryParseCellReference(string reference, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(reference))
            {
                return false;
            }

            var index = 0;
            while (index < reference.Length && char.IsLetter(reference[index]))
            {
                column = column * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1);
                index++;
            }

            if (column <= 0 || index >= reference.Length)
            {
                return false;
            }

            return int.TryParse(reference.Substring(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
        }

        private static string CellName(int column, int row)
        {
            var letters = string.Empty;
            var value = column;
            while (value > 0)
            {
                var mod = (value - 1) % 26;
                letters = (char)('A' + mod) + letters;
                value = (value - mod) / 26;
            }

            return letters + row.ToString(CultureInfo.InvariantCulture);
        }

        private static string FormatNumber(double value)
        {
            return value % 1 == 0
                ? ((int)value).ToString(CultureInfo.InvariantCulture)
                : value.ToString("0.#", CultureInfo.InvariantCulture);
        }
    }
}
