using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MyTools.Services
{
    /// <summary>
    /// 把排班表导出为带样式的 .xlsx。
    /// 不依赖 Office / OpenXML SDK，自行写 OOXML 包。
    /// 样式保持与 SchedulePage 视觉一致：
    /// - 周末 / 节假日列：表头淡橙底
    /// - 班次单元格：按 ShiftCodes 上色（夜/小用深底白字）
    /// - 姓名列、统计列：粗体居中
    /// - 全表四向细边框
    /// </summary>
    public static class ScheduleExcelExporter
    {
        // 颜色（RRGGBB）
        private const string HeaderBg = "F1F5F9";
        private const string HolidayBg = "FEF3C7";
        private const string CellBorder = "E2E8F0";
        private const string White = "FFFFFF";
        private const string EmptyHolidayBg = "FFFBEB";

        // Shift code → (bg, fg)
        private static readonly Dictionary<string, (string bg, string fg)> ShiftColors = new Dictionary<string, (string, string)>
        {
            { "白", (White, "000000") },
            { "卡", ("E0E7FF", "000000") },
            { "副", ("C0F0FC", "000000") },
            { "感", ("FEE2E2", "000000") },
            { "大", ("1E293B", White) },
            { "小", ("475569", White) },
            { "休", ("D1FAE5", "000000") },
            { "公", ("A7F3D0", "000000") },
            { "午", ("FEF3C7", "000000") },
        };

        // 节假日列"白"格子的蓝底
        private const string HolidayDayBg = "DBEAFE";

        private static readonly string[] DowZh = { "日", "一", "二", "三", "四", "五", "六" };

        public static async Task ExportAsync(ScheduleVersion sched, string filePath)
        {
            if (sched == null) throw new ArgumentNullException(nameof(sched));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException(nameof(filePath));

            Directory.CreateDirectory(Path.GetDirectoryName(filePath) ?? AppDomain.CurrentDomain.BaseDirectory);

            var styleTable = StyleTable.Build();

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create, false))
            {
                WriteText(zip, "[Content_Types].xml", ContentTypesXml());
                WriteText(zip, "_rels/.rels", RootRelsXml());
                WriteText(zip, "docProps/app.xml", AppXml());
                WriteText(zip, "docProps/core.xml", CoreXml());
                WriteText(zip, "xl/workbook.xml", WorkbookXml($"{sched.Year}-{sched.Month:00} {sched.VersionName}"));
                WriteText(zip, "xl/_rels/workbook.xml.rels", WorkbookRelsXml());
                WriteText(zip, "xl/styles.xml", styleTable.BuildStylesXml());

                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
                using (var s = entry.Open())
                using (var w = XmlWriter.Create(s, new XmlWriterSettings { Async = true, Encoding = new UTF8Encoding(false), CloseOutput = false }))
                {
                    await WriteWorksheetAsync(w, sched, styleTable).ConfigureAwait(false);
                }
            }
        }

        // ============================ Worksheet ============================
        private static async Task WriteWorksheetAsync(XmlWriter w, ScheduleVersion sched, StyleTable styles)
        {
            int days = sched.DayCount;
            // 列：1=姓名 | 2..days+1=日 | days+2=上班 | days+3=休息 | days+4=连上
            int totalCols = 1 + days + 3;

            w.WriteStartDocument(true);
            w.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            w.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            // ----- 冻结窗口 -----
            // OpenXML 顺序要求 sheetViews 在 cols 之前，否则 Excel 会提示修复工作簿。
            w.WriteStartElement("sheetViews");
            w.WriteStartElement("sheetView");
            w.WriteAttributeString("workbookViewId", "0");
            w.WriteStartElement("pane");
            w.WriteAttributeString("xSplit", "1");
            w.WriteAttributeString("ySplit", "4");
            w.WriteAttributeString("topLeftCell", "B5");
            w.WriteAttributeString("activePane", "bottomRight");
            w.WriteAttributeString("state", "frozen");
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();

            // ----- 列宽 -----
            w.WriteStartElement("cols");
            // 姓名列
            WriteCol(w, 1, 1, 10.5);
            // 日期列（按是否节假日，宽度都 4.2）
            WriteCol(w, 2, 1 + days, 4.2);
            // 统计列
            WriteCol(w, 2 + days, totalCols, 7.5);
            w.WriteEndElement();

            w.WriteStartElement("sheetData");

            // ---------- 第1行：星期 ----------
            w.WriteStartElement("row");
            w.WriteAttributeString("r", "1");
            w.WriteAttributeString("ht", "20");
            w.WriteAttributeString("customHeight", "1");
            WriteStr(w, Cell(1, 1), "姓名", styles.S("header"));
            for (int d = 0; d < days; d++)
            {
                var date = sched.DateOf(d);
                var holiday = HolidayService.IsHoliday(date);
                WriteStr(w, Cell(2 + d, 1), DowZh[(int)date.DayOfWeek], styles.S(holiday ? "headerHoliday" : "header"));
            }
            WriteStr(w, Cell(2 + days, 1), "上班", styles.S("header"));
            WriteStr(w, Cell(3 + days, 1), "休息", styles.S("header"));
            WriteStr(w, Cell(4 + days, 1), "连上", styles.S("header"));
            w.WriteEndElement();

            // ---------- 第2行：日期 ----------
            w.WriteStartElement("row");
            w.WriteAttributeString("r", "2");
            w.WriteAttributeString("ht", "20");
            w.WriteAttributeString("customHeight", "1");
            WriteStr(w, Cell(1, 2), "日期", styles.S("header"));
            for (int d = 0; d < days; d++)
            {
                var date = sched.DateOf(d);
                var holiday = HolidayService.IsHoliday(date);
                WriteNum(w, Cell(2 + d, 2), (d + 1).ToString(CultureInfo.InvariantCulture), styles.S(holiday ? "headerHoliday" : "header"));
            }
            WriteStr(w, Cell(2 + days, 2), "", styles.S("header"));
            WriteStr(w, Cell(3 + days, 2), "", styles.S("header"));
            WriteStr(w, Cell(4 + days, 2), "", styles.S("header"));
            w.WriteEndElement();

            // ---------- 第3行：实休 ----------
            w.WriteStartElement("row");
            w.WriteAttributeString("r", "3");
            WriteStr(w, Cell(1, 3), "实休", styles.S("header"));
            for (int d = 0; d < days; d++)
            {
                var actual = ComputeColumnRestCount(sched, d);
                var quota = d < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[d] : 0;
                WriteNum(w, Cell(2 + d, 3), FormatStat(actual), styles.S(IsDailyRestWithinQuota(actual, quota) ? "statGood" : "statBad"));
            }
            WriteStr(w, Cell(2 + days, 3), "", styles.S("header"));
            WriteStr(w, Cell(3 + days, 3), "", styles.S("header"));
            WriteStr(w, Cell(4 + days, 3), "", styles.S("header"));
            w.WriteEndElement();

            w.WriteStartElement("row");
            w.WriteAttributeString("r", "4");
            WriteStr(w, Cell(1, 4), "总休", styles.S("header"));
            for (int d = 0; d < days; d++)
            {
                var actual = ComputeColumnRestCount(sched, d);
                var quota = d < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[d] : 0;
                WriteNum(w, Cell(2 + d, 4), FormatStat(quota), styles.S(IsDailyRestWithinQuota(actual, quota) ? "statGood" : "statBad"));
            }
            WriteStr(w, Cell(2 + days, 4), "", styles.S("header"));
            WriteStr(w, Cell(3 + days, 4), "", styles.S("header"));
            WriteStr(w, Cell(4 + days, 4), "", styles.S("header"));
            w.WriteEndElement();

            // ---------- 员工行 ----------
            int rowIdx = 4;
            foreach (var emp in sched.Employees)
            {
                if (string.IsNullOrWhiteSpace(emp.Name)) continue;
                rowIdx++;
                w.WriteStartElement("row");
                w.WriteAttributeString("r", rowIdx.ToString(CultureInfo.InvariantCulture));

                // 姓名
                WriteStr(w, Cell(1, rowIdx), emp.Name ?? string.Empty, styles.S("name"));

                // 每日格
                double work = 0, rest = 0;
                int maxRun = 0, run = 0;
                for (int d = 0; d < days; d++)
                {
                    var cell = d < emp.Cells.Count ? emp.Cells[d] : new ShiftCell();
                    var code = ShiftCodes.Normalize(cell.Code);
                    var date = sched.DateOf(d);
                    var holiday = HolidayService.IsHoliday(date);
                    string styleKey;
                    string displayText;

                    if (code == "白")
                    {
                        // 节假日列白格子用浅蓝底+无字；非节假日列用白底+无字
                        styleKey = holiday ? "holidayDayCell" : "empty";
                        displayText = "";
                    }
                    else if (!string.IsNullOrEmpty(code) && ShiftColors.ContainsKey(code))
                    {
                        styleKey = "shift_" + code;
                        displayText = DisplayShiftCode(code);
                    }
                    else
                    {
                        styleKey = holiday ? "emptyHoliday" : "empty";
                        displayText = "";
                    }
                    WriteStr(w, Cell(2 + d, rowIdx), displayText, styles.S(styleKey));

                    work += ShiftCodes.WorkDays(code);
                    rest += ShiftCodes.RestDays(code);
                    if (ShiftCodes.IsWork(code)) { run++; if (run > maxRun) maxRun = run; }
                    else run = 0;
                }

                // 统计
                WriteNum(w, Cell(2 + days, rowIdx), FormatStat(work), styles.S("stat"));
                WriteNum(w, Cell(3 + days, rowIdx), FormatStat(rest), styles.S(rest < 9 ? "statBad" : "stat"));
                WriteNum(w, Cell(4 + days, rowIdx), maxRun.ToString(CultureInfo.InvariantCulture), styles.S(maxRun > 5 ? "statBad" : "stat"));

                w.WriteEndElement();
            }

            w.WriteEndElement(); // sheetData
            w.WriteEndElement(); // worksheet
            w.WriteEndDocument();
            await w.FlushAsync().ConfigureAwait(false);
        }

        private static string FormatStat(double v)
        {
            return v % 1 == 0 ? ((int)v).ToString(CultureInfo.InvariantCulture)
                              : v.ToString("0.#", CultureInfo.InvariantCulture);
        }

        private static string DisplayShiftCode(string code)
        {
            code = ShiftCodes.Normalize(code);
            return code == ShiftCodes.Big ? "夜" : code;
        }

        private static bool IsDailyRestWithinQuota(double actual, double quota)
        {
            return actual >= Math.Max(0, quota - 0.5) - 0.001 && actual <= quota + 0.001;
        }

        private static double ComputeColumnRestCount(ScheduleVersion sched, int dayIdx)
        {
            if (sched == null || dayIdx < 0)
            {
                return 0;
            }

            double sum = 0;
            foreach (var emp in sched.Employees)
            {
                if (string.IsNullOrWhiteSpace(emp.Name)) continue;
                if (emp.Cells != null && dayIdx < emp.Cells.Count)
                {
                    sum += ShiftCodes.RestDays(emp.Cells[dayIdx].Code);
                }
            }

            return sum;
        }

        // ============================ Cell helpers ============================
        private static string Cell(int col, int row)
        {
            var sb = new StringBuilder();
            int c = col;
            while (c > 0)
            {
                int mod = (c - 1) % 26;
                sb.Insert(0, (char)('A' + mod));
                c = (c - mod) / 26;
            }
            return sb.ToString() + row.ToString(CultureInfo.InvariantCulture);
        }

        private static void WriteCol(XmlWriter w, int min, int max, double width)
        {
            w.WriteStartElement("col");
            w.WriteAttributeString("min", min.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("max", max.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("width", width.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("customWidth", "1");
            w.WriteEndElement();
        }

        private static void WriteStr(XmlWriter w, string cellRef, string text, int styleIndex)
        {
            w.WriteStartElement("c");
            w.WriteAttributeString("r", cellRef);
            w.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("t", "inlineStr");
            w.WriteStartElement("is");
            w.WriteStartElement("t");
            if (!string.IsNullOrEmpty(text) && (text[0] == ' ' || text[text.Length - 1] == ' '))
                w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(Sanitize(text));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteNum(XmlWriter w, string cellRef, string value, int styleIndex)
        {
            w.WriteStartElement("c");
            w.WriteAttributeString("r", cellRef);
            w.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
            w.WriteElementString("v", value);
            w.WriteEndElement();
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var sb = new StringBuilder(value.Length);
            foreach (var ch in value) if (XmlConvert.IsXmlChar(ch)) sb.Append(ch);
            return sb.ToString();
        }

        private static void WriteText(ZipArchive zip, string path, string content)
        {
            var e = zip.CreateEntry(path, CompressionLevel.Fastest);
            using (var s = e.Open())
            using (var sw = new StreamWriter(s, new UTF8Encoding(false))) sw.Write(content);
        }

        // ============================ Style table ============================
        /// <summary>
        /// 集中管理 fonts / fills / borders / cellXfs；
        /// 通过 string Key 拿 cellXf 索引。
        /// </summary>
        private sealed class StyleTable
        {
            private readonly List<string> _fontsXml = new List<string>();
            private readonly List<string> _fillsXml = new List<string>();
            private readonly List<string> _xfsXml = new List<string>();
            private readonly Dictionary<string, int> _keyToIndex = new Dictionary<string, int>();

            public int S(string key) => _keyToIndex[key];

            public static StyleTable Build()
            {
                var t = new StyleTable();

                // 字体：0=黑色, 1=白色
                t._fontsXml.Add("<font><sz val=\"11\"/><name val=\"微软雅黑\"/><color rgb=\"FF000000\"/></font>");
                t._fontsXml.Add("<font><sz val=\"11\"/><name val=\"微软雅黑\"/><b/><color rgb=\"FF000000\"/></font>");
                t._fontsXml.Add("<font><sz val=\"11\"/><name val=\"微软雅黑\"/><b/><color rgb=\"FFFFFFFF\"/></font>");
                t._fontsXml.Add("<font><sz val=\"11\"/><name val=\"微软雅黑\"/><b/><color rgb=\"FFC62828\"/></font>"); // 红色统计
                const int FONT_REG = 0, FONT_BOLD = 1, FONT_BOLD_WHITE = 2, FONT_BOLD_RED = 3;

                // 填充：0/1 系统占位（none / gray125），后续动态
                t._fillsXml.Add("<fill><patternFill patternType=\"none\"/></fill>");
                t._fillsXml.Add("<fill><patternFill patternType=\"gray125\"/></fill>");
                int FillSolid(string rgb)
                {
                    var idx = t._fillsXml.Count;
                    t._fillsXml.Add($"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{rgb}\"/><bgColor indexed=\"64\"/></patternFill></fill>");
                    return idx;
                }

                int fillHeader = FillSolid(HeaderBg);
                int fillHoliday = FillSolid(HolidayBg);

                var shiftFillIdx = new Dictionary<string, int>();
                var shiftFontIdx = new Dictionary<string, int>();
                foreach (var kv in ShiftColors)
                {
                    shiftFillIdx[kv.Key] = FillSolid(kv.Value.bg);
                    // 字体：白底用黑字、深底用白字
                    shiftFontIdx[kv.Key] = string.Equals(kv.Value.fg, White, StringComparison.OrdinalIgnoreCase) ? FONT_BOLD_WHITE : FONT_BOLD;
                }
                int fillEmptyHoliday = FillSolid(EmptyHolidayBg);
                int fillHolidayDay = FillSolid(HolidayDayBg);

                // 边框：1=四向细边
                const int BORDER_THIN = 1;

                int RegisterXf(int fontId, int fillId, int borderId, string customAlign = null)
                {
                    var xf = new StringBuilder("<xf");
                    xf.Append($" numFmtId=\"0\" fontId=\"{fontId}\" fillId=\"{fillId}\" borderId=\"{borderId}\" xfId=\"0\"");
                    xf.Append(" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"");
                    if (customAlign != null) xf.Append(" applyAlignment=\"1\"");
                    xf.Append(">");
                    if (customAlign != null) xf.Append(customAlign);
                    xf.Append("</xf>");
                    t._xfsXml.Add(xf.ToString());
                    return t._xfsXml.Count - 1;
                }

                const string CenterAlign = "<alignment horizontal=\"center\" vertical=\"center\"/>";

                t._keyToIndex["header"] = RegisterXf(FONT_BOLD, fillHeader, BORDER_THIN, CenterAlign);
                t._keyToIndex["headerHoliday"] = RegisterXf(FONT_BOLD, fillHoliday, BORDER_THIN, CenterAlign);
                t._keyToIndex["name"] = RegisterXf(FONT_BOLD, fillHeader, BORDER_THIN, CenterAlign);
                t._keyToIndex["empty"] = RegisterXf(FONT_REG, 0, BORDER_THIN, CenterAlign);
                t._keyToIndex["emptyHoliday"] = RegisterXf(FONT_REG, fillEmptyHoliday, BORDER_THIN, CenterAlign);
                t._keyToIndex["holidayDayCell"] = RegisterXf(FONT_REG, fillHolidayDay, BORDER_THIN, CenterAlign);
                t._keyToIndex["stat"] = RegisterXf(FONT_BOLD, 0, BORDER_THIN, CenterAlign);
                t._keyToIndex["statGood"] = RegisterXf(FONT_BOLD, fillEmptyHoliday, BORDER_THIN, CenterAlign);
                t._keyToIndex["statBad"] = RegisterXf(FONT_BOLD_RED, 0, BORDER_THIN, CenterAlign);

                foreach (var kv in ShiftColors)
                {
                    var key = "shift_" + kv.Key;
                    t._keyToIndex[key] = RegisterXf(shiftFontIdx[kv.Key], shiftFillIdx[kv.Key], BORDER_THIN, CenterAlign);
                }

                return t;
            }

            public string BuildStylesXml()
            {
                var sb = new StringBuilder();
                sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sb.AppendLine("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                sb.AppendLine($"<fonts count=\"{_fontsXml.Count}\">{string.Concat(_fontsXml)}</fonts>");
                sb.AppendLine($"<fills count=\"{_fillsXml.Count}\">{string.Concat(_fillsXml)}</fills>");
                sb.AppendLine("<borders count=\"2\">");
                sb.AppendLine("<border/>");
                sb.AppendLine($"<border><left style=\"thin\"><color rgb=\"FF{CellBorder}\"/></left><right style=\"thin\"><color rgb=\"FF{CellBorder}\"/></right><top style=\"thin\"><color rgb=\"FF{CellBorder}\"/></top><bottom style=\"thin\"><color rgb=\"FF{CellBorder}\"/></bottom></border>");
                sb.AppendLine("</borders>");
                sb.AppendLine("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
                sb.AppendLine($"<cellXfs count=\"{_xfsXml.Count}\">{string.Concat(_xfsXml)}</cellXfs>");
                sb.AppendLine("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
                sb.AppendLine("</styleSheet>");
                return sb.ToString();
            }
        }

        // ============================ Static XMLs ============================
        private static string ContentTypesXml() => @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
  <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/>
  <Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml""/>
</Types>";

        private static string RootRelsXml() => @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/>
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml""/>
</Relationships>";

        private static string AppXml() => @"<?xml version=""1.0"" encoding=""utf-8""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties""><Application>MyTools</Application></Properties>";

        private static string CoreXml()
        {
            var t = XmlConvert.ToString(DateTime.UtcNow, XmlDateTimeSerializationMode.Utc);
            return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""
                   xmlns:dc=""http://purl.org/dc/elements/1.1/""
                   xmlns:dcterms=""http://purl.org/dc/terms/""
                   xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <dc:creator>MyTools</dc:creator>
  <cp:lastModifiedBy>MyTools</cp:lastModifiedBy>
  <dcterms:created xsi:type=""dcterms:W3CDTF"">{t}</dcterms:created>
  <dcterms:modified xsi:type=""dcterms:W3CDTF"">{t}</dcterms:modified>
</cp:coreProperties>";
        }

        private static string WorkbookXml(string sheetName) => $@"<?xml version=""1.0"" encoding=""utf-8""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets><sheet name=""{SecurityElement.Escape(NormalizeSheetName(sheetName))}"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>";

        private static string WorkbookRelsXml() => @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>";

        private static string NormalizeSheetName(string value)
        {
            var name = string.IsNullOrWhiteSpace(value) ? "Sheet1" : value.Trim();
            foreach (var c in new[] { '\\', '/', '?', '*', '[', ']', ':' }) name = name.Replace(c, '_');
            return name.Length <= 31 ? name : name.Substring(0, 31);
        }
    }
}
