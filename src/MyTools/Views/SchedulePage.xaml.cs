using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using MyTools.Services;
using MyTools.ViewModels;

namespace MyTools.Views
{
    public partial class SchedulePage : UserControl
    {
        private ScheduleViewModel _vm;
        private readonly DispatcherTimer _rebuildTimer;
        private int _focusedEmployeeIndex = -1;
        private int _focusedDayIndex = -1;
        private double _scheduleZoom = 1.0;
        private const double MinScheduleZoom = 0.65;
        private const double MaxScheduleZoom = 1.8;
        private const double ScheduleZoomStep = 0.1;
        private static readonly string[] DowZh = { "日", "一", "二", "三", "四", "五", "六" };

        public SchedulePage()
        {
            InitializeComponent();
            ApplyScheduleZoom();
            _rebuildTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(45) };
            _rebuildTimer.Tick += (s, e) =>
            {
                _rebuildTimer.Stop();
                Rebuild();
            };
            DataContextChanged += SchedulePage_DataContextChanged;
        }

        private void SchedulePage_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (_vm != null)
            {
                _vm.ScheduleStructureChanged -= Vm_ScheduleStructureChanged;
                _vm.ScheduleDataChanged -= Vm_ScheduleDataChanged;
                _vm.ScheduleCellFocusRequested -= Vm_ScheduleCellFocusRequested;
            }
            _vm = e.NewValue as ScheduleViewModel;
            if (_vm != null)
            {
                _vm.ScheduleStructureChanged += Vm_ScheduleStructureChanged;
                _vm.ScheduleDataChanged += Vm_ScheduleDataChanged;
                _vm.ScheduleCellFocusRequested += Vm_ScheduleCellFocusRequested;
                Rebuild();
            }
        }

        private void Vm_ScheduleStructureChanged(object sender, EventArgs e) => QueueRebuild();
        private void Vm_ScheduleDataChanged(object sender, EventArgs e) => QueueRebuild();
        private void Vm_ScheduleCellFocusRequested(object sender, ScheduleCellFocusRequestedEventArgs e)
        {
            _focusedEmployeeIndex = e.EmployeeIndex;
            _focusedDayIndex = e.DayIndex;
            QueueRebuild();
        }

        private void QueueRebuild()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(QueueRebuild));
                return;
            }

            _rebuildTimer.Stop();
            _rebuildTimer.Start();
        }

        private void VersionList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_vm == null) return;
            if (sender is ListBox lb && lb.SelectedItem is ScheduleVersionInfo info)
            {
                _ = _vm.LoadVersionAsync(info);
            }
        }

        private static readonly string[] RandomSurnames =
        {
            "王","李","张","刘","陈","杨","赵","黄","周","吴",
            "徐","孙","胡","朱","高","林","何","郭","马","罗",
            "梁","宋","郑","谢","韩","唐","冯","于","董","萧"
        };
        private static readonly string[] RandomGivenNames =
        {
            "伟","芳","娜","秀英","敏","静","丽","强","磊","军",
            "洋","勇","艳","杰","娟","涛","明","超","秀兰","霞",
            "平","刚","桂英","欣","怡","浩","宇","婷","雪","琳",
            "晨","瑶","俊","旭","睿","佳","璐","乐","萌","凯"
        };
        private static readonly Random _rand = new Random();

        private static string GenerateRandomName()
        {
            return RandomSurnames[_rand.Next(RandomSurnames.Length)] + RandomGivenNames[_rand.Next(RandomGivenNames.Length)];
        }

        private void AddEmployee_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null || _vm.Current == null)
            {
                MessageBox.Show("请先新建或加载排班表。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (!_vm.IsEditing)
            {
                MessageBox.Show("请先点编辑进入编辑模式。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int count = 1;
            if (!int.TryParse(AddCountBox?.Text, out count)) count = 1;
            if (count < 0) count = 0;
            if (count > 2000) count = 2000;
            AddCountBox.Text = count.ToString();

            if (count == 1)
            {
                _vm.AddEmployee(GenerateRandomName());
            }
            else if (count > 1)
            {
                var names = new System.Collections.Generic.List<string>(count);
                for (int i = 0; i < count; i++) names.Add(GenerateRandomName());
                _vm.AddEmployees(names);
            }
        }

        private void AddCountBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            // 仅允许数字输入
            foreach (var ch in e.Text)
            {
                if (!char.IsDigit(ch)) { e.Handled = true; return; }
            }
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e) => SetScheduleZoom(_scheduleZoom - ScheduleZoomStep);

        private void ZoomIn_Click(object sender, RoutedEventArgs e) => SetScheduleZoom(_scheduleZoom + ScheduleZoomStep);

        private void ZoomReset_Click(object sender, RoutedEventArgs e) => SetScheduleZoom(1.0);

        private void ScheduleScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
            {
                return;
            }

            SetScheduleZoom(_scheduleZoom + (e.Delta > 0 ? ScheduleZoomStep : -ScheduleZoomStep));
            e.Handled = true;
        }

        private void SetScheduleZoom(double value)
        {
            _scheduleZoom = Math.Max(MinScheduleZoom, Math.Min(MaxScheduleZoom, value));
            _scheduleZoom = Math.Round(_scheduleZoom / 0.05) * 0.05;
            ApplyScheduleZoom();
        }

        private void ApplyScheduleZoom()
        {
            if (ScheduleScale != null)
            {
                ScheduleScale.ScaleX = _scheduleZoom;
                ScheduleScale.ScaleY = _scheduleZoom;
            }

            if (ZoomText != null)
            {
                ZoomText.Text = Math.Round(_scheduleZoom * 100).ToString(CultureInfo.InvariantCulture) + "%";
            }
        }

        // ===================== Build =====================
        private void Rebuild()
        {
            ScheduleHost.Children.Clear();
            ScheduleHost.RowDefinitions.Clear();
            ScheduleHost.ColumnDefinitions.Clear();

            if (_vm == null || _vm.Current == null) return;
            var sched = _vm.Current;
            int days = sched.DayCount;

            // Columns: 0=姓名, 1..days=日期, days+1=上班, days+2=休息, days+3=连上, days+4=操作
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
            for (int d = 0; d < days; d++) ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 上班
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 休息
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 最长连上
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });  // 操作

            // Rows: 0=星期, 1=日期, 2=实休, 3=总休, 4..=员工
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            for (int i = 0; i < sched.Employees.Count; i++) ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(28) });

            var headerBg = new SolidColorBrush(Color.FromRgb(0xF1, 0xF3, 0xF6));
            var holidayBg = new SolidColorBrush(Color.FromRgb(0xFF, 0xF4, 0xE6));
            var quotaMismatchBg = new SolidColorBrush(Color.FromRgb(0xFF, 0xEB, 0xEE));

            // Row 0 — 星期
            ScheduleHost.Children.Add(MakeHeaderCell("姓名", 0, 0, headerBg, true));
            for (int d = 0; d < days; d++)
            {
                var date = sched.DateOf(d);
                var bg = HolidayService.IsHoliday(date) ? holidayBg : headerBg;
                ScheduleHost.Children.Add(MakeHeaderCell(DowZh[(int)date.DayOfWeek], 0, 1 + d, bg, false));
            }
            ScheduleHost.Children.Add(MakeHeaderCell("上班", 0, 1 + days, headerBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("休息", 0, 2 + days, headerBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("连上", 0, 3 + days, headerBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("", 0, 4 + days, headerBg, false));

            // Row 1 — 日期
            ScheduleHost.Children.Add(MakeHeaderCell("日期", 1, 0, headerBg, true));
            for (int d = 0; d < days; d++)
            {
                var date = sched.DateOf(d);
                var bg = HolidayService.IsHoliday(date) ? holidayBg : headerBg;
                ScheduleHost.Children.Add(MakeHeaderCell((d + 1).ToString(), 1, 1 + d, bg, false));
            }
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 1 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 2 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 3 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 4 + days, headerBg, false));

            // Row 2 — 实休（实际休息人数，只读；不在 [总休-0.5, 总休] 时标红）
            ScheduleHost.Children.Add(MakeHeaderCell("实休", 2, 0, headerBg, true));
            for (int d = 0; d < days; d++)
            {
                double quota = (d < sched.DailyRestQuotas.Count) ? sched.DailyRestQuotas[d] : 0;
                double actual = _vm.ComputeColumnRestCount(d);
                var matched = IsQuotaMatched(actual, quota);
                var actualCell = MakeHeaderCell(FormatStat(actual), 2, 1 + d, matched ? headerBg : quotaMismatchBg, false);
                if (actualCell is Border ab && ab.Child is TextBlock atb)
                {
                    atb.Foreground = matched ? Brushes.DarkSlateGray : Brushes.Crimson;
                    atb.FontWeight = FontWeights.SemiBold;
                    atb.ToolTip = BuildQuotaTooltip(actual, quota, matched);
                }
                ScheduleHost.Children.Add(actualCell);
            }
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 1 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 2 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 3 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 4 + days, headerBg, false));

            // Row 3 — 当日总休目标（可编辑；不在 [总休-0.5, 总休] 时标红）
            ScheduleHost.Children.Add(MakeHeaderCell("总休", 3, 0, headerBg, true));
            for (int d = 0; d < days; d++)
            {
                int dayIdx = d;
                double quota = (d < sched.DailyRestQuotas.Count) ? sched.DailyRestQuotas[d] : 0;
                double actual = _vm.ComputeColumnRestCount(d);
                var matched = IsQuotaMatched(actual, quota);
                var quotaTextBox = new TextBox
                {
                    Text = FormatStat(quota),
                    BorderThickness = new Thickness(0),
                    Background = Brushes.Transparent,
                    Foreground = matched ? Brushes.DarkGreen : Brushes.Crimson,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    IsReadOnly = !_vm.IsEditing,
                    Tag = d,
                    ToolTip = BuildQuotaTooltip(actual, quota, matched)
                };
                quotaTextBox.LostFocus += (s, e) =>
                {
                    if (TryParseQuota(quotaTextBox.Text, out var q))
                    {
                        _vm.UpdateDailyQuota(dayIdx, q);
                    }
                    else
                    {
                        quotaTextBox.Text = FormatStat(sched.DailyRestQuotas[dayIdx]);
                    }
                };
                ScheduleHost.Children.Add(WrapBorder(quotaTextBox, 3, 1 + d, matched ? headerBg : quotaMismatchBg));
            }
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 1 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 2 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 3 + days, headerBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 4 + days, headerBg, false));

            // Rows 4+ — 员工
            for (int e = 0; e < sched.Employees.Count; e++)
            {
                int row = 4 + e;
                int empIdx = e;

                // Name (editable in edit mode) prefixed with 1-based serial number.
                var nameBox = new TextBox
                {
                    Text = sched.Employees[e].Name,
                    BorderThickness = new Thickness(0),
                    Background = Brushes.Transparent,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    IsReadOnly = !_vm.IsEditing,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Padding = new Thickness(0)
                };
                nameBox.LostFocus += (s, ev) => _vm.UpdateEmployeeName(empIdx, nameBox.Text);

                var indexLabel = new TextBlock
                {
                    Text = (e + 1).ToString(CultureInfo.InvariantCulture),
                    FontSize = 9,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x90, 0x90, 0x90)),
                    Width = 18,
                    TextAlignment = TextAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false
                };

                var nameCellPanel = new DockPanel { LastChildFill = true };
                DockPanel.SetDock(indexLabel, Dock.Left);
                nameCellPanel.Children.Add(indexLabel);
                nameCellPanel.Children.Add(nameBox);
                ScheduleHost.Children.Add(WrapBorder(nameCellPanel, row, 0, headerBg));

                // Cells
                for (int d = 0; d < days; d++)
                {
                    int dayIdx = d;
                    var cell = sched.Employees[e].Cells[d];
                    var isFocusedCell = dayIdx == _focusedDayIndex &&
                                        (_focusedEmployeeIndex < 0 || empIdx == _focusedEmployeeIndex);
                    var btn = new Button
                    {
                        Content = string.IsNullOrEmpty(cell.Code) ? "·" : cell.Code,
                        FontSize = 12,
                        Padding = new Thickness(0),
                        BorderThickness = new Thickness(0.5),
                        BorderBrush = isFocusedCell
                            ? Brushes.Crimson
                            : new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                        Background = ResolveCellBg(cell, sched.DateOf(d)),
                        Foreground = ResolveCellFg(cell),
                        Cursor = _vm.IsEditing ? Cursors.Hand : Cursors.Arrow,
                        IsEnabled = _vm.IsEditing,
                        Tag = (empIdx, dayIdx),
                        FontWeight = cell.IsManual ? FontWeights.Bold : FontWeights.Normal
                    };
                    if (isFocusedCell)
                    {
                        btn.BorderThickness = new Thickness(2);
                        btn.ToolTip = "冲突定位";
                    }
                    btn.Click += CellButton_Click;
                    Grid.SetRow(btn, row);
                    Grid.SetColumn(btn, 1 + d);
                    ScheduleHost.Children.Add(btn);
                }

                // Stats
                var stats = _vm.ComputeRowStats(empIdx);
                var workTb = MakeStatCell(stats.work.ToString(stats.work % 1 == 0 ? "0" : "0.#", CultureInfo.InvariantCulture), row, 1 + days);
                ScheduleHost.Children.Add(workTb);
                var restTb = MakeStatCell(stats.rest.ToString(stats.rest % 1 == 0 ? "0" : "0.#", CultureInfo.InvariantCulture), row, 2 + days);
                if (stats.rest < 8) ((TextBlock)((Border)restTb).Child).Foreground = Brushes.Crimson;
                ScheduleHost.Children.Add(restTb);
                var runTb = MakeStatCell(stats.maxRun.ToString(), row, 3 + days);
                if (stats.maxRun > 5) ((TextBlock)((Border)runTb).Child).Foreground = Brushes.Crimson;
                ScheduleHost.Children.Add(runTb);

                // Delete employee
                var delBtn = new Button
                {
                    Content = "×",
                    Width = 22,
                    Height = 22,
                    Padding = new Thickness(0),
                    Margin = new Thickness(0),
                    Style = (Style)FindResource("MaterialDesignFlatButton"),
                    IsEnabled = _vm.IsEditing,
                    Tag = empIdx,
                    ToolTip = "删除该人员"
                };
                delBtn.Click += DeleteEmployeeButton_Click;
                Grid.SetRow(delBtn, row);
                Grid.SetColumn(delBtn, 4 + days);
                ScheduleHost.Children.Add(delBtn);
            }
        }

        private static Brush ResolveCellBg(ShiftCell cell, DateTime date)
        {
            if (cell == null || string.IsNullOrEmpty(cell.Code))
            {
                if (HolidayService.IsHoliday(date)) return new SolidColorBrush(Color.FromRgb(0xFF, 0xFA, 0xF0));
                return Brushes.White;
            }
            switch (cell.Code)
            {
                case "白": return new SolidColorBrush(Color.FromRgb(0xE3, 0xF2, 0xFD));
                case "卡": return new SolidColorBrush(Color.FromRgb(0xE8, 0xEA, 0xF6));
                case "副": return new SolidColorBrush(Color.FromRgb(0xE0, 0xF7, 0xFA));
                case "感": return new SolidColorBrush(Color.FromRgb(0xFF, 0xEB, 0xEE));
                case "大": return new SolidColorBrush(Color.FromRgb(0x37, 0x47, 0x4F));
                case "小": return new SolidColorBrush(Color.FromRgb(0x78, 0x90, 0x9C));
                case "休": return new SolidColorBrush(Color.FromRgb(0xC8, 0xE6, 0xC9));
                case "公": return new SolidColorBrush(Color.FromRgb(0xA5, 0xD6, 0xA7));
                case "午": return new SolidColorBrush(Color.FromRgb(0xFF, 0xF9, 0xC4));
                default: return Brushes.White;
            }
        }

        private static Brush ResolveCellFg(ShiftCell cell)
        {
            if (cell == null) return Brushes.Black;
            switch (cell.Code)
            {
                case "大":
                case "小":
                    return Brushes.White;
                default: return Brushes.Black;
            }
        }

        private static string FormatStat(double value)
        {
            return value % 1 == 0
                ? ((int)value).ToString(CultureInfo.InvariantCulture)
                : value.ToString("0.#", CultureInfo.InvariantCulture);
        }

        private static bool IsQuotaMatched(double actual, double quota)
        {
            return actual >= Math.Max(0, quota - 0.5) - 0.001 && actual <= quota + 0.001;
        }

        private static string BuildQuotaTooltip(double actual, double quota, bool matched)
        {
            var minAllowed = Math.Max(0, quota - 0.5);
            if (matched)
            {
                return actual < quota - 0.001
                    ? $"实际休息 {FormatStat(actual)} / 总休 {FormatStat(quota)}（允许范围 {FormatStat(minAllowed)}~{FormatStat(quota)}）"
                    : $"实际休息 {FormatStat(actual)} / 总休 {FormatStat(quota)}（允许范围 {FormatStat(minAllowed)}~{FormatStat(quota)}）";
            }
            return actual > quota + 0.001
                ? $"实际休息 {FormatStat(actual)} / 总休 {FormatStat(quota)}，多排 {FormatStat(actual - quota)}（超员）"
                : $"实际休息 {FormatStat(actual)} / 总休 {FormatStat(quota)}，少排 {FormatStat(minAllowed - actual)}（低于允许范围 {FormatStat(minAllowed)}~{FormatStat(quota)}）";
        }

        private static bool TryParseQuota(string text, out double quota)
        {
            quota = 0;
            text = (text ?? string.Empty).Trim();
            if (text.Length == 0)
            {
                return false;
            }

            if (!double.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.CurrentCulture, out quota) &&
                !double.TryParse(text.Replace(',', '.'), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out quota))
            {
                return false;
            }

            if (quota < 0)
            {
                return false;
            }

            return Math.Abs(quota * 2 - Math.Round(quota * 2)) < 0.001;
        }

        private static Border MakeHeaderCell(string text, int row, int col, Brush bg, bool bold)
        {
            var tb = new TextBlock
            {
                Text = text,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 12,
                FontWeight = bold ? FontWeights.SemiBold : FontWeights.Normal
            };
            var b = new Border
            {
                Background = bg,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                BorderThickness = new Thickness(0.5),
                Child = tb
            };
            Grid.SetRow(b, row); Grid.SetColumn(b, col);
            return b;
        }

        private static Border MakeStatCell(string text, int row, int col)
        {
            var tb = new TextBlock
            {
                Text = text,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold
            };
            var b = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                BorderThickness = new Thickness(0.5),
                Child = tb
            };
            Grid.SetRow(b, row); Grid.SetColumn(b, col);
            return b;
        }

        private static Border WrapBorder(UIElement child, int row, int col, Brush bg = null)
        {
            var b = new Border
            {
                Background = bg,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                BorderThickness = new Thickness(0.5),
                Child = child
            };
            Grid.SetRow(b, row); Grid.SetColumn(b, col);
            return b;
        }

        // ===================== Cell click → popup picker =====================
        private void CellButton_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null || !_vm.IsEditing) return;
            var btn = sender as Button;
            if (btn?.Tag is ValueTuple<int, int> tag)
            {
                ShowShiftPicker(btn, tag.Item1, tag.Item2);
            }
        }

        private void ShowShiftPicker(Button anchor, int empIdx, int dayIdx)
        {
            var popup = new Popup
            {
                Placement = PlacementMode.Bottom,
                PlacementTarget = anchor,
                StaysOpen = false,
                AllowsTransparency = true,
                PopupAnimation = PopupAnimation.Fade
            };

            var border = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xBD, 0xBD, 0xBD)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10),
                Effect = new System.Windows.Media.Effects.DropShadowEffect { ShadowDepth = 2, Opacity = 0.25, BlurRadius = 8 }
            };

            var panel = new WrapPanel { MaxWidth = 260 };
            void Add(string label, Action onClick, Brush bg = null, Brush fg = null)
            {
                var b = new Button
                {
                    Content = label,
                    Width = 46,
                    Height = 32,
                    Margin = new Thickness(2),
                    Padding = new Thickness(0),
                    FontSize = 12,
                    Background = bg,
                    Foreground = fg ?? Brushes.Black,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center
                };
                b.Click += (s, e) => { onClick(); popup.IsOpen = false; };
                panel.Children.Add(b);
            }

            Add("白", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Day), new SolidColorBrush(Color.FromRgb(0xE3, 0xF2, 0xFD)));
            Add("卡", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Card), new SolidColorBrush(Color.FromRgb(0xE8, 0xEA, 0xF6)));
            Add("副", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Deputy), new SolidColorBrush(Color.FromRgb(0xE0, 0xF7, 0xFA)));
            Add("感", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Infect), new SolidColorBrush(Color.FromRgb(0xFF, 0xEB, 0xEE)));
            Add("大", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Big), new SolidColorBrush(Color.FromRgb(0x37, 0x47, 0x4F)), Brushes.White);
            Add("小", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Small), new SolidColorBrush(Color.FromRgb(0x78, 0x90, 0x9C)), Brushes.White);
            Add("休", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Rest), new SolidColorBrush(Color.FromRgb(0xC8, 0xE6, 0xC9)));
            Add("公", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Public), new SolidColorBrush(Color.FromRgb(0xA5, 0xD6, 0xA7)));
            Add("午", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Half), new SolidColorBrush(Color.FromRgb(0xFF, 0xF9, 0xC4)));

            // separator
            panel.Children.Add(new Border { Width = 252, Height = 1, Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)), Margin = new Thickness(0, 4, 0, 4) });

            Add("大1", () => _vm.ApplyBigNight1(empIdx, dayIdx), new SolidColorBrush(Color.FromRgb(0xFF, 0xCC, 0x80)));
            Add("大2", () => _vm.ApplyBigNight2(empIdx, dayIdx), new SolidColorBrush(Color.FromRgb(0xFF, 0xB7, 0x4D)));
            Add("清空", () => _vm.SetCell(empIdx, dayIdx, ""), Brushes.WhiteSmoke);

            border.Child = panel;
            popup.Child = border;
            popup.IsOpen = true;
        }

        private void DeleteEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null || !_vm.IsEditing) return;
            if (sender is Button b && b.Tag is int idx)
            {
                if (MessageBox.Show($"确认删除该人员？", "删除", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                {
                    _vm.RemoveEmployee(idx);
                }
            }
        }
    }

    /// <summary>极简文本输入对话框。</summary>
    internal static class SimplePromptDialog
    {
        public static string Prompt(string message, string title, string defaultText)
        {
            var dlg = new Window
            {
                Title = title,
                Width = 360,
                Height = 170,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Owner = Application.Current?.MainWindow,
                Background = SystemColors.WindowBrush
            };
            var stack = new StackPanel { Margin = new Thickness(16) };
            stack.Children.Add(new TextBlock { Text = message, Margin = new Thickness(0, 0, 0, 8) });
            var tb = new TextBox { Text = defaultText ?? "" };
            tb.Focus();
            tb.SelectAll();
            stack.Children.Add(tb);
            var bp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var ok = new Button { Content = "确定", Width = 70, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
            var cancel = new Button { Content = "取消", Width = 70, IsCancel = true };
            ok.Click += (s, e) => { dlg.DialogResult = true; };
            bp.Children.Add(ok);
            bp.Children.Add(cancel);
            stack.Children.Add(bp);
            dlg.Content = stack;
            tb.Loaded += (s, e) => tb.Focus();
            return dlg.ShowDialog() == true ? tb.Text : null;
        }
    }
}
