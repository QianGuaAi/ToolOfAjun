using System;
using System.Collections.Generic;
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
        private readonly DispatcherTimer _refreshTimer;
        private int _focusedEmployeeIndex = -1;
        private int _focusedDayIndex = -1;
        private readonly HashSet<ValueTuple<int, int>> _selectedCells = new HashSet<ValueTuple<int, int>>();
        private ValueTuple<int, int>? _selectionAnchor;
        private ValueTuple<int, int>? _fillSourceCell;
        private bool _isSelectingCells;
        private bool _dragSelectionMoved;
        private bool _suppressNextCellClick;
        private double _scheduleZoom = 1.0;
        private int _renderedDayCount;
        private int _renderedEmployeeCount;
        private bool[] _holidayByDay;
        private TextBlock[] _actualRestTexts;
        private Border[] _actualRestBorders;
        private TextBox[] _quotaTextBoxes;
        private Border[] _quotaBorders;
        private TextBox[] _nameTextBoxes;
        private Button[,] _cellButtons;
        private TextBlock[] _workStatTexts;
        private TextBlock[] _restStatTexts;
        private TextBlock[] _runStatTexts;
        private TextBlock[] _weekendRestStatTexts;
        private Button[] _deleteButtons;
        private readonly List<UIElement> _employeeRowElements = new List<UIElement>();
        private int _firstRenderedEmployeeIndex = -1;
        private int _lastRenderedEmployeeIndex = -1;
        private const double MinScheduleZoom = 0.65;
        private const double MaxScheduleZoom = 1.8;
        private const double ScheduleZoomStep = 0.1;
        private const int HeaderRowCount = 4;
        private const double HeaderRowHeight = 24;
        private const double EmployeeRowHeight = 28;
        private const int EmployeeRowOverscan = 12; // Increased for smoother scrolling on large rosters (perf optimization)
        private static readonly string[] DowZh = { "日", "一", "二", "三", "四", "五", "六" };
        private static readonly Brush HeaderBg = FrozenBrush(0xF1, 0xF5, 0xF9);
        private static readonly Brush HolidayColumnBg = FrozenBrush(0xBF, 0xDB, 0xFE);
        private static readonly Brush QuotaMismatchBg = FrozenBrush(0xFE, 0xE2, 0xE2);
        private static readonly Brush BorderBg = FrozenBrush(0xE2, 0xE8, 0xF0);
        private static readonly Brush SelectionBorderBg = FrozenBrush(0x25, 0x63, 0xEB);
        private static readonly Brush FillSourceBorderBg = FrozenBrush(0x16, 0xA3, 0x4A);
        private static readonly Brush CellDayBg = FrozenBrush(0xFF, 0xFF, 0xFF);
        private static readonly Brush CellCardBg = FrozenBrush(0xA7, 0xF3, 0xD0);
        private static readonly Brush CellDeputyBg = FrozenBrush(0x99, 0x1B, 0x1B);
        private static readonly Brush CellInfectBg = FrozenBrush(0xFE, 0xE2, 0xE2);
        private static readonly Brush CellBigBg = FrozenBrush(0x1E, 0x29, 0x3B);
        private static readonly Brush CellSmallBg = FrozenBrush(0x47, 0x55, 0x69);
        private static readonly Brush CellRestBg = FrozenBrush(0xFD, 0xEC, 0xC8);
        private static readonly Brush CellPublicBg = FrozenBrush(0xFE, 0xF3, 0xC7);
        private static readonly Brush CellMaternityBg = FrozenBrush(0xFE, 0xD7, 0xA8);
        private static readonly Brush CellHalfBg = FrozenBrush(0xFE, 0xF3, 0xC7);
        private static readonly Brush SerialTextFg = FrozenBrush(0x64, 0x74, 0x8B);
        private static readonly string[] ShiftPickerVisibleLabels = { "白", "卡", "副", "感", "大夜", "小", "休", "公", "产假", "小1", "清空" };

        public SchedulePage()
        {
            InitializeComponent();
            Focusable = true;
            ApplyScheduleZoom();
            _rebuildTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(45) };
            _rebuildTimer.Tick += (s, e) =>
            {
                _rebuildTimer.Stop();
                Rebuild();
            };
            _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(25) };
            _refreshTimer.Tick += (s, e) =>
            {
                _refreshTimer.Stop();
                RefreshDataOnly();
            };
            PreviewKeyDown += SchedulePage_PreviewKeyDown;
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
                ClearScheduleSelection();
                Rebuild();
            }
        }

        private void Vm_ScheduleStructureChanged(object sender, EventArgs e)
        {
            ClearScheduleSelection();
            QueueRebuild();
        }
        private void Vm_ScheduleDataChanged(object sender, EventArgs e) => QueueDataRefresh();
        private void Vm_ScheduleCellFocusRequested(object sender, ScheduleCellFocusRequestedEventArgs e)
        {
            _focusedEmployeeIndex = e.EmployeeIndex;
            _focusedDayIndex = e.DayIndex;
            ScrollFocusedCellIntoView();
            QueueDataRefresh();
        }

        private void ScrollFocusedCellIntoView()
        {
            if (ScheduleScroll == null)
            {
                return;
            }

            var scale = _scheduleZoom <= 0 ? 1.0 : _scheduleZoom;
            if (_focusedEmployeeIndex >= 0)
            {
                var top = (HeaderRowCount * HeaderRowHeight + _focusedEmployeeIndex * EmployeeRowHeight) * scale;
                ScheduleScroll.ScrollToVerticalOffset(Math.Max(0, top - EmployeeRowHeight * scale));
            }

            if (_focusedDayIndex >= 0)
            {
                var left = (72 + _focusedDayIndex * 30) * scale;
                ScheduleScroll.ScrollToHorizontalOffset(Math.Max(0, left - 90 * scale));
            }
        }

        private void QueueRebuild()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(QueueRebuild));
                return;
            }

            _rebuildTimer.Stop();
            _refreshTimer.Stop();
            _rebuildTimer.Start();
        }

        private void QueueDataRefresh()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(QueueDataRefresh));
                return;
            }

            if (_rebuildTimer.IsEnabled)
            {
                return;
            }

            _refreshTimer.Stop();
            _refreshTimer.Start();
        }

        private static Brush FrozenBrush(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
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

        private void FillSelectedCells_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null || _selectedCells.Count == 0)
            {
                MessageBox.Show("请先在排班表中选择要填充的单元格。", "填充", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var source = _fillSourceCell ?? _selectionAnchor ?? FirstSelectedCell();
            if (source.Item1 < 0 || source.Item2 < 0)
            {
                MessageBox.Show("请先选择一个源单元格。", "填充", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _vm.FillCellsFromSource(source.Item1, source.Item2, _selectedCells);
        }

        private void ScheduleScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
            {
                return;
            }

            SetScheduleZoom(_scheduleZoom + (e.Delta > 0 ? ScheduleZoomStep : -ScheduleZoomStep));
            e.Handled = true;
        }

        private void ScheduleScroll_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (Math.Abs(e.VerticalChange) > 0.001 || Math.Abs(e.ViewportHeightChange) > 0.001)
            {
                RenderVisibleEmployeeRows(false);
            }
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

            RenderVisibleEmployeeRows(false);

            if (ZoomText != null)
            {
                ZoomText.Text = Math.Round(_scheduleZoom * 100).ToString(CultureInfo.InvariantCulture) + "%";
            }
        }

        private void SchedulePage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Delete)
            {
                return;
            }

            if (FindAncestor<TextBox>(e.OriginalSource as DependencyObject) != null)
            {
                return;
            }

            if (_selectedCells.Count == 0 || _vm == null)
            {
                return;
            }

            var cleared = _vm.ClearCells(_selectedCells);
            if (cleared > 0)
            {
                e.Handled = true;
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
            int employeeCount = sched.Employees.Count;

            _renderedDayCount = days;
            _renderedEmployeeCount = employeeCount;
            _holidayByDay = new bool[days];
            _actualRestTexts = new TextBlock[days];
            _actualRestBorders = new Border[days];
            _quotaTextBoxes = new TextBox[days];
            _quotaBorders = new Border[days];
            _nameTextBoxes = new TextBox[employeeCount];
            _cellButtons = new Button[employeeCount, days];
            _workStatTexts = new TextBlock[employeeCount];
            _restStatTexts = new TextBlock[employeeCount];
            _runStatTexts = new TextBlock[employeeCount];
            _weekendRestStatTexts = new TextBlock[employeeCount];
            _deleteButtons = new Button[employeeCount];

            // Columns: 0=姓名, 1..days=日期, days+1=上班, days+2=休息, days+3=连上, days+4=末休, days+5=操作
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
            for (int d = 0; d < days; d++) ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 上班
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 休息
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 最长连上
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });  // 周末休息
            ScheduleHost.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });  // 操作
            ScheduleHost.MinWidth = 72 + days * 30 + 42 * 4 + 34;

            // Rows: 0=星期, 1=日期, 2=实休, 3=总休, 4..=员工
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
            for (int i = 0; i < employeeCount; i++) ScheduleHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(28) });
            ScheduleHost.MinHeight = HeaderRowCount * HeaderRowHeight + employeeCount * EmployeeRowHeight;
            for (int d = 0; d < days; d++) _holidayByDay[d] = HolidayService.IsHoliday(sched.DateOf(d));

            // Row 0 — 星期
            ScheduleHost.Children.Add(MakeHeaderCell("姓名", 0, 0, HeaderBg, true));
            for (int d = 0; d < days; d++)
            {
                var date = sched.DateOf(d);
                var bg = _holidayByDay[d] ? HolidayColumnBg : HeaderBg;
                ScheduleHost.Children.Add(MakeHeaderCell(DowZh[(int)date.DayOfWeek], 0, 1 + d, bg, false));
            }
            ScheduleHost.Children.Add(MakeHeaderCell("上班", 0, 1 + days, HeaderBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("休息", 0, 2 + days, HeaderBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("连上", 0, 3 + days, HeaderBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("末休", 0, 4 + days, HeaderBg, true));
            ScheduleHost.Children.Add(MakeHeaderCell("", 0, 5 + days, HeaderBg, false));

            // Row 1 — 日期
            ScheduleHost.Children.Add(MakeHeaderCell("日期", 1, 0, HeaderBg, true));
            for (int d = 0; d < days; d++)
            {
                var bg = _holidayByDay[d] ? HolidayColumnBg : HeaderBg;
                ScheduleHost.Children.Add(MakeHeaderCell((d + 1).ToString(), 1, 1 + d, bg, false));
            }
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 1 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 2 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 3 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 4 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 1, 5 + days, HeaderBg, false));

            // Row 2 — 实休（实际休息人数，只读；不在 [总休-0.5, 总休] 时标红）
            ScheduleHost.Children.Add(MakeHeaderCell("实休", 2, 0, HeaderBg, true));
            for (int d = 0; d < days; d++)
            {
                double quota = (d < sched.DailyRestQuotas.Count) ? sched.DailyRestQuotas[d] : 0;
                double actual = _vm.ComputeColumnRestCount(d);
                var matched = IsQuotaMatched(actual, quota);
                var actualCell = MakeHeaderCell(FormatStat(actual), 2, 1 + d, GetColumnStatBg(d, matched), false);
                if (actualCell is Border ab && ab.Child is TextBlock atb)
                {
                    atb.Foreground = matched ? Brushes.DarkSlateGray : Brushes.Crimson;
                    atb.FontWeight = FontWeights.SemiBold;
                    atb.ToolTip = BuildQuotaTooltip(actual, quota, matched);
                    _actualRestTexts[d] = atb;
                    _actualRestBorders[d] = ab;
                }
                ScheduleHost.Children.Add(actualCell);
            }
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 1 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 2 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 3 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 4 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 2, 5 + days, HeaderBg, false));

            // Row 3 — 当日总休目标（可编辑；不在 [总休-0.5, 总休] 时标红）
            ScheduleHost.Children.Add(MakeHeaderCell("总休", 3, 0, HeaderBg, true));
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
                    IsReadOnly = false,
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
                var quotaBorder = WrapBorder(quotaTextBox, 3, 1 + d, GetColumnStatBg(d, matched));
                _quotaTextBoxes[d] = quotaTextBox;
                _quotaBorders[d] = quotaBorder;
                ScheduleHost.Children.Add(quotaBorder);
            }
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 1 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 2 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 3 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 4 + days, HeaderBg, false));
            ScheduleHost.Children.Add(MakeHeaderCell("", 3, 5 + days, HeaderBg, false));

            RenderVisibleEmployeeRows(true);
        }

        private void ClearVirtualizedEmployeeRows()
        {
            if (_employeeRowElements.Count > 0)
            {
                foreach (var element in _employeeRowElements)
                {
                    ScheduleHost.Children.Remove(element);
                }
                _employeeRowElements.Clear();
            }

            if (_firstRenderedEmployeeIndex >= 0 && _lastRenderedEmployeeIndex >= _firstRenderedEmployeeIndex)
            {
                for (var employeeIndex = _firstRenderedEmployeeIndex; employeeIndex <= _lastRenderedEmployeeIndex; employeeIndex++)
                {
                    if (_nameTextBoxes != null && employeeIndex < _nameTextBoxes.Length)
                    {
                        _nameTextBoxes[employeeIndex] = null;
                        _workStatTexts[employeeIndex] = null;
                        _restStatTexts[employeeIndex] = null;
                        _runStatTexts[employeeIndex] = null;
                        _weekendRestStatTexts[employeeIndex] = null;
                        _deleteButtons[employeeIndex] = null;
                        if (_cellButtons != null)
                        {
                            for (var day = 0; day < _renderedDayCount; day++)
                            {
                                _cellButtons[employeeIndex, day] = null;
                            }
                        }
                    }
                }
            }

            _firstRenderedEmployeeIndex = -1;
            _lastRenderedEmployeeIndex = -1;
        }

        private void RenderVisibleEmployeeRows(bool force)
        {
            if (_vm == null || _vm.Current == null || ScheduleHost == null)
            {
                return;
            }

            var employeeCount = _vm.Current.Employees.Count;
            if (employeeCount == 0 || _renderedDayCount <= 0)
            {
                ClearVirtualizedEmployeeRows();
                return;
            }

            var viewportHeight = ScheduleScroll?.ViewportHeight ?? 0;
            if (viewportHeight <= 0 || double.IsNaN(viewportHeight) || double.IsInfinity(viewportHeight))
            {
                viewportHeight = 700;
            }

            var scale = _scheduleZoom <= 0 ? 1.0 : _scheduleZoom;
            var logicalTop = (ScheduleScroll?.VerticalOffset ?? 0) / scale;
            var logicalBottom = logicalTop + viewportHeight / scale;
            var employeeAreaTop = HeaderRowCount * HeaderRowHeight;
            var first = (int)Math.Floor((logicalTop - employeeAreaTop) / EmployeeRowHeight) - EmployeeRowOverscan;
            var last = (int)Math.Ceiling((logicalBottom - employeeAreaTop) / EmployeeRowHeight) + EmployeeRowOverscan;

            first = Math.Max(0, first);
            last = Math.Min(employeeCount - 1, Math.Max(first, last));

            if (!force && first == _firstRenderedEmployeeIndex && last == _lastRenderedEmployeeIndex)
            {
                return;
            }

            ClearVirtualizedEmployeeRows();
            _firstRenderedEmployeeIndex = first;
            _lastRenderedEmployeeIndex = last;
            for (var employeeIndex = first; employeeIndex <= last; employeeIndex++)
            {
                AddEmployeeRow(_vm.Current, employeeIndex);
            }

            RefreshDataOnly();
        }

        private void AddEmployeeRow(ScheduleVersion sched, int employeeIndex)
        {
            var days = sched.DayCount;
            var row = HeaderRowCount + employeeIndex;
            var empIdx = employeeIndex;

            var nameBox = new TextBox
            {
                Text = sched.Employees[employeeIndex].Name,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                IsReadOnly = false,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Padding = new Thickness(0)
            };
            nameBox.LostFocus += (s, ev) => _vm.UpdateEmployeeName(empIdx, nameBox.Text);
            _nameTextBoxes[employeeIndex] = nameBox;

            var indexLabel = new TextBlock
            {
                Text = (employeeIndex + 1).ToString(CultureInfo.InvariantCulture),
                FontSize = 9,
                Foreground = SerialTextFg,
                Width = 18,
                TextAlignment = TextAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false
            };

            var nameCellPanel = new DockPanel { LastChildFill = true };
            DockPanel.SetDock(indexLabel, Dock.Left);
            nameCellPanel.Children.Add(indexLabel);
            nameCellPanel.Children.Add(nameBox);
            AddEmployeeElement(WrapBorder(nameCellPanel, row, 0, HeaderBg));

            for (var day = 0; day < days; day++)
            {
                var cell = sched.Employees[employeeIndex].Cells[day];
                var isFocusedCell = day == _focusedDayIndex &&
                                    (_focusedEmployeeIndex < 0 || empIdx == _focusedEmployeeIndex);
                var btn = new Button
                {
                    Content = DisplayCellCode(cell.Code),
                    Style = (Style)FindResource("ScheduleCellButton"),
                    Background = ResolveCellBg(cell, _holidayByDay != null && day < _holidayByDay.Length && _holidayByDay[day]),
                    Foreground = ResolveCellFg(cell, _holidayByDay != null && day < _holidayByDay.Length && _holidayByDay[day]),
                    IsEnabled = true,
                    Tag = (empIdx, day),
                    FontWeight = cell.IsManual ? FontWeights.Bold : FontWeights.Normal,
                    ToolTip = isFocusedCell ? "冲突定位" : null
                };
                btn.PreviewMouseLeftButtonDown += CellButton_PreviewMouseLeftButtonDown;
                btn.PreviewMouseMove += CellButton_PreviewMouseMove;
                btn.PreviewMouseLeftButtonUp += CellButton_PreviewMouseLeftButtonUp;
                btn.Click += CellButton_Click;
                Grid.SetRow(btn, row);
                Grid.SetColumn(btn, 1 + day);
                _cellButtons[employeeIndex, day] = btn;
                AddEmployeeElement(btn);
            }

            var stats = _vm.ComputeRowStats(empIdx);
            var workTb = MakeStatCell(stats.work.ToString(stats.work % 1 == 0 ? "0" : "0.#", CultureInfo.InvariantCulture), row, 1 + days);
            _workStatTexts[employeeIndex] = GetBorderText(workTb);
            AddEmployeeElement(workTb);
            var restTb = MakeStatCell(stats.rest.ToString(stats.rest % 1 == 0 ? "0" : "0.#", CultureInfo.InvariantCulture), row, 2 + days);
            _restStatTexts[employeeIndex] = GetBorderText(restTb);
            if (stats.rest < 9) _restStatTexts[employeeIndex].Foreground = Brushes.Crimson;
            AddEmployeeElement(restTb);
            var runTb = MakeStatCell(stats.maxRun.ToString(), row, 3 + days);
            _runStatTexts[employeeIndex] = GetBorderText(runTb);
            if (stats.maxRun > 5) _runStatTexts[employeeIndex].Foreground = Brushes.Crimson;
            AddEmployeeElement(runTb);
            var weekendRestTb = MakeStatCell(FormatStat(stats.weekendRest), row, 4 + days);
            _weekendRestStatTexts[employeeIndex] = GetBorderText(weekendRestTb);
            AddEmployeeElement(weekendRestTb);

            var delBtn = new Button
            {
                Content = "×",
                Width = 22,
                Height = 22,
                Padding = new Thickness(0),
                Margin = new Thickness(0),
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                IsEnabled = true,
                Tag = empIdx,
                ToolTip = "删除该人员"
            };
            delBtn.Click += DeleteEmployeeButton_Click;
            Grid.SetRow(delBtn, row);
            Grid.SetColumn(delBtn, 5 + days);
            _deleteButtons[employeeIndex] = delBtn;
            AddEmployeeElement(delBtn);
        }

        private void AddEmployeeElement(UIElement element)
        {
            ScheduleHost.Children.Add(element);
            _employeeRowElements.Add(element);
        }

        private void RefreshDataOnly()
        {
            if (_vm == null || _vm.Current == null)
            {
                return;
            }

            var sched = _vm.Current;
            var days = sched.DayCount;
            var employeeCount = sched.Employees.Count;
            if (days != _renderedDayCount ||
                employeeCount != _renderedEmployeeCount ||
                _cellButtons == null ||
                _cellButtons.GetLength(0) != employeeCount ||
                _cellButtons.GetLength(1) != days)
            {
                QueueRebuild();
                return;
            }

            for (var day = 0; day < days; day++)
            {
                var quota = day < sched.DailyRestQuotas.Count ? sched.DailyRestQuotas[day] : 0;
                var actual = _vm.ComputeColumnRestCount(day);
                var matched = IsQuotaMatched(actual, quota);
                var bg = GetColumnStatBg(day, matched);
                var tooltip = BuildQuotaTooltip(actual, quota, matched);

                if (_actualRestTexts != null && _actualRestTexts[day] != null)
                {
                    _actualRestTexts[day].Text = FormatStat(actual);
                    _actualRestTexts[day].Foreground = matched ? Brushes.DarkSlateGray : Brushes.Crimson;
                    _actualRestTexts[day].ToolTip = tooltip;
                }

                if (_actualRestBorders != null && _actualRestBorders[day] != null)
                {
                    _actualRestBorders[day].Background = bg;
                }

                if (_quotaTextBoxes != null && _quotaTextBoxes[day] != null)
                {
                    if (!_quotaTextBoxes[day].IsKeyboardFocusWithin)
                    {
                        _quotaTextBoxes[day].Text = FormatStat(quota);
                    }

                    _quotaTextBoxes[day].Foreground = matched ? Brushes.DarkGreen : Brushes.Crimson;
                    _quotaTextBoxes[day].ToolTip = tooltip;
                    _quotaTextBoxes[day].IsReadOnly = false;
                }

                if (_quotaBorders != null && _quotaBorders[day] != null)
                {
                    _quotaBorders[day].Background = bg;
                }
            }

            var firstEmployee = _firstRenderedEmployeeIndex < 0 ? 0 : _firstRenderedEmployeeIndex;
            var lastEmployee = _lastRenderedEmployeeIndex < firstEmployee
                ? Math.Min(employeeCount - 1, firstEmployee)
                : Math.Min(employeeCount - 1, _lastRenderedEmployeeIndex);

            for (var empIdx = firstEmployee; empIdx <= lastEmployee; empIdx++)
            {
                var employee = sched.Employees[empIdx];
                if (_nameTextBoxes != null && _nameTextBoxes[empIdx] != null && !_nameTextBoxes[empIdx].IsKeyboardFocusWithin)
                {
                    _nameTextBoxes[empIdx].Text = employee.Name;
                    _nameTextBoxes[empIdx].IsReadOnly = false;
                }

                for (var day = 0; day < days; day++)
                {
                    var button = _cellButtons[empIdx, day];
                    if (button == null || employee.Cells == null || day >= employee.Cells.Count)
                    {
                        continue;
                    }

                    var cell = employee.Cells[day];
                    var isFocusedCell = day == _focusedDayIndex &&
                                        (_focusedEmployeeIndex < 0 || empIdx == _focusedEmployeeIndex);
                    button.Content = DisplayCellCode(cell.Code);
                    button.Background = ResolveCellBg(cell, _holidayByDay != null && day < _holidayByDay.Length && _holidayByDay[day]);
                    button.Foreground = ResolveCellFg(cell, _holidayByDay != null && day < _holidayByDay.Length && _holidayByDay[day]);
                    button.Cursor = Cursors.Hand;
                    button.IsEnabled = true;
                    button.FontWeight = cell.IsManual ? FontWeights.Bold : FontWeights.Normal;
                    ApplyCellButtonState(button, empIdx, day, isFocusedCell);
                }

                var stats = _vm.ComputeRowStats(empIdx);
                SetStatText(_workStatTexts, empIdx, stats.work);
                SetStatText(_restStatTexts, empIdx, stats.rest, stats.rest < 9);
                SetStatText(_weekendRestStatTexts, empIdx, stats.weekendRest);
                if (_runStatTexts != null && _runStatTexts[empIdx] != null)
                {
                    _runStatTexts[empIdx].Text = stats.maxRun.ToString(CultureInfo.InvariantCulture);
                    _runStatTexts[empIdx].Foreground = stats.maxRun > 5 ? Brushes.Crimson : Brushes.Black;
                }

                if (_deleteButtons != null && _deleteButtons[empIdx] != null)
                {
                    _deleteButtons[empIdx].IsEnabled = true;
                }
            }
        }

        private static void SetStatText(TextBlock[] textBlocks, int index, double value, bool alert = false)
        {
            if (textBlocks == null || index < 0 || index >= textBlocks.Length || textBlocks[index] == null)
            {
                return;
            }

            textBlocks[index].Text = value.ToString(value % 1 == 0 ? "0" : "0.#", CultureInfo.InvariantCulture);
            textBlocks[index].Foreground = alert ? Brushes.Crimson : Brushes.Black;
        }

        private void ApplyCellButtonState(Button button, int empIdx, int dayIdx, bool isFocusedCell)
        {
            if (button == null)
            {
                return;
            }

            var key = ValueTuple.Create(empIdx, dayIdx);
            var isSelected = _selectedCells.Contains(key);
            var isFillSource = _fillSourceCell.HasValue && _fillSourceCell.Value.Equals(key);

            if (isFillSource)
            {
                button.BorderBrush = FillSourceBorderBg;
                button.BorderThickness = new Thickness(2.5);
            }
            else if (isSelected)
            {
                button.BorderBrush = SelectionBorderBg;
                button.BorderThickness = new Thickness(2);
            }
            else if (isFocusedCell)
            {
                button.BorderBrush = Brushes.Crimson;
                button.BorderThickness = new Thickness(2);
            }
            else
            {
                button.BorderBrush = BorderBg;
                button.BorderThickness = new Thickness(0.5);
            }

            if (isFillSource)
            {
                button.ToolTip = "填充源单元格";
            }
            else if (isSelected)
            {
                button.ToolTip = _selectedCells.Count > 1 ? $"已选择 {_selectedCells.Count} 个单元格" : "已选择单元格";
            }
            else
            {
                button.ToolTip = isFocusedCell ? "冲突定位" : null;
            }
        }

        private void ClearScheduleSelection()
        {
            if (_selectedCells.Count == 0 && !_fillSourceCell.HasValue && !_selectionAnchor.HasValue)
            {
                return;
            }

            _selectedCells.Clear();
            _selectionAnchor = null;
            _fillSourceCell = null;
            RefreshSelectionBorders();
        }

        private void SetSingleSelectedCell(int empIdx, int dayIdx)
        {
            var key = ValueTuple.Create(empIdx, dayIdx);
            _selectedCells.Clear();
            _selectedCells.Add(key);
            _selectionAnchor = key;
            _fillSourceCell = key;
            RefreshSelectionBorders();
        }

        private void ToggleSelectedCell(int empIdx, int dayIdx)
        {
            var key = ValueTuple.Create(empIdx, dayIdx);
            if (_selectedCells.Contains(key))
            {
                _selectedCells.Remove(key);
                if (_fillSourceCell.HasValue && _fillSourceCell.Value.Equals(key))
                {
                    _fillSourceCell = _selectedCells.Count > 0 ? FirstSelectedCell() : (ValueTuple<int, int>?)null;
                }
            }
            else
            {
                _selectedCells.Add(key);
                _selectionAnchor = key;
                _fillSourceCell = key;
            }

            if (_selectedCells.Count == 0)
            {
                _selectionAnchor = null;
                _fillSourceCell = null;
            }

            RefreshSelectionBorders();
        }

        private void SelectCellRange(ValueTuple<int, int> start, ValueTuple<int, int> end, bool append)
        {
            if (_vm?.Current == null)
            {
                return;
            }

            if (!append)
            {
                _selectedCells.Clear();
            }

            var minEmployee = Math.Max(0, Math.Min(start.Item1, end.Item1));
            var maxEmployee = Math.Min(_vm.Current.Employees.Count - 1, Math.Max(start.Item1, end.Item1));
            var minDay = Math.Max(0, Math.Min(start.Item2, end.Item2));
            var maxDay = Math.Min(_vm.Current.DayCount - 1, Math.Max(start.Item2, end.Item2));

            for (var employee = minEmployee; employee <= maxEmployee; employee++)
            {
                for (var day = minDay; day <= maxDay; day++)
                {
                    _selectedCells.Add(ValueTuple.Create(employee, day));
                }
            }

            _selectionAnchor = start;
            _fillSourceCell = start;
            RefreshSelectionBorders();
        }

        private void RefreshSelectionBorders()
        {
            if (_cellButtons == null || _vm?.Current == null)
            {
                return;
            }

            if (_cellButtons.GetLength(0) == 0 || _cellButtons.GetLength(1) == 0)
            {
                return;
            }

            var firstEmployee = _firstRenderedEmployeeIndex < 0 ? 0 : _firstRenderedEmployeeIndex;
            var lastEmployee = _lastRenderedEmployeeIndex < firstEmployee
                ? Math.Min(_vm.Current.Employees.Count - 1, firstEmployee)
                : Math.Min(_vm.Current.Employees.Count - 1, _lastRenderedEmployeeIndex);
            firstEmployee = Math.Min(firstEmployee, _cellButtons.GetLength(0) - 1);
            lastEmployee = Math.Min(lastEmployee, _cellButtons.GetLength(0) - 1);
            if (firstEmployee < 0 || lastEmployee < firstEmployee)
            {
                return;
            }

            var maxDay = Math.Min(_vm.Current.DayCount, _cellButtons.GetLength(1));

            for (var empIdx = firstEmployee; empIdx <= lastEmployee; empIdx++)
            {
                for (var day = 0; day < maxDay; day++)
                {
                    var button = _cellButtons[empIdx, day];
                    if (button == null)
                    {
                        continue;
                    }

                    var isFocusedCell = day == _focusedDayIndex &&
                                        (_focusedEmployeeIndex < 0 || empIdx == _focusedEmployeeIndex);
                    ApplyCellButtonState(button, empIdx, day, isFocusedCell);
                }
            }
        }

        private ValueTuple<int, int> FirstSelectedCell()
        {
            foreach (var cell in _selectedCells)
            {
                return cell;
            }

            return ValueTuple.Create(-1, -1);
        }

        private bool TryGetCellTag(object sender, out int empIdx, out int dayIdx)
        {
            empIdx = -1;
            dayIdx = -1;
            if (sender is Button button && button.Tag is ValueTuple<int, int> tag)
            {
                empIdx = tag.Item1;
                dayIdx = tag.Item2;
                return true;
            }

            return false;
        }

        private static Brush ResolveCellBg(ShiftCell cell, bool isHoliday)
        {
            var code = ShiftCodes.Normalize(cell?.Code);
            if (string.IsNullOrEmpty(code))
            {
                return isHoliday ? HolidayColumnBg : Brushes.White;
            }

            switch (code)
            {
                case ShiftCodes.Day: return isHoliday ? HolidayColumnBg : Brushes.White;
                case ShiftCodes.Card: return CellCardBg;
                case ShiftCodes.Deputy: return CellDeputyBg;
                case ShiftCodes.Infect: return CellInfectBg;
                case ShiftCodes.Big: return CellBigBg;
                case ShiftCodes.Small: return CellSmallBg;
                case ShiftCodes.Rest: return CellRestBg;
                case ShiftCodes.Public: return CellPublicBg;
                case ShiftCodes.Maternity: return CellMaternityBg;
                case ShiftCodes.Half: return CellHalfBg;
                default: return isHoliday ? HolidayColumnBg : Brushes.White;
            }
        }

        private Brush GetColumnStatBg(int day, bool matched)
        {
            if (_holidayByDay != null && day >= 0 && day < _holidayByDay.Length && _holidayByDay[day])
            {
                return HolidayColumnBg;
            }

            return matched ? HeaderBg : QuotaMismatchBg;
        }

        private static Brush ResolveCellFg(ShiftCell cell, bool isHoliday)
        {
            var code = ShiftCodes.Normalize(cell?.Code);
            switch (code)
            {
                case "副":
                case "大":
                case "小":
                    return Brushes.White;
                default: return Brushes.Black;
            }
        }

        private static string[] GetShiftPickerVisibleLabels()
        {
            return ShiftPickerVisibleLabels;
        }

        private static string DisplayCellCode(string code)
        {
            code = ShiftCodes.Normalize(code);
            if (string.IsNullOrEmpty(code)) return "·";
            if (code == ShiftCodes.Day) return "";
            return code == ShiftCodes.Big ? "大" : code;
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
                BorderBrush = BorderBg,
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
                BorderBrush = BorderBg,
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
                BorderBrush = BorderBg,
                BorderThickness = new Thickness(0.5),
                Child = child
            };
            Grid.SetRow(b, row); Grid.SetColumn(b, col);
            return b;
        }

        private static TextBlock GetBorderText(Border border)
        {
            return border?.Child as TextBlock;
        }

        // ===================== Cell click → popup picker =====================
        private void CellButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!TryGetCellTag(sender, out var empIdx, out var dayIdx))
            {
                return;
            }

            Focus();
            var key = ValueTuple.Create(empIdx, dayIdx);
            var modifiers = Keyboard.Modifiers;
            if ((modifiers & ModifierKeys.Shift) == ModifierKeys.Shift && _selectionAnchor.HasValue)
            {
                SelectCellRange(_selectionAnchor.Value, key, (modifiers & ModifierKeys.Control) == ModifierKeys.Control);
                _isSelectingCells = false;
                _dragSelectionMoved = true;
                SuppressNextCellClick();
                e.Handled = true;
                return;
            }

            if ((modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                ToggleSelectedCell(empIdx, dayIdx);
                _isSelectingCells = false;
                _dragSelectionMoved = true;
                SuppressNextCellClick();
                e.Handled = true;
                return;
            }

            SetSingleSelectedCell(empIdx, dayIdx);
            _isSelectingCells = true;
            _dragSelectionMoved = false;
        }

        private void CellButton_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isSelectingCells || !_selectionAnchor.HasValue || e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            var hit = ScheduleHost == null
                ? null
                : VisualTreeHelper.HitTest(ScheduleHost, e.GetPosition(ScheduleHost));
            var element = hit?.VisualHit ?? Mouse.DirectlyOver as DependencyObject;
            var button = FindAncestor<Button>(element);
            if (button == null || !TryGetCellTag(button, out var empIdx, out var dayIdx))
            {
                return;
            }

            var current = ValueTuple.Create(empIdx, dayIdx);
            if (_selectedCells.Count == 1 && _selectedCells.Contains(current))
            {
                return;
            }

            _dragSelectionMoved = true;
            if (sender is Button dragSource && !dragSource.IsMouseCaptured)
            {
                dragSource.CaptureMouse();
            }
            SelectCellRange(_selectionAnchor.Value, current, false);
            e.Handled = true;
        }

        private void CellButton_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var wasDragSelection = _isSelectingCells && _dragSelectionMoved;
            if (wasDragSelection && sender is Button button && button.IsMouseCaptured)
            {
                button.ReleaseMouseCapture();
            }

            if (wasDragSelection)
            {
                SuppressNextCellClick();
                e.Handled = true;
            }

            _isSelectingCells = false;
            _dragSelectionMoved = false;
        }

        private void SuppressNextCellClick()
        {
            _suppressNextCellClick = true;
            Dispatcher.BeginInvoke(new Action(() => _suppressNextCellClick = false), DispatcherPriority.Background);
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T typed)
                {
                    return typed;
                }

                if (current is Visual || current is System.Windows.Media.Media3D.Visual3D)
                {
                    current = VisualTreeHelper.GetParent(current);
                }
                else
                {
                    current = (current as FrameworkElement)?.Parent ?? (current as FrameworkContentElement)?.Parent;
                }
            }

            return null;
        }

        private void CellButton_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null) return;
            if (_suppressNextCellClick)
            {
                _suppressNextCellClick = false;
                return;
            }

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

            Add("白", () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Day), Brushes.White);
            var labels = GetShiftPickerVisibleLabels();
            for (var i = 1; i < labels.Length; i++)
            {
                var label = labels[i];
                if (label == "小1")
                {
                    panel.Children.Add(new Border { Width = 252, Height = 1, Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)), Margin = new Thickness(0, 4, 0, 4) });
                }

                switch (label)
                {
                    case "卡":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Card), new SolidColorBrush(Color.FromRgb(0xA7, 0xF3, 0xD0)));
                        break;
                    case "副":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Deputy), new SolidColorBrush(Color.FromRgb(0x99, 0x1B, 0x1B)), Brushes.White);
                        break;
                    case "感":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Infect), new SolidColorBrush(Color.FromRgb(0xFF, 0xEB, 0xEE)));
                        break;
                    case "大夜":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Big), new SolidColorBrush(Color.FromRgb(0x37, 0x47, 0x4F)), Brushes.White);
                        break;
                    case "小":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Small), new SolidColorBrush(Color.FromRgb(0x78, 0x90, 0x9C)), Brushes.White);
                        break;
                    case "休":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Rest), new SolidColorBrush(Color.FromRgb(0xFD, 0xEC, 0xC8)));
                        break;
                    case "公":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Public), new SolidColorBrush(Color.FromRgb(0xFE, 0xF3, 0xC7)));
                        break;
                    case "产假":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ShiftCodes.Maternity), new SolidColorBrush(Color.FromRgb(0xFE, 0xD7, 0xA8)));
                        break;
                    case "小1":
                        Add(label, () => _vm.ApplySmallNight1(empIdx, dayIdx), new SolidColorBrush(Color.FromRgb(0xFF, 0xCC, 0x80)));
                        break;
                    case "清空":
                        Add(label, () => _vm.SetCell(empIdx, dayIdx, ""), Brushes.WhiteSmoke);
                        break;
                }
            }

            border.Child = panel;
            popup.Child = border;
            popup.IsOpen = true;
        }

        private void DeleteEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null) return;
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
