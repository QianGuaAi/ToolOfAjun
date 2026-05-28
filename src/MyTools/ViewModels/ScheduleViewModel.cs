using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using MyTools.Services;

namespace MyTools.ViewModels
{
    /// <summary>
    /// 排班模块 ViewModel。负责：版本列表、新建/加载/保存、编辑模式切换、
    /// 单元格修改 + 夜班联动、自动设置休息日、统计实时刷新。
    /// </summary>
    public class ScheduleViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>当数据集结构（员工数 / 月份 / 列）变化时触发，View 应重建表格。</summary>
        public event EventHandler ScheduleStructureChanged;

        /// <summary>当统计需要刷新时触发（不重建结构）。</summary>
        public event EventHandler ScheduleDataChanged;
        public event EventHandler<ScheduleCellFocusRequestedEventArgs> ScheduleCellFocusRequested;

        public ScheduleViewModel()
        {
            NewCommand = new RelayCommand(NewSchedule);
            EditCommand = new RelayCommand(ToggleEditMode, () => Current != null);
            SaveCommand = new AsyncRelayCommand(SaveAsync, () => Current != null);
            SaveAsCommand = new AsyncRelayCommand(SaveAsAsync, () => Current != null);
            DeleteVersionCommand = new RelayParameterCommand(DeleteVersion);
            RefreshVersionsCommand = new RelayCommand(LoadVersions);
            AutoOptimizeCommand = new AsyncRelayCommand(AutoOptimizeAsync, () => Current != null && !IsAutoOptimizeWaiting);
            ClearAllCellsCommand = new RelayCommand(ClearAllCells, () => Current != null);
            ClearRandomCellsCommand = new RelayCommand(ClearRandomCells, () => Current != null);
            LoadVersionCommand = new AsyncRelayParameterCommand(LoadVersionAsync);
            ExportExcelCommand = new AsyncRelayCommand(ExportExcelAsync, () => Current != null);
            ImportEmployeeTemplateCommand = new AsyncRelayCommand(ImportEmployeeTemplateAsync, () => Current != null);
            ExportEmployeeTemplateCommand = new AsyncRelayCommand(ExportEmployeeTemplateAsync, () => Current != null && Current.Employees.Count > 0);
            CopyPreviousMonthEmployeesCommand = new AsyncRelayCommand(CopyPreviousMonthEmployeesAsync, () => Current != null);
            CopyEmployeeMonthCommand = new RelayCommand(CopyEmployeeMonth, () => Current != null && CopySourceEmployee != null && CopyTargetEmployee != null && !ReferenceEquals(CopySourceEmployee, CopyTargetEmployee));
            LocateScheduleConflictCommand = new RelayParameterCommand(LocateScheduleConflict, parameter => parameter is ScheduleConflictItem);

            LoadVersions();
        }

        // ============================ Versions list ============================
        public ObservableCollection<ScheduleVersionInfo> Versions { get; } = new ObservableCollection<ScheduleVersionInfo>();

        public void LoadVersions()
        {
            Versions.Clear();
            foreach (var v in ScheduleService.ListVersions()) Versions.Add(v);
        }

        // ============================ Current schedule ============================
        private ScheduleVersion _current;
        public ScheduleVersion Current
        {
            get => _current;
            private set
            {
                _current = value;
                var shouldEdit = _current != null;
                if (_isEditing != shouldEdit)
                {
                    _isEditing = shouldEdit;
                    OnPropertyChanged(nameof(IsEditing));
                    OnPropertyChanged(nameof(EditButtonLabel));
                }
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasCurrent));
                OnPropertyChanged(nameof(HasEmployees));
                OnPropertyChanged(nameof(HeaderTitle));
                CopySourceEmployee = null;
                CopyTargetEmployee = null;
                NormalizeScheduleRows(_current);
                RefreshScheduleConflicts();
                ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool HasCurrent => _current != null;
        public bool HasEmployees => _current != null && _current.Employees.Count > 0;

        public string HeaderTitle => _current == null
            ? "排班"
            : $"{_current.Year}-{_current.Month:00} · {_current.VersionName}";

        // ============================ Edit mode ============================
        private bool _isEditing;
        public bool IsEditing
        {
            get => _isEditing;
            private set
            {
                if (_isEditing == value) return;
                _isEditing = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HeaderTitle));
                OnPropertyChanged(nameof(EditButtonLabel));
                ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string EditButtonLabel => IsEditing ? "退出编辑" : "编辑";

        private string _statusMessage = "尚未加载排班表，点击【新建】开始。";
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public ObservableCollection<ScheduleConflictItem> ScheduleConflicts { get; } = new ObservableCollection<ScheduleConflictItem>();
        public ObservableCollection<ScheduleWeekRestItem> WeeklyRestStats { get; } = new ObservableCollection<ScheduleWeekRestItem>();
        private const int MaxVisibleScheduleConflicts = 300;
        private const int MaxWarningFreeOptimizeAttempts = 99;
        private const int MinMonthlyRestDays = 9;
        private int _scheduleConflictTotalCount;
        private bool _requireWarningFreeSchedule;
        private bool _isAutoOptimizeWaiting;
        private int _autoOptimizeAttempt;
        private int _autoOptimizeMaxAttempts = MaxWarningFreeOptimizeAttempts;
        private double _autoOptimizeProgress;
        private string _autoOptimizeWaitMessage = "正在生成正确的结果，请等待！";
        public bool HasScheduleConflicts => _scheduleConflictTotalCount > 0;
        public bool HasWeeklyRestStats => WeeklyRestStats.Count > 0;
        public string ScheduleConflictSummary => Current == null
            ? "未加载排班"
            : HasScheduleConflicts
                ? (_scheduleConflictTotalCount > ScheduleConflicts.Count
                    ? $"{_scheduleConflictTotalCount} 项需处理（显示前 {ScheduleConflicts.Count} 项）"
                    : $"{_scheduleConflictTotalCount} 项需处理")
                : "未发现明显冲突";

        public bool RequireWarningFreeSchedule
        {
            get => _requireWarningFreeSchedule;
            set
            {
                if (_requireWarningFreeSchedule == value) return;
                _requireWarningFreeSchedule = value;
                OnPropertyChanged();
            }
        }

        public bool IsAutoOptimizeWaiting
        {
            get => _isAutoOptimizeWaiting;
            private set
            {
                if (_isAutoOptimizeWaiting == value) return;
                _isAutoOptimizeWaiting = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public int AutoOptimizeAttempt
        {
            get => _autoOptimizeAttempt;
            private set
            {
                if (_autoOptimizeAttempt == value) return;
                _autoOptimizeAttempt = value;
                OnPropertyChanged();
            }
        }

        public int AutoOptimizeMaxAttempts
        {
            get => _autoOptimizeMaxAttempts;
            private set
            {
                if (_autoOptimizeMaxAttempts == value) return;
                _autoOptimizeMaxAttempts = value;
                OnPropertyChanged();
            }
        }

        public double AutoOptimizeProgress
        {
            get => _autoOptimizeProgress;
            private set
            {
                if (Math.Abs(_autoOptimizeProgress - value) < 0.001) return;
                _autoOptimizeProgress = value;
                OnPropertyChanged();
            }
        }

        public string AutoOptimizeWaitMessage
        {
            get => _autoOptimizeWaitMessage;
            private set
            {
                if (string.Equals(_autoOptimizeWaitMessage, value, StringComparison.Ordinal)) return;
                _autoOptimizeWaitMessage = value;
                OnPropertyChanged();
            }
        }

        public int SelectedConflictEmployeeIndex { get; private set; } = -1;
        public int SelectedConflictDayIndex { get; private set; } = -1;

        private EmployeeRow _copySourceEmployee;
        public EmployeeRow CopySourceEmployee
        {
            get => _copySourceEmployee;
            set
            {
                _copySourceEmployee = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private EmployeeRow _copyTargetEmployee;
        public EmployeeRow CopyTargetEmployee
        {
            get => _copyTargetEmployee;
            set
            {
                _copyTargetEmployee = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        // ============================ Commands ============================
        public ICommand NewCommand { get; }
        public ICommand EditCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand SaveAsCommand { get; }
        public ICommand DeleteVersionCommand { get; }
        public ICommand RefreshVersionsCommand { get; }
        public ICommand AutoOptimizeCommand { get; }
        public ICommand ClearAllCellsCommand { get; }
        public ICommand ClearRandomCellsCommand { get; }
        public ICommand LoadVersionCommand { get; }
        public ICommand ExportExcelCommand { get; }
        public ICommand ImportEmployeeTemplateCommand { get; }
        public ICommand ExportEmployeeTemplateCommand { get; }
        public ICommand CopyPreviousMonthEmployeesCommand { get; }
        public ICommand CopyEmployeeMonthCommand { get; }
        public ICommand LocateScheduleConflictCommand { get; }

        private async System.Threading.Tasks.Task ExportExcelAsync()
        {
            if (Current == null) return;
            try
            {
                if (!Current.HasGenerated)
                {
                    StatusMessage = "导出被拒绝：本月排班尚未完成自动生成。";
                    MessageBox.Show("本月排班尚未完成自动生成，不能导出 Excel。", "导出排班", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel 文件 (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx",
                    AddExtension = true,
                    OverwritePrompt = true,
                    FileName = BuildExportFileName(Current),
                    Title = "导出排班表"
                };
                if (dlg.ShowDialog() != true) return;

                await ScheduleExcelExporter.ExportAsync(Current, dlg.FileName).ConfigureAwait(true);
                StatusMessage = "已导出：" + dlg.FileName;

                var open = System.Windows.MessageBox.Show(
                    "导出成功：\n" + dlg.FileName + "\n\n是否打开该文件？",
                    "导出排班", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Information);
                if (open == System.Windows.MessageBoxResult.Yes)
                {
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true }); }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Schedule export failed: {Msg}", ex.Message);
                StatusMessage = "导出失败：" + ex.Message;
                System.Windows.MessageBox.Show(ex.Message, "导出失败", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private static string BuildExportFileName(ScheduleVersion schedule)
        {
            var versionName = SanitizeFileName(schedule?.VersionName);
            if (string.IsNullOrWhiteSpace(versionName))
            {
                versionName = "v1";
            }

            if (versionName.Length > 60)
            {
                versionName = versionName.Substring(0, 60).Trim();
            }

            var year = schedule?.Year ?? DateTime.Now.Year;
            var month = schedule?.Month ?? DateTime.Now.Month;
            return $"排班_{year}-{month:00}_{versionName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            var chars = value.Trim().Select(ch => invalidChars.Contains(ch) || char.IsControl(ch) ? '_' : ch).ToArray();
            return new string(chars).Trim().Trim('.');
        }

        private async Task ImportEmployeeTemplateAsync()
        {
            if (Current == null) return;
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "导入人员模板",
                Filter = "CSV / 文本 (*.csv;*.txt)|*.csv;*.txt|所有文件 (*.*)|*.*"
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                var text = await ReadAllTextAsync(dialog.FileName).ConfigureAwait(true);
                var names = ParseEmployeeNames(text)
                    .Where(name => !Current.Employees.Any(emp => string.Equals(emp.Name, name, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                AddEmployees(names);
                StatusMessage = names.Count > 0
                    ? $"已导入 {names.Count} 位人员。"
                    : "未导入新人员：文件为空或人员已存在。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Import schedule employee template failed: {Msg}", ex.Message);
                StatusMessage = "导入人员模板失败：" + ex.Message;
                MessageBox.Show(ex.Message, "导入人员模板失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ExportEmployeeTemplateAsync()
        {
            if (Current == null) return;
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "导出人员模板",
                Filter = "CSV 文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                AddExtension = true,
                OverwritePrompt = true,
                FileName = $"排班人员_{Current.Year}-{Current.Month:00}_{DateTime.Now:yyyyMMddHHmmss}.csv"
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                var lines = new List<string> { "姓名" };
                lines.AddRange(Current.Employees.Select(emp => EscapeCsv(emp.Name)));
                await WriteAllTextAsync(dialog.FileName, string.Join(Environment.NewLine, lines)).ConfigureAwait(true);
                StatusMessage = $"已导出人员模板：{dialog.FileName}";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Export schedule employee template failed: {Msg}", ex.Message);
                StatusMessage = "导出人员模板失败：" + ex.Message;
                MessageBox.Show(ex.Message, "导出人员模板失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task CopyPreviousMonthEmployeesAsync()
        {
            if (Current == null) return;
            try
            {
                var previous = await FindPreviousMonthScheduleAsync(Current.Year, Current.Month).ConfigureAwait(true);
                if (previous == null || previous.Employees.Count == 0)
                {
                    StatusMessage = "未找到上月人员名单。";
                    return;
                }

                var existing = new HashSet<string>(Current.Employees.Select(emp => emp.Name), StringComparer.OrdinalIgnoreCase);
                var names = previous.Employees
                    .Select(emp => emp.Name)
                    .Where(name => !string.IsNullOrWhiteSpace(name) && !existing.Contains(name.Trim()))
                    .ToList();
                AddEmployees(names);
                StatusMessage = names.Count > 0
                    ? $"已从上月复制 {names.Count} 位人员。"
                    : "上月人员已全部在当前排班中。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Copy previous month employees failed: {Msg}", ex.Message);
                StatusMessage = "复制上月人员失败：" + ex.Message;
            }
        }

        private void CopyEmployeeMonth()
        {
            if (Current == null || CopySourceEmployee == null || CopyTargetEmployee == null)
            {
                return;
            }

            if (ReferenceEquals(CopySourceEmployee, CopyTargetEmployee))
            {
                StatusMessage = "来源和目标人员不能相同。";
                return;
            }

            EnsureEmployeeCellCount(CopySourceEmployee, Current.DayCount);
            EnsureEmployeeCellCount(CopyTargetEmployee, Current.DayCount);
            for (var day = 0; day < Current.DayCount; day++)
            {
                CopyTargetEmployee.Cells[day].Code = CopySourceEmployee.Cells[day].Code;
                CopyTargetEmployee.Cells[day].IsManual = !string.IsNullOrEmpty(CopyTargetEmployee.Cells[day].Code);
            }

            StatusMessage = $"已将 {CopySourceEmployee.Name} 的整月班次复制给 {CopyTargetEmployee.Name}。";
            NotifyScheduleDataChanged();
        }

        private static async Task<string> ReadAllTextAsync(string filePath)
        {
            byte[] bytes;
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
            {
                bytes = new byte[stream.Length];
                var offset = 0;
                while (offset < bytes.Length)
                {
                    var read = await stream.ReadAsync(bytes, offset, bytes.Length - offset).ConfigureAwait(false);
                    if (read <= 0)
                    {
                        break;
                    }

                    offset += read;
                }
            }

            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                return Encoding.Default.GetString(bytes);
            }
        }

        private static async Task WriteAllTextAsync(string filePath, string text)
        {
            using (var writer = new StreamWriter(filePath, false, new UTF8Encoding(true)))
            {
                await writer.WriteAsync(text ?? string.Empty).ConfigureAwait(false);
            }
        }

        private static IEnumerable<string> ParseEmployeeNames(string text)
        {
            return (text ?? string.Empty)
                .Split(new[] { '\r', '\n', ',', ';', '\t', '，', '；' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim().Trim('"'))
                .Where(item => item.Length > 0 && !string.Equals(item, "姓名", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static string EscapeCsv(string value)
        {
            value = value ?? string.Empty;
            return value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0
                ? "\"" + value.Replace("\"", "\"\"") + "\""
                : value;
        }

        private static async Task<ScheduleVersion> FindPreviousMonthScheduleAsync(int year, int month)
        {
            var previousMonth = month == 1 ? 12 : month - 1;
            var previousYear = month == 1 ? year - 1 : year;
            var info = ScheduleService.ListVersions()
                .Where(item => item.Year == previousYear && item.Month == previousMonth)
                .OrderByDescending(item => item.UpdatedAt)
                .FirstOrDefault();
            return info == null ? null : await ScheduleService.LoadAsync(info.FilePath).ConfigureAwait(false);
        }

        // ============================ New ============================
        private async void NewSchedule()
        {
            try
            {
                var dlg = new MyTools.Views.NewScheduleDialog();
                var owner = Application.Current?.MainWindow;
                if (owner != null) dlg.Owner = owner;
                if (dlg.ShowDialog() != true) return;

                int year = dlg.SelectedYear;
                int month = dlg.SelectedMonth;
                string versionName = string.IsNullOrWhiteSpace(dlg.VersionName) ? "v1" : dlg.VersionName.Trim();

                if (ScheduleService.VersionExists(year, month, versionName))
                {
                    var resp = MessageBox.Show($"{year}-{month:00} 已有版本【{versionName}】。覆盖？",
                        "新建排班", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                    if (resp != MessageBoxResult.OK) return;
                }

                var sched = await ScheduleService.CreateNewAsync(year, month, versionName).ConfigureAwait(true);
                Current = sched;
                IsEditing = true;
                StatusMessage = sched.Employees.Count == 0
                    ? $"已新建 {year}-{month:00}/{versionName}（请先添加人员）。"
                    : $"已新建 {year}-{month:00}/{versionName}，复制了 {sched.Employees.Count} 位人员，可直接编辑。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("NewSchedule failed: {Msg}", ex.Message);
                StatusMessage = "新建失败：" + ex.Message;
            }
        }

        // ============================ Load ============================
        public async Task LoadVersionAsync(object parameter)
        {
            ScheduleVersionInfo info = parameter as ScheduleVersionInfo;
            if (info == null) return;
            try
            {
                var sched = await ScheduleService.LoadAsync(info.FilePath).ConfigureAwait(true);
                NormalizeScheduleRows(sched);
                if (sched == null)
                {
                    StatusMessage = "加载失败：文件读取错误。";
                    return;
                }
                Current = sched;
                IsEditing = true;
                StatusMessage = $"已加载 {sched.Year}-{sched.Month:00}/{sched.VersionName}，可直接编辑。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("LoadVersion failed: {Msg}", ex.Message);
                StatusMessage = "加载失败：" + ex.Message;
            }
        }

        // ============================ Edit ============================
        private void ToggleEditMode()
        {
            if (Current == null) return;
            EnterEditMode();
        }

        private void EnterEditMode()
        {
            if (Current == null) return;
            IsEditing = true;
            StatusMessage = "当前排班表可直接编辑：可单击班次格子选择班次，也可修改姓名和每日需休人数。";
        }

        private void ExitEditMode()
        {
            if (Current == null) return;
            IsEditing = true;
            StatusMessage = "当前排班表保持可编辑；如需保留修改请点击保存或另存。";
        }

        // ============================ Save ============================
        private async Task SaveAsync()
        {
            if (Current == null) return;
            try
            {
                var validationIssues = BuildSaveValidationIssues(Current).ToList();
                if (validationIssues.Count > 0)
                {
                    var detail = string.Join(Environment.NewLine, validationIssues.Take(12).Select(item => "· " + item));
                    if (validationIssues.Count > 12)
                    {
                        detail += Environment.NewLine + $"· 另有 {validationIssues.Count - 12} 项未显示";
                    }

                    StatusMessage = $"保存被阻止：{validationIssues.Count} 项基础数据校验未通过。";
                    MessageBox.Show(
                        "排班基础数据不完整，不能保存：\n\n" + detail,
                        "排班保存校验",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                var path = await ScheduleService.SaveAsync(Current).ConfigureAwait(true);
                StatusMessage = $"已保存：{path}";
                LoadVersions();
                RefreshScheduleConflicts();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("SaveSchedule failed: {Msg}", ex.Message);
                StatusMessage = "保存失败：" + ex.Message;
            }
        }

        private async Task SaveAsAsync()
        {
            if (Current == null) return;

            var defaultName = BuildDefaultSaveAsName(Current);
            var input = MyTools.Views.SimplePromptDialog.Prompt(
                $"请输入 {Current.Year}-{Current.Month:00} 下的新版本名：",
                "另存排班",
                defaultName);
            if (input == null) return;

            var newName = SanitizeFileName(input);
            if (string.IsNullOrWhiteSpace(newName))
            {
                StatusMessage = "另存失败：版本名不能为空。";
                MessageBox.Show("版本名不能为空。", "另存排班", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.Equals(newName, Current.VersionName, StringComparison.OrdinalIgnoreCase))
            {
                StatusMessage = "另存失败：请使用不同于当前表的新名字。";
                MessageBox.Show("请给当前排班表另起一个不同的名字。", "另存排班", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (ScheduleService.VersionExists(Current.Year, Current.Month, newName))
            {
                StatusMessage = $"另存失败：{Current.Year}-{Current.Month:00} 下已存在 {newName}。";
                MessageBox.Show($"{Current.Year}-{Current.Month:00} 下已存在【{newName}】，请换一个名字。", "另存排班", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var oldName = Current.VersionName;
            try
            {
                Current.VersionName = newName;
                OnPropertyChanged(nameof(HeaderTitle));
                var path = await ScheduleService.SaveAsync(Current).ConfigureAwait(true);
                LoadVersions();
                StatusMessage = $"已另存到当前月份：{path}";
            }
            catch (Exception ex)
            {
                Current.VersionName = oldName;
                OnPropertyChanged(nameof(HeaderTitle));
                AppLogService.Warning("SaveScheduleAs failed: {Msg}", ex.Message);
                StatusMessage = "另存失败：" + ex.Message;
            }
        }

        private static string BuildDefaultSaveAsName(ScheduleVersion schedule)
        {
            var baseName = SanitizeFileName(schedule?.VersionName);
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "v1";
            }

            var candidate = baseName + "-副本";
            for (var index = 2; index <= 99 && ScheduleService.VersionExists(schedule.Year, schedule.Month, candidate); index++)
            {
                candidate = $"{baseName}-副本{index}";
            }

            return candidate;
        }

        // ============================ Delete ============================
        private void DeleteVersion(object parameter)
        {
            var info = parameter as ScheduleVersionInfo;
            if (info == null) return;
            var resp = MessageBox.Show($"确认删除版本 {info.DisplayLabel}？",
                "删除版本", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (resp != MessageBoxResult.OK) return;
            ScheduleService.Delete(info.FilePath);
            LoadVersions();
            StatusMessage = $"已删除 {info.DisplayLabel}。";
        }

        // ============================ Auto optimize ============================
        private async Task AutoOptimizeAsync()
        {
            if (Current == null) return;
            if (!RequireWarningFreeSchedule)
            {
                AutoOptimizeCurrentOnce();
                return;
            }

            await AutoOptimizeUntilWarningFreeAsync().ConfigureAwait(true);
        }

        private void AutoOptimizeCurrentOnce()
        {
            if (Current == null) return;
            var r = ShiftAutoOptimizer.Optimize(Current, BuildPreservedCells(Current));
            NotifyScheduleDataChanged();
            if (!r.Success)
            {
                StatusMessage = r.Warnings.Count > 0 ? r.Message + " " + r.Warnings[0] : r.Message;
                foreach (var w in r.Warnings) AppLogService.Warning("Schedule optimize rejected: {W}", w);
                return;
            }

            var hardIssues = BuildHardConstraintIssues(Current).Take(6).ToList();
            if (hardIssues.Count > 0)
            {
                foreach (var issue in hardIssues)
                {
                    AppLogService.Warning("Schedule hard rule after optimize: {Issue}", issue);
                }

                StatusMessage = $"自动休息后仍有硬约束未满足：{hardIssues[0]}。请按冲突侧栏调整。";
                return;
            }

            StatusMessage = r.Message + (r.Warnings.Count > 0 ? "（详情见日志）" : "");
            if (r.Warnings.Count > 0)
            {
                foreach (var w in r.Warnings) AppLogService.Warning("Schedule warning: {W}", w);
            }
        }

        private async Task AutoOptimizeUntilWarningFreeAsync()
        {
            if (Current == null) return;

            var target = Current;
            var source = CloneScheduleVersion(target);
            var manualHardIssues = BuildManualHardConstraintIssues(source).ToList();
            if (manualHardIssues.Count > 0)
            {
                var detail = BuildIssuePreview(manualHardIssues, 6);
                StatusMessage = $"手工设置的作息已违反硬性规定，请检查：{manualHardIssues[0]}";
                MessageBox.Show(
                    "手工设置的作息已违反硬性规定，请检查！\n\n" + detail,
                    "设置休息日",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            AutoOptimizeMaxAttempts = MaxWarningFreeOptimizeAttempts;
            AutoOptimizeAttempt = 0;
            AutoOptimizeProgress = 0;
            AutoOptimizeWaitMessage = "正在生成正确的结果，请等待！";
            IsAutoOptimizeWaiting = true;

            var progress = new Progress<int>(attempt =>
            {
                AutoOptimizeAttempt = attempt;
                AutoOptimizeProgress = Math.Min(100, attempt * 100.0 / AutoOptimizeMaxAttempts);
                AutoOptimizeWaitMessage = $"正在生成正确的结果，请等待！已尝试 {attempt}/{AutoOptimizeMaxAttempts} 次";
            });

            try
            {
                var result = await Task.Run(() => GenerateWarningFreeCandidate(source, AutoOptimizeMaxAttempts, progress)).ConfigureAwait(true);
                if (!ReferenceEquals(Current, target))
                {
                    StatusMessage = "已切换到其他排班版本，本次自动生成结果已丢弃。";
                    return;
                }

                if (result.Schedule == null)
                {
                    var lastIssue = string.IsNullOrWhiteSpace(result.LastIssue)
                        ? "未找到明确的未通过项，请查看冲突侧栏。"
                        : "最后未通过项：" + result.LastIssue;
                    StatusMessage = $"已尝试 {result.Attempts} 次，仍未生成无警告结果。{lastIssue}";
                    AppLogService.Warning("Schedule warning-free optimize failed after {Attempts} attempts: {Issue}", result.Attempts, result.LastIssue);
                    MessageBox.Show(
                        $"已尝试 {result.Attempts} 次，仍未生成无警告结果。\n\n{lastIssue}",
                        "设置休息日",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                ApplyScheduleVersionCells(target, result.Schedule);
                NotifyScheduleDataChanged();
                StatusMessage = $"已尝试 {result.Attempts} 次，生成无警告结果。{result.Result.Message}";
            }
            finally
            {
                IsAutoOptimizeWaiting = false;
                AutoOptimizeWaitMessage = "正在生成正确的结果，请等待！";
                AutoOptimizeProgress = 0;
            }
        }

        private static (ScheduleVersion Schedule, ShiftAutoOptimizer.OptimizeResult Result, int Attempts, string LastIssue) GenerateWarningFreeCandidate(
            ScheduleVersion source,
            int maxAttempts,
            IProgress<int> progress)
        {
            string lastIssue = string.Empty;
            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var candidate = CloneScheduleVersion(source);
                var result = ShiftAutoOptimizer.Optimize(candidate, BuildPreservedCells(candidate));
                lastIssue = GetScheduleOptimizeIssue(candidate, result);
                if (string.IsNullOrWhiteSpace(lastIssue))
                {
                    progress?.Report(attempt);
                    return (candidate, result, attempt, string.Empty);
                }

                if (attempt == 1 || attempt % 5 == 0 || attempt == maxAttempts)
                {
                    progress?.Report(attempt);
                }
            }

            return (null, null, maxAttempts, lastIssue);
        }

        private static string GetScheduleOptimizeIssue(ScheduleVersion schedule, ShiftAutoOptimizer.OptimizeResult result)
        {
            if (result == null)
            {
                return "自动生成未返回结果。";
            }

            if (!result.Success)
            {
                return result.Warnings.Count > 0 ? result.Warnings[0] : result.Message;
            }

            if (result.Warnings.Count > 0)
            {
                return result.Warnings[0];
            }

            return BuildFirstScheduleConflictIssue(schedule);
        }

        private static string BuildFirstScheduleConflictIssue(ScheduleVersion schedule)
        {
            if (schedule == null)
            {
                return "当前排班为空。";
            }

            NormalizeScheduleRows(schedule);
            for (var day = 0; day < schedule.DayCount; day++)
            {
                var actual = ComputeColumnRestCount(schedule, day);
                var quota = day < schedule.DailyRestQuotas.Count ? schedule.DailyRestQuotas[day] : 0;
                if (actual > quota + 0.001)
                {
                    return $"{day + 1} 日总休超额：实际 {FormatNumber(actual)} / 目标 {FormatNumber(quota)}。";
                }

                if (!IsDailyRestWithinQuota(actual, quota))
                {
                    var minAllowed = MinDailyRestAllowed(quota);
                    return $"{day + 1} 日总休不足：实际 {FormatNumber(actual)} / 允许范围 {FormatNumber(minAllowed)}~{FormatNumber(quota)}。";
                }
            }

            for (var empIndex = 0; empIndex < schedule.Employees.Count; empIndex++)
            {
                var employee = schedule.Employees[empIndex];
                if (string.IsNullOrWhiteSpace(employee?.Name)) continue;

                var maxRun = ComputeMaxConsecutiveWork(employee);
                if (maxRun > 5)
                {
                    return $"{employee.Name} 连续上班 {maxRun} 天。";
                }

                var rest = employee.Cells == null ? 0 : employee.Cells.Sum(cell => ShiftCodes.RestDays(cell.Code));
                if (rest < MinMonthlyRestDays)
                {
                    return $"{employee.Name} 本月休息 {FormatNumber(rest)} 天，低于硬性要求 {MinMonthlyRestDays} 天。";
                }
            }

            var softIssue = BuildLowPriorityScheduleWarnings(schedule)
                .Select(item => string.IsNullOrWhiteSpace(item.Detail) ? item.Title : item.Title + "：" + item.Detail)
                .FirstOrDefault();
            return softIssue ?? string.Empty;
        }

        private static ScheduleVersion CloneScheduleVersion(ScheduleVersion source)
        {
            if (source == null)
            {
                return null;
            }

            return new ScheduleVersion
            {
                Year = source.Year,
                Month = source.Month,
                VersionName = source.VersionName,
                DailyRestQuotas = source.DailyRestQuotas == null ? new List<double>() : source.DailyRestQuotas.ToList(),
                Employees = source.Employees == null
                    ? new List<EmployeeRow>()
                    : source.Employees.Select(employee => new EmployeeRow
                    {
                        Name = employee?.Name ?? string.Empty,
                        Cells = employee?.Cells == null
                            ? new List<ShiftCell>()
                            : employee.Cells.Select(cell => new ShiftCell
                            {
                                Code = cell?.Code ?? string.Empty,
                                IsManual = cell?.IsManual ?? false
                            }).ToList()
                    }).ToList(),
                CreatedAt = source.CreatedAt,
                UpdatedAt = source.UpdatedAt,
                GeneratedAt = source.GeneratedAt
            };
        }

        private static void ApplyScheduleVersionCells(ScheduleVersion target, ScheduleVersion source)
        {
            if (target == null || source == null)
            {
                return;
            }

            NormalizeScheduleRows(target);
            NormalizeScheduleRows(source);
            for (var employeeIndex = 0; employeeIndex < target.Employees.Count && employeeIndex < source.Employees.Count; employeeIndex++)
            {
                var targetEmployee = target.Employees[employeeIndex];
                var sourceEmployee = source.Employees[employeeIndex];
                EnsureEmployeeCellCount(targetEmployee, target.DayCount);
                EnsureEmployeeCellCount(sourceEmployee, source.DayCount);
                for (var day = 0; day < target.DayCount && day < sourceEmployee.Cells.Count; day++)
                {
                    targetEmployee.Cells[day].Code = sourceEmployee.Cells[day].Code;
                    targetEmployee.Cells[day].IsManual = sourceEmployee.Cells[day].IsManual;
                }
            }

            target.GeneratedAt = source.GeneratedAt;
            target.UpdatedAt = DateTime.Now;
        }
        // ============================ Clear all cells ============================
        private void ClearAllCells()
        {
            if (Current == null) return;
            if (Current.Employees == null || Current.Employees.Count == 0) return;

            var resp = MessageBox.Show(
                "确认清空当前排班表的所有班次单元格？用户手动设置的单元格也会恢复为空白初始状态，此操作不可撤销。",
                "清空排班",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (resp != MessageBoxResult.OK) return;

            var cleared = 0;
            foreach (var emp in Current.Employees)
            {
                if (emp.Cells == null) continue;
                foreach (var cell in emp.Cells)
                {
                    if (cell == null) continue;
                    if (!string.IsNullOrEmpty(cell.Code) || cell.IsManual)
                    {
                        cleared++;
                    }

                    cell.Code = string.Empty;
                    cell.IsManual = false;
                }
            }

            Current.GeneratedAt = null;
            NotifyScheduleDataChanged();
            StatusMessage = $"已清空 {cleared} 个单元格，排班表已恢复为空白初始状态。";
        }

        // ============================ Clear random cells ============================
        private void ClearRandomCells()
        {
            if (Current == null) return;
            if (Current.Employees == null || Current.Employees.Count == 0) return;

            var resp = MessageBox.Show(
                "确认清空自动生成的随机单元格？用户手动设置的单元格会保留，此操作不可撤销。",
                "清空随机单元",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (resp != MessageBoxResult.OK) return;

            var cleared = 0;
            foreach (var emp in Current.Employees)
            {
                if (emp.Cells == null) continue;
                foreach (var cell in emp.Cells)
                {
                    if (cell == null || cell.IsManual) continue;
                    if (!string.IsNullOrEmpty(cell.Code))
                    {
                        cleared++;
                    }

                    cell.Code = string.Empty;
                    cell.IsManual = false;
                }
            }

            Current.GeneratedAt = null;
            NotifyScheduleDataChanged();
            StatusMessage = $"已清空 {cleared} 个随机单元格，用户设置的单元格已保留。";
        }

        // ============================ Cell update / shortcuts ============================
        /// <summary>设置单元格为指定代码（用户手动）。</summary>
        public void SetCell(int empIdx, int dayIdx, string code)
        {
            if (Current == null) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            if (dayIdx < 0 || dayIdx >= Current.DayCount) return;

            var cell = Current.Employees[empIdx].Cells[dayIdx];
            cell.Code = ShiftCodes.Normalize(code);
            cell.IsManual = !string.IsNullOrEmpty(cell.Code);
            NotifyScheduleDataChanged();
        }

        /// <summary>小1：从当前格开始设置 小→夜→休→休。</summary>
        public void ApplySmallNight1(int empIdx, int dayIdx)
        {
            if (Current == null) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            int days = Current.DayCount;
            if (dayIdx < 0 || dayIdx >= days) return;

            var emp = Current.Employees[empIdx];
            SetManualCell(emp, dayIdx, ShiftCodes.Small);
            SetManualCell(emp, dayIdx + 1, ShiftCodes.Big);
            SetManualCell(emp, dayIdx + 2, ShiftCodes.Rest);
            SetManualCell(emp, dayIdx + 3, ShiftCodes.Rest);

            NotifyScheduleDataChanged();
        }

        private static void SetManualCell(EmployeeRow employee, int dayIdx, string code)
        {
            if (employee == null || employee.Cells == null) return;
            if (dayIdx < 0 || dayIdx >= employee.Cells.Count) return;

            employee.Cells[dayIdx].Code = ShiftCodes.Normalize(code);
            employee.Cells[dayIdx].IsManual = !string.IsNullOrEmpty(employee.Cells[dayIdx].Code);
        }

        // ============================ Employee management ============================
        public void AddEmployee(string name)
        {
            if (Current == null) return;
            name = (name ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(name)) return;
            var row = new EmployeeRow { Name = name };
            for (int i = 0; i < Current.DayCount; i++) row.Cells.Add(new ShiftCell());
            Current.Employees.Add(row);
            NotifyScheduleStructureChanged();
        }

        public void AddEmployees(System.Collections.Generic.IEnumerable<string> names)
        {
            if (Current == null || names == null) return;
            int added = 0;
            foreach (var raw in names)
            {
                var name = (raw ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(name)) continue;
                var row = new EmployeeRow { Name = name };
                for (int i = 0; i < Current.DayCount; i++) row.Cells.Add(new ShiftCell());
                Current.Employees.Add(row);
                added++;
            }
            if (added > 0) NotifyScheduleStructureChanged();
        }

        public void RemoveEmployee(int empIdx)
        {
            if (Current == null) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            var removed = Current.Employees[empIdx];
            Current.Employees.RemoveAt(empIdx);
            if (ReferenceEquals(CopySourceEmployee, removed))
            {
                CopySourceEmployee = null;
            }

            if (ReferenceEquals(CopyTargetEmployee, removed))
            {
                CopyTargetEmployee = null;
            }

            NotifyScheduleStructureChanged();
        }

        public void UpdateEmployeeName(int empIdx, string newName)
        {
            if (Current == null) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            Current.Employees[empIdx].Name = (newName ?? string.Empty).Trim();
            NotifyScheduleDataChanged();
        }

        public void UpdateDailyQuota(int dayIdx, double quota)
        {
            if (Current == null) return;
            if (dayIdx < 0 || dayIdx >= Current.DailyRestQuotas.Count) return;
            Current.DailyRestQuotas[dayIdx] = Math.Max(0, quota);
            NotifyScheduleDataChanged();
        }

        // ============================ Statistics ============================
        public (double work, double rest, int maxRun) ComputeRowStats(int empIdx)
        {
            double work = 0, rest = 0;
            int maxRun = 0, run = 0;
            if (Current == null || empIdx < 0 || empIdx >= Current.Employees.Count) return (0, 0, 0);
            if (string.IsNullOrWhiteSpace(Current.Employees[empIdx].Name)) return (0, 0, 0);
            foreach (var c in Current.Employees[empIdx].Cells)
            {
                work += ShiftCodes.WorkDays(c.Code);
                rest += ShiftCodes.RestDays(c.Code);
                if (ShiftCodes.IsWork(c.Code)) { run++; if (run > maxRun) maxRun = run; }
                else run = 0;
            }
            return (work, rest, maxRun);
        }

        public double ComputeColumnRestCount(int dayIdx)
        {
            if (Current == null) return 0;
            double sum = 0;
            foreach (var emp in Current.Employees)
            {
                if (string.IsNullOrWhiteSpace(emp.Name)) continue;
                if (dayIdx < emp.Cells.Count) sum += ShiftCodes.RestDays(emp.Cells[dayIdx].Code);
            }
            return sum;
        }

        private void NotifyScheduleStructureChanged()
        {
            RefreshScheduleConflicts();
            OnPropertyChanged(nameof(Current));
            OnPropertyChanged(nameof(HasEmployees));
            CommandManager.InvalidateRequerySuggested();
            ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
        }

        private void NotifyScheduleDataChanged()
        {
            RefreshScheduleConflicts();
            OnPropertyChanged(nameof(Current));
            ScheduleDataChanged?.Invoke(this, EventArgs.Empty);
        }

        private static IReadOnlyCollection<(int emp, int day)> BuildPreservedCells(ScheduleVersion schedule)
        {
            return new List<(int emp, int day)>();
        }

        private static IEnumerable<string> BuildManualHardConstraintIssues(ScheduleVersion schedule)
        {
            if (schedule == null)
            {
                yield break;
            }

            NormalizeScheduleRows(schedule);
            for (var day = 0; day < schedule.DayCount; day++)
            {
                var quota = schedule.DailyRestQuotas != null && day < schedule.DailyRestQuotas.Count
                    ? schedule.DailyRestQuotas[day]
                    : 0;
                double manualRest = 0;
                foreach (var employee in schedule.Employees)
                {
                    if (string.IsNullOrWhiteSpace(employee?.Name)) continue;
                    if (employee.Cells == null || day >= employee.Cells.Count) continue;
                    var cell = employee.Cells[day];
                    if (cell != null && cell.IsManual)
                    {
                        manualRest += ShiftCodes.RestDays(cell.Code);
                    }
                }

                if (manualRest > quota + 0.001)
                {
                    yield return $"{day + 1} 日手工休息 {FormatNumber(manualRest)} / 总休目标 {FormatNumber(quota)}，多 {FormatNumber(manualRest - quota)}";
                }
            }

            for (var empIndex = 0; empIndex < schedule.Employees.Count; empIndex++)
            {
                var employee = schedule.Employees[empIndex];
                if (string.IsNullOrWhiteSpace(employee?.Name) || employee.Cells == null) continue;

                var run = 0;
                var runStart = 0;
                for (var day = 0; day < employee.Cells.Count; day++)
                {
                    var cell = employee.Cells[day];
                    if (cell != null && cell.IsManual && ShiftCodes.IsWork(cell.Code))
                    {
                        if (run == 0)
                        {
                            runStart = day;
                        }

                        run++;
                        if (run == 6)
                        {
                            yield return $"{employee.Name} 手工设置 {runStart + 1}-{day + 1} 日连续上班 {run} 天，超过 5 天上限";
                            break;
                        }
                    }
                    else
                    {
                        run = 0;
                    }
                }
            }
        }

        private static string BuildIssuePreview(IReadOnlyList<string> issues, int maxCount)
        {
            if (issues == null || issues.Count == 0)
            {
                return string.Empty;
            }

            var lines = issues.Take(maxCount).Select(item => "· " + item).ToList();
            if (issues.Count > maxCount)
            {
                lines.Add($"· 另有 {issues.Count - maxCount} 项未显示");
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static IEnumerable<string> BuildSaveValidationIssues(ScheduleVersion schedule)
        {
            if (schedule == null)
            {
                yield return "当前排班为空。";
                yield break;
            }

            if (schedule.Year < 1900 || schedule.Year > 2100)
            {
                yield return $"年份 {schedule.Year} 无效。";
                yield break;
            }

            if (schedule.Month < 1 || schedule.Month > 12)
            {
                yield return $"月份 {schedule.Month} 无效。";
                yield break;
            }

            var dayCount = DateTime.DaysInMonth(schedule.Year, schedule.Month);
            if (schedule.DailyRestQuotas == null)
            {
                yield return "每日总休数组为空。";
            }
            else
            {
                if (schedule.DailyRestQuotas.Count != dayCount)
                {
                    yield return $"每日总休数组长度 {schedule.DailyRestQuotas.Count} 与当月天数 {dayCount} 不一致。";
                }

                for (var day = 0; day < schedule.DailyRestQuotas.Count; day++)
                {
                    var quota = schedule.DailyRestQuotas[day];
                    if (double.IsNaN(quota) || double.IsInfinity(quota) || quota < 0)
                    {
                        yield return $"{day + 1} 日总休值 {quota} 无效。";
                    }
                    else if (Math.Abs(quota * 2 - Math.Round(quota * 2)) > 0.001)
                    {
                        yield return $"{day + 1} 日总休 {FormatNumber(quota)} 不是 0.5 步进。";
                    }
                }
            }

            if (schedule.Employees == null)
            {
                yield return "人员行集合为空。";
                yield break;
            }

            for (var row = 0; row < schedule.Employees.Count; row++)
            {
                var employee = schedule.Employees[row];
                if (employee == null)
                {
                    yield return $"第 {row + 1} 个人员行为空。";
                    continue;
                }

                if (employee.Cells == null)
                {
                    yield return $"{DisplayEmployeeName(employee, row)} 的日期单元格为空。";
                }
                else if (employee.Cells.Count != dayCount)
                {
                    yield return $"{DisplayEmployeeName(employee, row)} 的日期单元格数量 {employee.Cells.Count} 与当月天数 {dayCount} 不一致。";
                }
            }
        }

        private static string DisplayEmployeeName(EmployeeRow employee, int index)
        {
            return string.IsNullOrWhiteSpace(employee?.Name) ? $"第 {index + 1} 行" : employee.Name.Trim();
        }

        private static double MinDailyRestAllowed(double quota)
        {
            return Math.Max(0, quota - 0.5);
        }

        private static bool IsDailyRestWithinQuota(double actual, double quota)
        {
            return actual >= MinDailyRestAllowed(quota) - 0.001 && actual <= quota + 0.001;
        }

        private static IEnumerable<string> BuildHardConstraintIssues(ScheduleVersion schedule)
        {
            if (schedule == null)
            {
                yield break;
            }

            NormalizeScheduleRows(schedule);
            for (var day = 0; day < schedule.DayCount; day++)
            {
                var actual = ComputeColumnRestCount(schedule, day);
                var quota = day < schedule.DailyRestQuotas.Count ? schedule.DailyRestQuotas[day] : 0;
                if (actual > quota + 0.001)
                {
                    yield return $"{day + 1} 日实际休息 {FormatNumber(actual)} / 总休目标 {FormatNumber(quota)}，多 {FormatNumber(actual - quota)}";
                }
                else if (!IsDailyRestWithinQuota(actual, quota))
                {
                    var minAllowed = MinDailyRestAllowed(quota);
                    yield return $"{day + 1} 日实际休息 {FormatNumber(actual)} / 总休目标 {FormatNumber(quota)}，少 {FormatNumber(minAllowed - actual)}，允许范围 {FormatNumber(minAllowed)}~{FormatNumber(quota)}";
                }
            }

            for (var empIndex = 0; empIndex < schedule.Employees.Count; empIndex++)
            {
                var employee = schedule.Employees[empIndex];
                if (string.IsNullOrWhiteSpace(employee?.Name)) continue;
                var maxRun = ComputeMaxConsecutiveWork(employee);
                if (maxRun > 5)
                {
                    yield return $"{employee.Name} 连续上班 {maxRun} 天，超过 5 天上限";
                }

                var rest = employee.Cells == null ? 0 : employee.Cells.Sum(cell => ShiftCodes.RestDays(cell.Code));
                if (rest < MinMonthlyRestDays)
                {
                    yield return $"{employee.Name} 本月休息 {FormatNumber(rest)} 天，低于硬性要求 {MinMonthlyRestDays} 天";
                }
            }
        }

        private static double ComputeColumnRestCount(ScheduleVersion schedule, int dayIdx)
        {
            if (schedule == null || dayIdx < 0)
            {
                return 0;
            }

            double sum = 0;
            foreach (var emp in schedule.Employees)
            {
                if (string.IsNullOrWhiteSpace(emp.Name)) continue;
                if (emp.Cells != null && dayIdx < emp.Cells.Count)
                {
                    sum += ShiftCodes.RestDays(emp.Cells[dayIdx].Code);
                }
            }

            return sum;
        }

        private static int ComputeMaxConsecutiveWork(EmployeeRow employee)
        {
            if (employee == null || employee.Cells == null)
            {
                return 0;
            }

            var maxRun = 0;
            var run = 0;
            foreach (var cell in employee.Cells)
            {
                if (ShiftCodes.IsWork(cell.Code))
                {
                    run++;
                    if (run > maxRun)
                    {
                        maxRun = run;
                    }
                }
                else
                {
                    run = 0;
                }
            }

            return maxRun;
        }

        private static IEnumerable<ScheduleConflictItem> BuildLowPriorityScheduleWarnings(ScheduleVersion schedule)
        {
            if (schedule == null)
            {
                yield break;
            }

            NormalizeScheduleRows(schedule);
            for (var day = 0; day < schedule.DayCount; day++)
            {
                var specialCodes = new List<string>();
                foreach (var employee in schedule.Employees)
                {
                    if (string.IsNullOrWhiteSpace(employee?.Name) || employee.Cells == null || day >= employee.Cells.Count) continue;
                    var code = ShiftCodes.Normalize(employee.Cells[day].Code);
                    if (IsSpecialWarningShift(code))
                    {
                        specialCodes.Add(code);
                    }
                }

                if (specialCodes.Count >= 2)
                {
                    var detail = string.Join("、", specialCodes
                        .GroupBy(code => code)
                        .OrderBy(group => SpecialWarningShiftOrder(group.Key))
                        .Select(group => group.Key + group.Count() + "次"));
                    yield return new ScheduleConflictItem
                    {
                        Level = "低",
                        Title = $"{day + 1} 日副/卡/感合计 {specialCodes.Count} 次",
                        Detail = $"副、卡、感任意一项每出现一次计 1 次（{detail}），建议低优先级调整。",
                        Category = "特殊班次",
                        DayIndex = day
                    };
                }
            }

            for (var empIndex = 0; empIndex < schedule.Employees.Count; empIndex++)
            {
                var employee = schedule.Employees[empIndex];
                if (string.IsNullOrWhiteSpace(employee?.Name)) continue;
                for (var day = 0; day < schedule.DayCount; day++)
                {
                    if (IsFullRest(employee, day) && !IsFullRest(employee, day - 1) && !IsFullRest(employee, day + 1))
                    {
                        yield return new ScheduleConflictItem
                        {
                            Level = "低",
                            Title = $"{employee.Name} {day + 1} 日单天休",
                            Detail = "建议将休息日尽量安排为连续 2 天或以上。",
                            Category = "休息连续性",
                            EmployeeIndex = empIndex,
                            DayIndex = day
                        };
                    }

                    if (ShiftCodes.IsWork(employee.Cells[day].Code) && IsFullRest(employee, day - 1) && IsFullRest(employee, day + 1))
                    {
                        yield return new ScheduleConflictItem
                        {
                            Level = "低",
                            Title = $"{employee.Name} {day + 1} 日上一天夹在休息日之间",
                            Detail = "建议尽量避免上一天休一天的节奏。",
                            Category = "休息连续性",
                            EmployeeIndex = empIndex,
                            DayIndex = day
                        };
                    }
                }
            }
        }

        private static bool IsFullRest(EmployeeRow employee, int day)
        {
            return employee != null
                && employee.Cells != null
                && day >= 0
                && day < employee.Cells.Count
                && ShiftCodes.RestDays(employee.Cells[day].Code) >= 1.0;
        }

        private static bool IsSpecialWarningShift(string code)
        {
            code = ShiftCodes.Normalize(code);
            return code == ShiftCodes.Deputy || code == ShiftCodes.Card || code == ShiftCodes.Infect;
        }

        private static int SpecialWarningShiftOrder(string code)
        {
            code = ShiftCodes.Normalize(code);
            if (code == ShiftCodes.Deputy) return 0;
            if (code == ShiftCodes.Card) return 1;
            if (code == ShiftCodes.Infect) return 2;
            return 3;
        }

        private void RefreshScheduleConflicts()
        {
            ScheduleConflicts.Clear();
            WeeklyRestStats.Clear();
            _scheduleConflictTotalCount = 0;
            if (Current == null)
            {
                OnPropertyChanged(nameof(HasScheduleConflicts));
                OnPropertyChanged(nameof(ScheduleConflictSummary));
                OnPropertyChanged(nameof(HasWeeklyRestStats));
                return;
            }

            RefreshWeeklyRestStats();
            for (var day = 0; day < Current.DayCount; day++)
            {
                var actual = ComputeColumnRestCount(day);
                var quota = day < Current.DailyRestQuotas.Count ? Current.DailyRestQuotas[day] : 0;
                if (actual > quota + 0.001)
                {
                    AddScheduleConflict(new ScheduleConflictItem
                    {
                        Level = "高",
                        Title = $"{day + 1} 日总休超额",
                        Detail = $"实际 {FormatNumber(actual)} / 目标 {FormatNumber(quota)}，多 {FormatNumber(actual - quota)}。",
                        Category = "每日总休",
                        DayIndex = day
                    });
                }
                else if (!IsDailyRestWithinQuota(actual, quota))
                {
                    var minAllowed = MinDailyRestAllowed(quota);
                    AddScheduleConflict(new ScheduleConflictItem
                    {
                        Level = "高",
                        Title = $"{day + 1} 日总休不足",
                        Detail = $"实际 {FormatNumber(actual)} / 目标 {FormatNumber(quota)}，允许范围 {FormatNumber(minAllowed)}~{FormatNumber(quota)}，少 {FormatNumber(minAllowed - actual)}。",
                        Category = "每日总休",
                        DayIndex = day
                    });
                }

                // 规则：同一天"副"+"卡"+"感"同时出现次数 >= 2 时警告
                int specialCount = 0;
                foreach (var emp in Current.Employees)
                {
                    if (string.IsNullOrWhiteSpace(emp?.Name)) continue;
                    if (day < emp.Cells.Count)
                    {
                        var code = ShiftCodes.Normalize(emp.Cells[day].Code);
                        if (code == ShiftCodes.Deputy || code == ShiftCodes.Card || code == ShiftCodes.Infect)
                        {
                            specialCount++;
                        }
                    }
                }
                if (specialCount >= 2)
                {
                    AddScheduleConflict(new ScheduleConflictItem
                    {
                        Level = "高",
                        Title = $"{day + 1} 日副/卡/感同天 ≥ 2 人",
                        Detail = $"副+卡+感 共 {specialCount} 人同时排班，疑似重复安排。",
                        Category = "特殊班次",
                        DayIndex = day
                    });
                }
            }

            for (var empIndex = 0; empIndex < Current.Employees.Count; empIndex++)
            {
                var employee = Current.Employees[empIndex];
                if (string.IsNullOrWhiteSpace(employee?.Name)) continue;
                var stats = ComputeRowStats(empIndex);
                if (stats.maxRun > 5)
                {
                    AddScheduleConflict(new ScheduleConflictItem
                    {
                        Level = "高",
                        Title = $"{employee.Name} 连续上班 {stats.maxRun} 天",
                        Detail = "建议控制在 5 天以内。",
                        Category = "连续上班",
                        EmployeeIndex = empIndex
                    });
                }

                if (stats.rest < MinMonthlyRestDays)
                {
                    AddScheduleConflict(new ScheduleConflictItem
                    {
                        Level = "高",
                        Title = $"{employee.Name} 休息不足",
                        Detail = $"本月休息 {FormatNumber(stats.rest)} 天，低于硬性要求 {MinMonthlyRestDays} 天。",
                        Category = "人员休息",
                        EmployeeIndex = empIndex
                    });
                }
            }

            foreach (var warning in BuildLowPriorityScheduleWarnings(Current))
            {
                AddScheduleConflict(warning);
            }

            OnPropertyChanged(nameof(HasScheduleConflicts));
            OnPropertyChanged(nameof(ScheduleConflictSummary));
            OnPropertyChanged(nameof(HasWeeklyRestStats));
        }

        private void AddScheduleConflict(ScheduleConflictItem item)
        {
            _scheduleConflictTotalCount++;
            if (ScheduleConflicts.Count < MaxVisibleScheduleConflicts)
            {
                ScheduleConflicts.Add(item);
            }
        }

        private void RefreshWeeklyRestStats()
        {
            if (Current == null)
            {
                return;
            }

            var weekIndex = 1;
            for (var start = 0; start < Current.DayCount; start += 7)
            {
                var end = Math.Min(Current.DayCount - 1, start + 6);
                var employeeRests = Current.Employees
                    .Where(employee => !string.IsNullOrWhiteSpace(employee.Name))
                    .Select(employee => ComputeEmployeeRestRange(employee, start, end))
                    .ToList();
                var totalRest = employeeRests.Sum();
                var minRest = employeeRests.Count == 0 ? 0 : employeeRests.Min();
                var maxRest = employeeRests.Count == 0 ? 0 : employeeRests.Max();
                WeeklyRestStats.Add(new ScheduleWeekRestItem
                {
                    WeekText = $"第 {weekIndex} 周",
                    DateRangeText = $"{start + 1}-{end + 1} 日",
                    TotalRestText = FormatNumber(totalRest),
                    BalanceText = $"人均 {FormatNumber(employeeRests.Count == 0 ? 0 : totalRest / employeeRests.Count)} / 差 {FormatNumber(maxRest - minRest)}"
                });
                weekIndex++;
            }
        }

        private static double ComputeEmployeeRestRange(EmployeeRow employee, int startDay, int endDay)
        {
            if (employee == null || employee.Cells == null)
            {
                return 0;
            }

            double total = 0;
            for (var day = Math.Max(0, startDay); day <= endDay && day < employee.Cells.Count; day++)
            {
                total += ShiftCodes.RestDays(employee.Cells[day].Code);
            }

            return total;
        }

        private void LocateScheduleConflict(object parameter)
        {
            var item = parameter as ScheduleConflictItem;
            if (item == null)
            {
                return;
            }

            ResolveConflictLocation(item);
            SelectedConflictEmployeeIndex = item.EmployeeIndex;
            SelectedConflictDayIndex = item.DayIndex;
            OnPropertyChanged(nameof(SelectedConflictEmployeeIndex));
            OnPropertyChanged(nameof(SelectedConflictDayIndex));
            ScheduleCellFocusRequested?.Invoke(this, new ScheduleCellFocusRequestedEventArgs(item.EmployeeIndex, item.DayIndex));
        }

        private void ResolveConflictLocation(ScheduleConflictItem item)
        {
            if (item == null || Current == null)
            {
                return;
            }

            if (item.DayIndex < 0)
            {
                item.DayIndex = ExtractFirstNumber(item.Title) - 1;
                if (item.DayIndex < 0 || item.DayIndex >= Current.DayCount)
                {
                    item.DayIndex = -1;
                }
            }

            if (item.EmployeeIndex < 0)
            {
                item.EmployeeIndex = Current.Employees
                    .Select((employee, index) => new { employee, index })
                    .FirstOrDefault(pair => !string.IsNullOrWhiteSpace(pair.employee.Name) &&
                                            (item.Title ?? string.Empty).IndexOf(pair.employee.Name, StringComparison.OrdinalIgnoreCase) >= 0)
                    ?.index ?? -1;
            }
        }

        private static int ExtractFirstNumber(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return 0;
            }

            var digits = new StringBuilder();
            foreach (var ch in value)
            {
                if (char.IsDigit(ch))
                {
                    digits.Append(ch);
                }
                else if (digits.Length > 0)
                {
                    break;
                }
            }

            return int.TryParse(digits.ToString(), out var number) ? number : 0;
        }

        private static string FormatNumber(double value)
        {
            return value % 1 == 0 ? ((int)value).ToString() : value.ToString("0.#");
        }

        private static void NormalizeScheduleRows(ScheduleVersion schedule)
        {
            if (schedule == null)
            {
                return;
            }

            if (schedule.DailyRestQuotas == null)
            {
                schedule.DailyRestQuotas = new List<double>();
            }

            while (schedule.DailyRestQuotas.Count < schedule.DayCount)
            {
                schedule.DailyRestQuotas.Add(0);
            }

            foreach (var employee in schedule.Employees)
            {
                EnsureEmployeeCellCount(employee, schedule.DayCount);
            }
        }

        private static void EnsureEmployeeCellCount(EmployeeRow employee, int dayCount)
        {
            if (employee == null)
            {
                return;
            }

            if (employee.Cells == null)
            {
                employee.Cells = new List<ShiftCell>();
            }

            while (employee.Cells.Count < dayCount)
            {
                employee.Cells.Add(new ShiftCell());
            }

            while (employee.Cells.Count > dayCount)
            {
                employee.Cells.RemoveAt(employee.Cells.Count - 1);
            }

            for (var i = 0; i < employee.Cells.Count; i++)
            {
                if (employee.Cells[i] == null)
                {
                    employee.Cells[i] = new ShiftCell();
                }
                employee.Cells[i].Code = ShiftCodes.Normalize(employee.Cells[i].Code);
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public sealed class ScheduleConflictItem
    {
        public string Level { get; set; }
        public string Category { get; set; }
        public string Title { get; set; }
        public string Detail { get; set; }
        public int EmployeeIndex { get; set; } = -1;
        public int DayIndex { get; set; } = -1;
        public bool CanLocate => EmployeeIndex >= 0 || DayIndex >= 0;
    }

    public sealed class ScheduleWeekRestItem
    {
        public string WeekText { get; set; }
        public string DateRangeText { get; set; }
        public string TotalRestText { get; set; }
        public string BalanceText { get; set; }
    }

    public sealed class ScheduleCellFocusRequestedEventArgs : EventArgs
    {
        public ScheduleCellFocusRequestedEventArgs(int employeeIndex, int dayIndex)
        {
            EmployeeIndex = employeeIndex;
            DayIndex = dayIndex;
        }

        public int EmployeeIndex { get; }
        public int DayIndex { get; }
    }
}
