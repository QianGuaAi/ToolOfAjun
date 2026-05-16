using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
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

        public ScheduleViewModel()
        {
            NewCommand = new RelayCommand(NewSchedule);
            EditCommand = new RelayCommand(EnterEditMode, () => Current != null && !IsEditing);
            SaveCommand = new AsyncRelayCommand(SaveAsync, () => Current != null && IsEditing);
            DeleteVersionCommand = new RelayParameterCommand(DeleteVersion);
            RefreshVersionsCommand = new RelayCommand(LoadVersions);
            AutoOptimizeCommand = new RelayCommand(AutoOptimize, () => Current != null && IsEditing);
            LoadVersionCommand = new AsyncRelayParameterCommand(LoadVersionAsync);
            ExportExcelCommand = new AsyncRelayCommand(ExportExcelAsync, () => Current != null);

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
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasCurrent));
                OnPropertyChanged(nameof(HeaderTitle));
                ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool HasCurrent => _current != null;

        public string HeaderTitle => _current == null
            ? "排班"
            : $"{_current.Year}-{_current.Month:00} · {_current.VersionName}（{(IsEditing ? "编辑中" : "查看")}）";

        // ============================ Edit mode ============================
        private bool _isEditing;
        public bool IsEditing
        {
            get => _isEditing;
            private set
            {
                _isEditing = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HeaderTitle));
                OnPropertyChanged(nameof(EditButtonLabel));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string EditButtonLabel => IsEditing ? "编辑中" : "编辑";

        private string _statusMessage = "尚未加载排班表，点击【新建】开始。";
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        // ============================ Commands ============================
        public ICommand NewCommand { get; }
        public ICommand EditCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand DeleteVersionCommand { get; }
        public ICommand RefreshVersionsCommand { get; }
        public ICommand AutoOptimizeCommand { get; }
        public ICommand LoadVersionCommand { get; }
        public ICommand ExportExcelCommand { get; }

        private async System.Threading.Tasks.Task ExportExcelAsync()
        {
            if (Current == null) return;
            try
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel 文件 (*.xlsx)|*.xlsx",
                    FileName = $"排班_{Current.Year}-{Current.Month:00}_{Current.VersionName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx",
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
                    : $"已新建 {year}-{month:00}/{versionName}，复制了 {sched.Employees.Count} 位人员。";
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
                if (sched == null)
                {
                    StatusMessage = "加载失败：文件读取错误。";
                    return;
                }
                Current = sched;
                IsEditing = false; // 加载后只读
                StatusMessage = $"已加载 {sched.Year}-{sched.Month:00}/{sched.VersionName}（只读，请点编辑修改）。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("LoadVersion failed: {Msg}", ex.Message);
                StatusMessage = "加载失败：" + ex.Message;
            }
        }

        // ============================ Edit ============================
        private void EnterEditMode()
        {
            if (Current == null) return;
            IsEditing = true;
            StatusMessage = "已进入编辑模式。";
        }

        // ============================ Save ============================
        private async Task SaveAsync()
        {
            if (Current == null) return;
            try
            {
                var path = await ScheduleService.SaveAsync(Current).ConfigureAwait(true);
                StatusMessage = $"已保存：{path}";
                LoadVersions();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("SaveSchedule failed: {Msg}", ex.Message);
                StatusMessage = "保存失败：" + ex.Message;
            }
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
        private void AutoOptimize()
        {
            if (Current == null || !IsEditing) return;
            var r = ShiftAutoOptimizer.Optimize(Current);
            ScheduleDataChanged?.Invoke(this, EventArgs.Empty);
            if (r.Success)
            {
                StatusMessage = r.Message + (r.Warnings.Count > 0 ? "（详情见日志）" : "");
                if (r.Warnings.Count > 0)
                {
                    foreach (var w in r.Warnings) AppLogService.Warning("Schedule warning: {W}", w);
                }
            }
            else
            {
                StatusMessage = r.Message;
            }
        }

        // ============================ Cell update / shortcuts ============================
        /// <summary>设置单元格为指定代码（用户手动）。</summary>
        public void SetCell(int empIdx, int dayIdx, string code)
        {
            if (Current == null || !IsEditing) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            if (dayIdx < 0 || dayIdx >= Current.DayCount) return;

            var cell = Current.Employees[empIdx].Cells[dayIdx];
            cell.Code = code ?? string.Empty;
            cell.IsManual = !string.IsNullOrEmpty(cell.Code);
            ScheduleDataChanged?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>大1：当天=大，前一天=小，后两天=休（小→大→休→休）。</summary>
        public void ApplyBigNight1(int empIdx, int dayIdx) => ApplyNightShift(empIdx, dayIdx, restDaysAfter: 2);

        /// <summary>大2：当天=大，前一天=小，后一天=休。</summary>
        public void ApplyBigNight2(int empIdx, int dayIdx) => ApplyNightShift(empIdx, dayIdx, restDaysAfter: 1);

        private void ApplyNightShift(int empIdx, int dayIdx, int restDaysAfter)
        {
            if (Current == null || !IsEditing) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            int days = Current.DayCount;
            if (dayIdx < 0 || dayIdx >= days) return;

            var emp = Current.Employees[empIdx];

            // 当天：大
            emp.Cells[dayIdx].Code = ShiftCodes.Big;
            emp.Cells[dayIdx].IsManual = true;

            // 前一天：小（如在月内）
            if (dayIdx - 1 >= 0)
            {
                emp.Cells[dayIdx - 1].Code = ShiftCodes.Small;
                emp.Cells[dayIdx - 1].IsManual = true;
            }

            // 后续 N 天：休
            for (int k = 1; k <= restDaysAfter; k++)
            {
                int idx = dayIdx + k;
                if (idx >= days) break;
                emp.Cells[idx].Code = ShiftCodes.Rest;
                emp.Cells[idx].IsManual = true;
            }

            ScheduleDataChanged?.Invoke(this, EventArgs.Empty);
        }

        // ============================ Employee management ============================
        public void AddEmployee(string name)
        {
            if (Current == null || !IsEditing) return;
            name = (name ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(name)) return;
            var row = new EmployeeRow { Name = name };
            for (int i = 0; i < Current.DayCount; i++) row.Cells.Add(new ShiftCell());
            Current.Employees.Add(row);
            ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
        }

        public void AddEmployees(System.Collections.Generic.IEnumerable<string> names)
        {
            if (Current == null || !IsEditing || names == null) return;
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
            if (added > 0) ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
        }

        public void RemoveEmployee(int empIdx)
        {
            if (Current == null || !IsEditing) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            Current.Employees.RemoveAt(empIdx);
            ScheduleStructureChanged?.Invoke(this, EventArgs.Empty);
        }

        public void UpdateEmployeeName(int empIdx, string newName)
        {
            if (Current == null || !IsEditing) return;
            if (empIdx < 0 || empIdx >= Current.Employees.Count) return;
            Current.Employees[empIdx].Name = (newName ?? string.Empty).Trim();
            ScheduleDataChanged?.Invoke(this, EventArgs.Empty);
        }

        public void UpdateDailyQuota(int dayIdx, int quota)
        {
            if (Current == null || !IsEditing) return;
            if (dayIdx < 0 || dayIdx >= Current.DailyRestQuotas.Count) return;
            Current.DailyRestQuotas[dayIdx] = Math.Max(0, quota);
            ScheduleDataChanged?.Invoke(this, EventArgs.Empty);
        }

        // ============================ Statistics ============================
        public (double work, double rest, int maxRun) ComputeRowStats(int empIdx)
        {
            double work = 0, rest = 0;
            int maxRun = 0, run = 0;
            if (Current == null || empIdx < 0 || empIdx >= Current.Employees.Count) return (0, 0, 0);
            foreach (var c in Current.Employees[empIdx].Cells)
            {
                work += ShiftCodes.WorkDays(c.Code);
                rest += ShiftCodes.RestDays(c.Code);
                if (ShiftCodes.IsWork(c.Code)) { run++; if (run > maxRun) maxRun = run; }
                else run = 0;
            }
            return (work, rest, maxRun);
        }

        public int ComputeColumnRestCount(int dayIdx)
        {
            if (Current == null) return 0;
            double sum = 0;
            foreach (var emp in Current.Employees)
            {
                if (dayIdx < emp.Cells.Count) sum += ShiftCodes.RestDays(emp.Cells[dayIdx].Code);
            }
            // 总休按 0.5 精度求和后向最近整数下取（午半天计 0.5 也允许显示）
            return (int)Math.Round(sum);
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
