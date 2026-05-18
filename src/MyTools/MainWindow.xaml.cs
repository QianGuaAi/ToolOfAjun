using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using MahApps.Metro.Controls;
using MyTools.Services;
using MyTools.Shared;
using MyTools.ViewModels;

namespace MyTools
{
    public partial class MainWindow : MetroWindow
    {
        private const int WM_HOTKEY = 0x0312;
        private HwndSource _windowSource;
        private readonly DispatcherTimer _videoProgressTimer;
        private bool _isDraggingVideoProgress;
        private bool _isDraggingVideoWaveform;
        private bool _isSelectingVideoWaveformRange;
        private bool _isDraggingVideoWaveformRangeHandle;
        private bool _isDraggingVideoWaveformStartHandle;
        private double _videoWaveformRangeStartSeconds;
        private Point _videoPlaylistDragStartPoint;
        private VideoPlaylistItem _videoPlaylistDragItem;
        private Point _convertQueueDragStartPoint;
        private ConvertQueueItem _convertQueueDragItem;

        public MainWindow()
        {
            InitializeComponent();
            _videoProgressTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(350)
            };
            _videoProgressTimer.Tick += VideoProgressTimer_OnTick;
            InitializeTrayIcon();
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.PropertyChanged += ViewModel_OnPropertyChanged;
            }
        }

        private void InitializeTrayIcon()
        {
            var executablePath = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                return;
            }

            var icon = System.Drawing.Icon.ExtractAssociatedIcon(executablePath);
            if (icon != null)
            {
                TrayIcon.Icon = icon;
            }
        }

        private void SqlPasswordBox_OnPasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel && sender is PasswordBox passwordBox)
            {
                viewModel.SqlPassword = passwordBox.Password;
            }
        }

        private void SqlPasswordHistoryButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.IsOpen = true;
            }
        }

        private void SqlPasswordHistoryMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel && sender is MenuItem menuItem && menuItem.DataContext is string password)
            {
                viewModel.SqlPassword = password;
                SqlPasswordBox.Password = password;
            }
        }

        private void CopySqlQueryResultButton_OnClick(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as MainViewModel;
            if (viewModel == null || viewModel.SqlQueryResult == null || viewModel.SqlQueryResult.Count == 0)
            {
                MessageBox.Show("没有可复制的查询结果。", "复制结果", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var text = BuildSqlQueryResultClipboardText(viewModel.SqlQueryResult, SqlQueryResultGrid);
                if (string.IsNullOrWhiteSpace(text))
                {
                    MessageBox.Show("没有可复制的查询结果。", "复制结果", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Clipboard.SetText(text);
                viewModel.QueryStatusMessage = SqlQueryResultGrid.SelectedCells.Count > 0
                    ? $"已复制选中结果，共 {SqlQueryResultGrid.SelectedCells.Count} 个单元格。"
                    : $"已复制全部查询结果，共 {viewModel.SqlQueryResult.Count} 行。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Copy SQL query result failed: {Msg}", ex.Message);
                MessageBox.Show(ex.Message, "复制结果失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string BuildSqlQueryResultClipboardText(DataView result, DataGrid grid)
        {
            var table = result.Table;
            if (grid != null && grid.SelectedCells.Count > 0)
            {
                return BuildSelectedSqlQueryClipboardText(grid);
            }

            var builder = new StringBuilder();
            AppendSqlQueryHeader(builder, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
            foreach (DataRowView rowView in result)
            {
                AppendSqlQueryRow(builder, table.Columns.Cast<DataColumn>().Select(column => rowView.Row[column]));
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private static string BuildSelectedSqlQueryClipboardText(DataGrid grid)
        {
            var selectedCells = grid.SelectedCells
                .Where(cell => cell.Item is DataRowView && cell.Column != null)
                .ToList();
            if (selectedCells.Count == 0)
            {
                return string.Empty;
            }

            var rowOrder = grid.Items
                .OfType<DataRowView>()
                .Select((row, index) => new { row, index })
                .ToDictionary(item => item.row, item => item.index);
            var columnOrder = grid.Columns
                .Select((column, index) => new { column, index })
                .ToDictionary(item => item.column, item => item.index);
            var selectedColumns = selectedCells
                .Select(cell => cell.Column)
                .Distinct()
                .OrderBy(column => columnOrder.ContainsKey(column) ? columnOrder[column] : int.MaxValue)
                .ToList();
            var selectedRows = selectedCells
                .Select(cell => (DataRowView)cell.Item)
                .Distinct()
                .OrderBy(row => rowOrder.ContainsKey(row) ? rowOrder[row] : int.MaxValue)
                .ToList();
            var selectedSet = new HashSet<string>(
                selectedCells.Select(cell => GetSqlCellKey((DataRowView)cell.Item, cell.Column, rowOrder, columnOrder)));

            var builder = new StringBuilder();
            var firstSelectedRow = selectedRows.FirstOrDefault();
            var selectedColumnMaps = selectedColumns
                .Select(column => new { GridColumn = column, DataColumn = GetSqlDataColumnFromGridColumn(column, firstSelectedRow) })
                .Where(item => item.DataColumn != null)
                .ToList();

            AppendSqlQueryHeader(builder, selectedColumnMaps.Select(item => item.DataColumn.ColumnName));
            foreach (var row in selectedRows)
            {
                var values = new List<object>();
                foreach (var columnMap in selectedColumnMaps)
                {
                    var dataColumn = GetSqlDataColumnFromGridColumn(columnMap.GridColumn, row);
                    if (dataColumn == null)
                    {
                        values.Add(string.Empty);
                        continue;
                    }

                    values.Add(selectedSet.Contains(GetSqlCellKey(row, columnMap.GridColumn, rowOrder, columnOrder))
                        ? row.Row[dataColumn]
                        : string.Empty);
                }

                AppendSqlQueryRow(builder, values);
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private static DataColumn GetSqlDataColumnFromGridColumn(DataGridColumn column, DataRowView row)
        {
            if (column == null || row?.Row?.Table == null)
            {
                return null;
            }

            if (column.Header is DataColumn dataColumn && row.Row.Table.Columns.Contains(dataColumn.ColumnName))
            {
                return row.Row.Table.Columns[dataColumn.ColumnName];
            }

            var columnName = GetSqlDataColumnName(column);
            if (string.IsNullOrWhiteSpace(columnName))
            {
                return null;
            }

            return row.Row.Table.Columns.Contains(columnName)
                ? row.Row.Table.Columns[columnName]
                : null;
        }

        private static string GetSqlDataColumnName(DataGridColumn column)
        {
            if (column == null)
            {
                return string.Empty;
            }

            var header = Convert.ToString(column.Header);
            if (!string.IsNullOrWhiteSpace(header))
            {
                return header;
            }

            if (!string.IsNullOrWhiteSpace(column.SortMemberPath))
            {
                return CleanSqlDataColumnPath(column.SortMemberPath);
            }

            var boundColumn = column as DataGridBoundColumn;
            var binding = boundColumn?.Binding as Binding;
            return binding?.Path == null ? string.Empty : CleanSqlDataColumnPath(binding.Path.Path);
        }

        private static string CleanSqlDataColumnPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            var trimmed = path.Trim();
            return trimmed.Length > 2 && trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal)
                ? trimmed.Substring(1, trimmed.Length - 2)
                : trimmed;
        }

        private static string GetSqlCellKey(
            DataRowView row,
            DataGridColumn column,
            IDictionary<DataRowView, int> rowOrder,
            IDictionary<DataGridColumn, int> columnOrder)
        {
            var rowIndex = rowOrder.ContainsKey(row) ? rowOrder[row] : -1;
            var columnIndex = columnOrder.ContainsKey(column) ? columnOrder[column] : -1;
            return rowIndex + ":" + columnIndex;
        }

        private static void AppendSqlQueryHeader(StringBuilder builder, IEnumerable<string> columns)
        {
            builder.AppendLine(string.Join("\t", columns.Select(EscapeClipboardCell)));
        }

        private static void AppendSqlQueryRow(StringBuilder builder, IEnumerable<object> values)
        {
            builder.AppendLine(string.Join("\t", values.Select(FormatSqlClipboardValue)));
        }

        private static string FormatSqlClipboardValue(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            if (value is DateTime dateTime)
            {
                return EscapeClipboardCell(dateTime.ToString("yyyy-MM-dd HH:mm:ss"));
            }

            return EscapeClipboardCell(Convert.ToString(value));
        }

        private static string EscapeClipboardCell(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r\n", " ")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
        }

        private void ViewModel_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is MainViewModel moduleVm && e.PropertyName == nameof(MainViewModel.CurrentModule)
                && moduleVm.CurrentModule != "VideoViewer")
            {
                PauseVideoViewer();
            }

            if (sender is MainViewModel videoVm && e.PropertyName == nameof(MainViewModel.VideoViewerSource))
            {
                try
                {
                    VideoPlayer.Stop();
                }
                catch
                {
                }

                _videoProgressTimer.Stop();
                videoVm.IsVideoViewerPlaying = false;
                videoVm.VideoViewerPositionSeconds = 0;
                videoVm.UpdateVideoViewerSubtitle(0);
                ApplyVideoSpeed(videoVm.VideoViewerSpeedRatio);
            }

            if (sender is MainViewModel speedVm && e.PropertyName == nameof(MainViewModel.VideoViewerSpeedRatio))
            {
                ApplyVideoSpeed(speedVm.VideoViewerSpeedRatio);
            }

            if (e.PropertyName != nameof(MainViewModel.SqlPassword) || !(sender is MainViewModel viewModel))
            {
                return;
            }

            if (SqlPasswordBox.Password != (viewModel.SqlPassword ?? string.Empty))
            {
                SqlPasswordBox.Password = viewModel.SqlPassword ?? string.Empty;
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var handle = new WindowInteropHelper(this).Handle;
            HotkeyService.Initialize(handle);
            _windowSource = HwndSource.FromHwnd(handle);
            _windowSource?.AddHook(WndProc);
            EnableDragDropForElevatedProcess(handle);
            if (DataContext is MainViewModel vm)
                vm.ReRegisterHotkey();
        }

        // ===== UIPI: 允许低权限窗口（如普通资源管理器）向本提升进程拖文件 =====
        private const int WM_DROPFILES = 0x0233;
        private const int WM_COPYDATA = 0x004A;
        private const int WM_COPYGLOBALDATA = 0x0049;
        private const int MSGFLT_ALLOW = 1;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct CHANGEFILTERSTRUCT { public uint cbSize; public uint ExtStatus; }

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool ChangeWindowMessageFilterEx(IntPtr hWnd, uint message, int action, ref CHANGEFILTERSTRUCT changeInfo);

        private static void EnableDragDropForElevatedProcess(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return;
            try
            {
                var cfs = new CHANGEFILTERSTRUCT { cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(CHANGEFILTERSTRUCT)) };
                ChangeWindowMessageFilterEx(hwnd, WM_DROPFILES, MSGFLT_ALLOW, ref cfs);
                ChangeWindowMessageFilterEx(hwnd, WM_COPYDATA, MSGFLT_ALLOW, ref cfs);
                ChangeWindowMessageFilterEx(hwnd, WM_COPYGLOBALDATA, MSGFLT_ALLOW, ref cfs);
            }
            catch (Exception ex)
            {
                AppLogService.Warning("ChangeWindowMessageFilterEx failed: {Msg}", ex.Message);
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY && (int)wParam == HotkeyService.ScreenshotHotkeyId)
            {
                if (DataContext is MainViewModel vm)
                    _ = vm.TriggerScreenshotAsync();
                handled = true;
            }
            return IntPtr.Zero;
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (TryHandleVideoViewerShortcut(e))
            {
                return;
            }

            if (!(DataContext is MainViewModel vm) || !vm.IsCapturingHotkey) return;

            var key = ResolveHotkeyKey(e);

            if (key == Key.LeftShift || key == Key.RightShift ||
                key == Key.LeftCtrl  || key == Key.RightCtrl  ||
                key == Key.LeftAlt   || key == Key.RightAlt   ||
                key == Key.LWin      || key == Key.RWin       ||
                key == Key.Escape)
            {
                if (key == Key.Escape)
                {
                    vm.IsCapturingHotkey = false;
                    e.Handled = true;
                }

                return;
            }

            var modifiers = Keyboard.Modifiers;
            uint fsModifiers = 0;
            if ((modifiers & ModifierKeys.Control) != 0) fsModifiers |= 0x0002;
            if ((modifiers & ModifierKeys.Shift)   != 0) fsModifiers |= 0x0004;
            if ((modifiers & ModifierKeys.Alt)     != 0) fsModifiers |= 0x0001;
            if ((modifiers & ModifierKeys.Windows) != 0) fsModifiers |= 0x0008;

            if (fsModifiers == 0)
            {
                e.Handled = true;
                return;
            }

            var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            if (vk == 0)
            {
                e.Handled = true;
                return;
            }

            vm.ApplyPendingHotkey(fsModifiers, vk);
            e.Handled = true;
        }

        private bool TryHandleVideoViewerShortcut(KeyEventArgs e)
        {
            if (!(DataContext is MainViewModel vm)
                || vm.CurrentModule != "VideoViewer"
                || !vm.HasVideoViewerVideo
                || Keyboard.Modifiers != ModifierKeys.None)
            {
                return false;
            }

            switch (e.Key)
            {
                case Key.Space:
                    VideoPlayPauseButton_OnClick(this, new RoutedEventArgs());
                    e.Handled = true;
                    return true;
                case Key.Left:
                    var previousSeconds = Math.Max(0, vm.VideoViewerPositionSeconds - 5);
                    SeekVideoViewer(previousSeconds);
                    vm.VideoViewerPositionSeconds = previousSeconds;
                    vm.UpdateVideoViewerSubtitle(previousSeconds);
                    e.Handled = true;
                    return true;
                case Key.Right:
                    var nextSeconds = vm.VideoViewerDurationSeconds > 0
                        ? Math.Min(vm.VideoViewerDurationSeconds, vm.VideoViewerPositionSeconds + 5)
                        : vm.VideoViewerPositionSeconds + 5;
                    SeekVideoViewer(nextSeconds);
                    vm.VideoViewerPositionSeconds = nextSeconds;
                    vm.UpdateVideoViewerSubtitle(nextSeconds);
                    e.Handled = true;
                    return true;
                case Key.Up:
                    vm.VideoViewerVolume = Math.Min(1.0, vm.VideoViewerVolume + 0.05);
                    e.Handled = true;
                    return true;
                case Key.Down:
                    vm.VideoViewerVolume = Math.Max(0.0, vm.VideoViewerVolume - 0.05);
                    e.Handled = true;
                    return true;
                default:
                    return false;
            }
        }

        private static Key ResolveHotkeyKey(KeyEventArgs e)
        {
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (key == Key.ImeProcessed && e.ImeProcessedKey != Key.None)
            {
                key = e.ImeProcessedKey;
            }
            else if (key == Key.DeadCharProcessed && e.DeadCharProcessedKey != Key.None)
            {
                key = e.DeadCharProcessedKey;
            }

            return key;
        }

        public void OpenMediaFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            if (!MediaFileAssociationCore.IsSupportedMediaExtension(Path.GetExtension(filePath)))
            {
                return;
            }

            if (DataContext is MainViewModel vm && vm.TryOpenVideoViewerFile(filePath))
            {
                Show();
                if (WindowState == WindowState.Minimized)
                {
                    WindowState = WindowState.Normal;
                }

                Activate();
                Focus();
            }
        }

        private void ImageFolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is ImageFolderNode node && DataContext is MainViewModel vm)
                vm.OnImageFolderTreeSelected(node);
        }

        private void Window_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (DataContext is MainViewModel videoVm
                && videoVm.CurrentModule == "VideoViewer"
                && TryGetDroppedVideoFile(e, out _))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }

            if (DataContext is MainViewModel imageVm
                && imageVm.CurrentModule == "ImageViewer"
                && TryGetDroppedImageFile(e, out _))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }

            // 仅在 CodexConfig 模块且确实是配置文件夹时拦截；其它模块继续让子元素处理拖放。
            if (DataContext is MainViewModel vm && vm.CurrentModule == "CodexConfig"
                && TryGetDroppedFolders(e, out _))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (DataContext is MainViewModel videoVm
                && videoVm.CurrentModule == "VideoViewer"
                && TryGetDroppedVideoFiles(e, out var videoPaths))
            {
                videoVm.TryOpenVideoViewerFiles(videoPaths);
                VideoPlayer.Stop();
                videoVm.IsVideoViewerPlaying = false;
                e.Handled = true;
                return;
            }

            if (DataContext is MainViewModel imageVm
                && imageVm.CurrentModule == "ImageViewer"
                && TryGetDroppedImageFile(e, out var imagePath))
            {
                imageVm.TryOpenImageViewerFile(imagePath);
                e.Handled = true;
                return;
            }

            if (DataContext is MainViewModel viewModel && viewModel.CurrentModule == "CodexConfig"
                && TryGetDroppedFolders(e, out var folders))
            {
                await viewModel.AddCodexProfileFoldersAsync(folders);
                e.Handled = true;
            }
            // 其它情况让事件继续传递，FileVerify_OnDrop 等子处理器才能收到。
        }

        /// <summary>
        /// 同时支持文件夹拖入与文件（config.toml / auth.json）拖入。
        /// 文件会映射为其父目录，再由 ViewModel 校验该目录是否同时包含两个必需文件。
        /// </summary>
        private static bool TryGetDroppedFolders(DragEventArgs e, out string[] folderPaths)
        {
            folderPaths = new string[0];
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return false;
            }

            var droppedPaths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (droppedPaths == null || droppedPaths.Length == 0)
            {
                return false;
            }

            var collected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in droppedPaths)
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                if (Directory.Exists(path))
                {
                    collected.Add(path);
                }
                else if (File.Exists(path))
                {
                    var name = Path.GetFileName(path);
                    if (string.Equals(name, "config.toml", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(name, "auth.json", StringComparison.OrdinalIgnoreCase))
                    {
                        var parent = Path.GetDirectoryName(path);
                        if (!string.IsNullOrWhiteSpace(parent) && Directory.Exists(parent))
                        {
                            collected.Add(parent);
                        }
                    }
                }
            }

            folderPaths = collected.ToArray();
            return folderPaths.Length > 0;
        }

        private static bool TryGetDroppedImageFile(DragEventArgs e, out string imagePath)
        {
            imagePath = null;
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return false;
            }

            var droppedPaths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (droppedPaths == null || droppedPaths.Length == 0)
            {
                return false;
            }

            foreach (var path in droppedPaths)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    continue;
                }

                switch ((Path.GetExtension(path) ?? string.Empty).ToLowerInvariant())
                {
                    case ".png":
                    case ".jpg":
                    case ".jpeg":
                    case ".bmp":
                    case ".gif":
                    case ".tif":
                    case ".tiff":
                        imagePath = path;
                        return true;
                }
            }

            return false;
        }

        private static bool TryGetDroppedVideoFile(DragEventArgs e, out string videoPath)
        {
            videoPath = null;
            if (TryGetDroppedVideoFiles(e, out var videoPaths) && videoPaths.Length > 0)
            {
                videoPath = videoPaths[0];
                return true;
            }

            return false;
        }

        private static bool TryGetDroppedVideoFiles(DragEventArgs e, out string[] videoPaths)
        {
            videoPaths = new string[0];
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return false;
            }

            var droppedPaths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (droppedPaths == null || droppedPaths.Length == 0)
            {
                return false;
            }

            videoPaths = droppedPaths
                .Where(path => !string.IsNullOrWhiteSpace(path)
                    && File.Exists(path)
                    && MediaFileAssociationCore.IsSupportedMediaExtension(Path.GetExtension(path)))
                .ToArray();
            return videoPaths.Length > 0;
        }

        private void VideoPlayer_MediaOpened(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                vm.VideoViewerDurationSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            }

            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
            vm.VideoViewerStatusMessage = "媒体已就绪。";
        }

        private void VideoPlayer_MediaEnded(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            var endSeconds = vm.VideoViewerDurationSeconds > 0
                ? vm.VideoViewerDurationSeconds
                : VideoPlayer.Position.TotalSeconds;
            if (vm.ShouldLoopVideoViewerAt(endSeconds, out var loopTargetSeconds))
            {
                SeekVideoViewer(loopTargetSeconds);
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
                VideoPlayer.Play();
                _videoProgressTimer.Start();
                vm.IsVideoViewerPlaying = true;
                vm.VideoViewerPositionSeconds = loopTargetSeconds;
                vm.UpdateVideoViewerSubtitle(loopTargetSeconds);
                vm.VideoViewerStatusMessage = "正在 A/B 循环。";
                return;
            }

            VideoPlayer.Stop();
            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            if (vm.OpenNextVideoViewerPlaylistItem())
            {
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
                VideoPlayer.Play();
                _videoProgressTimer.Start();
                vm.IsVideoViewerPlaying = true;
                vm.VideoViewerStatusMessage = vm.VideoViewerAutoAdvanceText;
            }
            else
            {
                vm.VideoViewerStatusMessage = "播放结束。";
            }
        }

        private void VideoPlayer_MediaFailed(object sender, ExceptionRoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            var message = e.ErrorException?.Message ?? "系统缺少对应解码器";
            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerStatusMessage = "播放失败：" + message;
            MessageBox.Show(
                "播放失败：" + message + "\n\n可点击“外部打开”用系统默认播放器尝试播放。",
                "音视频播放",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void VideoPlayPauseButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            if (vm.IsVideoViewerPlaying)
            {
                PauseVideoViewer();
                vm.VideoViewerStatusMessage = "已暂停。";
            }
            else
            {
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
                VideoPlayer.Play();
                _videoProgressTimer.Start();
                vm.IsVideoViewerPlaying = true;
                vm.VideoViewerStatusMessage = "正在播放。";
            }
        }

        private void VideoStopButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            VideoPlayer.Stop();
            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            vm.VideoViewerStatusMessage = "已停止。";
        }

        private void VideoProgressSlider_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isDraggingVideoProgress = true;
        }

        private void VideoProgressSlider_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                _isDraggingVideoProgress = false;
                return;
            }

            _isDraggingVideoProgress = false;
            SeekVideoViewer(VideoProgressSlider.Value);
            vm.VideoViewerPositionSeconds = VideoProgressSlider.Value;
            vm.UpdateVideoViewerSubtitle(VideoProgressSlider.Value);
        }

        private void VideoWaveform_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is UIElement element))
            {
                return;
            }

            element.CaptureMouse();
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                BeginVideoWaveformRangeSelection(sender, e.GetPosition(element));
            }
            else
            {
                _isDraggingVideoWaveform = true;
                SeekVideoViewerByWaveformPoint(sender, e.GetPosition(element), "正在按波形定位：");
            }

            e.Handled = true;
        }

        private void VideoWaveform_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is UIElement element))
            {
                return;
            }

            element.CaptureMouse();
            BeginVideoWaveformRangeSelection(sender, e.GetPosition(element));
            e.Handled = true;
        }

        private void VideoWaveform_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (_isSelectingVideoWaveformRange)
            {
                if (e.LeftButton != MouseButtonState.Pressed && e.RightButton != MouseButtonState.Pressed)
                {
                    EndVideoWaveformInteraction(sender);
                    return;
                }

                if (sender is UIElement rangeElement)
                {
                    UpdateVideoWaveformRangeSelection(sender, e.GetPosition(rangeElement), false);
                }

                e.Handled = true;
                return;
            }

            if (!_isDraggingVideoWaveform)
            {
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                EndVideoWaveformInteraction(sender);
                return;
            }

            if (!(sender is UIElement element))
            {
                return;
            }

            SeekVideoViewerByWaveformPoint(sender, e.GetPosition(element), "正在按波形定位：");
            e.Handled = true;
        }

        private void VideoWaveform_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSelectingVideoWaveformRange)
            {
                if (sender is UIElement rangeElement)
                {
                    UpdateVideoWaveformRangeSelection(sender, e.GetPosition(rangeElement), true);
                }

                EndVideoWaveformInteraction(sender);
                e.Handled = true;
                return;
            }

            if (sender is UIElement element)
            {
                SeekVideoViewerByWaveformPoint(sender, e.GetPosition(element), "已按波形跳转：");
            }

            EndVideoWaveformInteraction(sender);
            e.Handled = true;
        }

        private void VideoWaveform_OnMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSelectingVideoWaveformRange)
            {
                if (sender is UIElement element)
                {
                    UpdateVideoWaveformRangeSelection(sender, e.GetPosition(element), true);
                }

                EndVideoWaveformInteraction(sender);
            }

            e.Handled = true;
        }

        private void VideoWaveformStartHandle_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginVideoWaveformHandleDrag(sender, true);
            e.Handled = true;
        }

        private void VideoWaveformEndHandle_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginVideoWaveformHandleDrag(sender, false);
            e.Handled = true;
        }

        private void VideoWaveformHandle_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingVideoWaveformRangeHandle)
            {
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                EndVideoWaveformInteraction(sender);
                return;
            }

            UpdateVideoWaveformHandleDrag(e.GetPosition(VideoWaveformImage), false);
            e.Handled = true;
        }

        private void VideoWaveformHandle_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDraggingVideoWaveformRangeHandle)
            {
                UpdateVideoWaveformHandleDrag(e.GetPosition(VideoWaveformImage), true);
            }

            EndVideoWaveformInteraction(sender);
            e.Handled = true;
        }

        private void SeekVideoViewerByWaveformPoint(object sender, Point point, string statusPrefix)
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo || vm.VideoViewerDurationSeconds <= 0)
            {
                return;
            }

            if (!(sender is FrameworkElement element) || element.ActualWidth <= 1)
            {
                return;
            }

            var ratio = Math.Max(0, Math.Min(1, point.X / element.ActualWidth));
            var seconds = ratio * vm.VideoViewerDurationSeconds;
            SeekVideoViewer(seconds);
            vm.VideoViewerPositionSeconds = seconds;
            vm.UpdateVideoViewerSubtitle(seconds);
            vm.VideoViewerStatusMessage = (statusPrefix ?? "已按波形跳转：") + vm.VideoViewerPositionText;
        }

        private void BeginVideoWaveformRangeSelection(object sender, Point point)
        {
            _isDraggingVideoWaveform = false;
            if (!TryGetVideoWaveformSeconds(sender, point, out var seconds))
            {
                _isSelectingVideoWaveformRange = false;
                return;
            }

            _isSelectingVideoWaveformRange = true;
            _videoWaveformRangeStartSeconds = seconds;
            if (DataContext is MainViewModel vm)
            {
                var previewEndSeconds = Math.Min(vm.VideoViewerDurationSeconds, seconds + 0.25);
                if (previewEndSeconds <= seconds + 0.2)
                {
                    previewEndSeconds = Math.Max(0, seconds - 0.25);
                }

                vm.SetVideoViewerLoopRange(seconds, previewEndSeconds, false);
                vm.VideoViewerStatusMessage = "正在选择 A/B 区间：" + vm.VideoViewerLoopRangeText;
            }
        }

        private void UpdateVideoWaveformRangeSelection(object sender, Point point, bool enableLoop)
        {
            if (!_isSelectingVideoWaveformRange || !TryGetVideoWaveformSeconds(sender, point, out var seconds))
            {
                return;
            }

            if (DataContext is MainViewModel vm)
            {
                vm.SetVideoViewerLoopRange(_videoWaveformRangeStartSeconds, seconds, enableLoop);
                if (!enableLoop && vm.HasVideoViewerLoopRange)
                {
                    vm.VideoViewerStatusMessage = "正在选择 A/B 区间：" + vm.VideoViewerLoopRangeText;
                }
            }
        }

        private void BeginVideoWaveformHandleDrag(object sender, bool isStartHandle)
        {
            _isDraggingVideoWaveform = false;
            _isSelectingVideoWaveformRange = false;
            _isDraggingVideoWaveformRangeHandle = true;
            _isDraggingVideoWaveformStartHandle = isStartHandle;
            if (sender is UIElement element)
            {
                element.CaptureMouse();
            }

            if (DataContext is MainViewModel vm)
            {
                vm.VideoViewerStatusMessage = isStartHandle
                    ? "正在微调 A 点：" + vm.VideoViewerLoopStartText
                    : "正在微调 B 点：" + vm.VideoViewerLoopEndText;
            }
        }

        private void UpdateVideoWaveformHandleDrag(Point point, bool finalize)
        {
            if (!_isDraggingVideoWaveformRangeHandle || !TryGetVideoWaveformSeconds(VideoWaveformImage, point, out var seconds))
            {
                return;
            }

            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerLoopRange)
            {
                return;
            }

            seconds = _isDraggingVideoWaveformStartHandle
                ? Math.Min(seconds, Math.Max(0, vm.VideoViewerLoopEndSeconds - 0.25))
                : Math.Max(seconds, vm.VideoViewerLoopStartSeconds + 0.25);
            var startSeconds = _isDraggingVideoWaveformStartHandle ? seconds : vm.VideoViewerLoopStartSeconds;
            var endSeconds = _isDraggingVideoWaveformStartHandle ? vm.VideoViewerLoopEndSeconds : seconds;
            if (vm.SetVideoViewerLoopRange(startSeconds, endSeconds, vm.VideoViewerIsLoopEnabled))
            {
                vm.VideoViewerStatusMessage = finalize
                    ? "已微调 A/B 循环：" + vm.VideoViewerLoopRangeText
                    : (_isDraggingVideoWaveformStartHandle
                        ? "正在微调 A 点：" + vm.VideoViewerLoopStartText
                        : "正在微调 B 点：" + vm.VideoViewerLoopEndText);
            }
        }

        private bool TryGetVideoWaveformSeconds(object sender, Point point, out double seconds)
        {
            seconds = 0;
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo || vm.VideoViewerDurationSeconds <= 0)
            {
                return false;
            }

            if (!(sender is FrameworkElement element) || element.ActualWidth <= 1)
            {
                return false;
            }

            var ratio = Math.Max(0, Math.Min(1, point.X / element.ActualWidth));
            seconds = ratio * vm.VideoViewerDurationSeconds;
            return true;
        }

        private void EndVideoWaveformInteraction(object sender)
        {
            _isDraggingVideoWaveform = false;
            _isSelectingVideoWaveformRange = false;
            _isDraggingVideoWaveformRangeHandle = false;
            if (sender is UIElement element)
            {
                element.ReleaseMouseCapture();
            }
        }

        private void VideoProgressTimer_OnTick(object sender, EventArgs e)
        {
            if (_isDraggingVideoProgress || _isDraggingVideoWaveform || _isSelectingVideoWaveformRange || _isDraggingVideoWaveformRangeHandle || !(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            var positionSeconds = VideoPlayer.Position.TotalSeconds;
            vm.VideoViewerPositionSeconds = positionSeconds;
            vm.UpdateVideoViewerSubtitle(positionSeconds);
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                vm.VideoViewerDurationSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            }

            if (vm.ShouldLoopVideoViewerAt(positionSeconds, out var loopTargetSeconds))
            {
                SeekVideoViewer(loopTargetSeconds);
                vm.VideoViewerPositionSeconds = loopTargetSeconds;
                vm.UpdateVideoViewerSubtitle(loopTargetSeconds);
                return;
            }
        }

        private void PauseVideoViewer()
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            try
            {
                VideoPlayer.Pause();
            }
            catch
            {
            }

            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
        }

        private void VideoSpeedComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is MainViewModel vm)
            {
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
            }
        }

        private void ApplyVideoSpeed(double speedRatio)
        {
            if (double.IsNaN(speedRatio) || double.IsInfinity(speedRatio) || speedRatio <= 0)
            {
                speedRatio = 1.0;
            }

            try
            {
                VideoPlayer.SpeedRatio = speedRatio;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Video speed change failed: {Msg}", ex.Message);
            }
        }

        private void SeekVideoViewer(double seconds)
        {
            if (double.IsNaN(seconds) || double.IsInfinity(seconds) || seconds < 0)
            {
                return;
            }

            try
            {
                VideoPlayer.Position = TimeSpan.FromSeconds(seconds);
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Video seek failed: {Msg}", ex.Message);
            }
        }

        private void VideoPlaylistListBox_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _videoPlaylistDragStartPoint = e.GetPosition(null);
            _videoPlaylistDragItem = FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject)?.DataContext as VideoPlaylistItem;
        }

        private void VideoPlaylistListBox_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _videoPlaylistDragItem == null)
            {
                return;
            }

            var currentPosition = e.GetPosition(null);
            if (Math.Abs(currentPosition.X - _videoPlaylistDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(currentPosition.Y - _videoPlaylistDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            DragDrop.DoDragDrop(VideoPlaylistListBox, _videoPlaylistDragItem, DragDropEffects.Move);
            _videoPlaylistDragItem = null;
        }

        private void VideoPlaylistListBox_OnDrop(object sender, DragEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            var item = e.Data.GetData(typeof(VideoPlaylistItem)) as VideoPlaylistItem;
            var target = FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject)?.DataContext as VideoPlaylistItem;
            if (item == null || target == null)
            {
                return;
            }

            vm.MoveVideoViewerPlaylistItem(item, target);
            e.Handled = true;
        }

        private void VideoPlaylistItem_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ListBoxItem item)
            {
                vm.OpenVideoViewerPlaylistItemCommand.Execute(item.DataContext);
                e.Handled = true;
            }
        }

        private void VideoLoopRangeItem_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ListBoxItem item && item.DataContext is VideoLoopRangeItem range)
            {
                SeekVideoViewer(range.StartSeconds);
                vm.OpenVideoViewerLoopRangeCommand.Execute(range);
                e.Handled = true;
            }
        }

        private void ConvertQueueGrid_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _convertQueueDragStartPoint = e.GetPosition(null);
            _convertQueueDragItem = FindAncestor<DataGridRow>(e.OriginalSource as DependencyObject)?.DataContext as ConvertQueueItem;
        }

        private void ConvertQueueGrid_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _convertQueueDragItem == null)
            {
                return;
            }

            if (DataContext is MainViewModel vm && vm.IsConvertBusy)
            {
                return;
            }

            var currentPosition = e.GetPosition(null);
            if (Math.Abs(currentPosition.X - _convertQueueDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(currentPosition.Y - _convertQueueDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            DragDrop.DoDragDrop(ConvertQueueGrid, _convertQueueDragItem, DragDropEffects.Move);
            _convertQueueDragItem = null;
        }

        private void ConvertQueueGrid_OnDrop(object sender, DragEventArgs e)
        {
            if (!(DataContext is MainViewModel vm) || vm.IsConvertBusy)
            {
                return;
            }

            var item = e.Data.GetData(typeof(ConvertQueueItem)) as ConvertQueueItem;
            var target = FindAncestor<DataGridRow>(e.OriginalSource as DependencyObject)?.DataContext as ConvertQueueItem;
            if (item == null || target == null || ReferenceEquals(item, target))
            {
                return;
            }

            vm.MoveConvertQueueItem(item, target);
            e.Handled = true;
        }

        private void VideoLoopRangeOpenButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is VideoLoopRangeItem range)
            {
                SeekVideoViewer(range.StartSeconds);
            }
        }

        private void VideoBookmarkItem_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ListBoxItem item && item.DataContext is VideoBookmarkItem bookmark)
            {
                SeekVideoViewer(bookmark.PositionSeconds);
                vm.OpenVideoViewerBookmarkCommand.Execute(bookmark);
                e.Handled = true;
            }
        }

        private void VideoBookmarkJumpButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is VideoBookmarkItem bookmark)
            {
                SeekVideoViewer(bookmark.PositionSeconds);
            }
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T match)
                {
                    return match;
                }

                current = System.Windows.Media.VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        protected override void OnClosed(EventArgs e)
        {
            _videoProgressTimer.Stop();
            try { VideoPlayer.Stop(); } catch { }
            _windowSource?.RemoveHook(WndProc);
            _windowSource = null;
            HotkeyService.Unregister();
            TrayIcon.Dispose();
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.PropertyChanged -= ViewModel_OnPropertyChanged;
                viewModel.Dispose();
            }

            base.OnClosed(e);
        }

        private void FileHashResult_OnClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && !string.IsNullOrWhiteSpace(vm.FileHashResult))
            {
                try
                {
                    Clipboard.SetText(vm.FileHashResult);
                    vm.FileHashStatusMessage = "已复制到剪贴板。";
                }
                catch { }
            }
        }

        private void FileVerify_OnDragOver(object sender, System.Windows.DragEventArgs e)
        {
            bool ok = e.Data != null && (
                e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop) ||
                e.Data.GetDataPresent(System.Windows.DataFormats.UnicodeText) ||
                e.Data.GetDataPresent(System.Windows.DataFormats.Text));
            e.Effects = ok ? System.Windows.DragDropEffects.Copy : System.Windows.DragDropEffects.None;
            e.Handled = true;
        }

        private async void FileVerify_OnDrop(object sender, System.Windows.DragEventArgs e)
        {
            e.Handled = true;
            var vm = DataContext as MainViewModel;
            if (vm == null)
            {
                MyTools.Services.AppLogService.Warning("FileVerify Drop: DataContext is not MainViewModel.");
                return;
            }

            try
            {
                // 列出所有可用格式以便诊断 UIPI/OLE 兼容性问题。
                var fmts = e.Data?.GetFormats() ?? new string[0];
                MyTools.Services.AppLogService.Information("FileVerify Drop fired. formats=[{Fmts}]", string.Join(",", fmts));

                var paths = new List<string>();

                // 1) 标准 FileDrop
                if (e.Data != null && e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
                {
                    var files = e.Data.GetData(System.Windows.DataFormats.FileDrop) as string[];
                    if (files != null && files.Length > 0)
                    {
                        paths.AddRange(files.Where(System.IO.File.Exists));
                    }
                }

                // 2) 兜底：UIPI 中 OLE 可能只传 Text/Unicode 路径字符串。
                if (paths.Count == 0 && e.Data != null)
                {
                    foreach (var fmt in new[] { System.Windows.DataFormats.UnicodeText, System.Windows.DataFormats.Text })
                    {
                        if (!e.Data.GetDataPresent(fmt)) continue;
                        var s = e.Data.GetData(fmt) as string;
                        if (string.IsNullOrWhiteSpace(s)) continue;
                        var first = s.Split(new[] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries)[0].Trim('"', ' ');
                        if (System.IO.File.Exists(first)) { paths.Add(first); break; }
                    }
                }

                paths = paths
                    .Where(System.IO.File.Exists)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (paths.Count == 0)
                {
                    vm.FileHashStatusMessage = "未取到可用文件路径（格式：" + string.Join(",", fmts) + "）。";
                    MyTools.Services.AppLogService.Warning("FileVerify Drop: no usable file path. formats=[{Fmts}]", string.Join(",", fmts));
                    return;
                }

                vm.FileHashStatusMessage = paths.Count == 1
                    ? "已接收：" + System.IO.Path.GetFileName(paths[0])
                    : $"已接收 {paths.Count} 个文件。";
                await vm.VerifyFromPathsAsync(paths);
            }
            catch (System.Exception ex)
            {
                MyTools.Services.AppLogService.Error(ex, "FileVerify Drop failed");
                vm.FileHashStatusMessage = "拖放校验失败：" + ex.Message;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!App.IsExiting)
            {
                e.Cancel = true;
                Hide();
            }

            base.OnClosing(e);
        }
    }
}
