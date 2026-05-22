using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using MyTools.Services;
using MyTools.ViewModels;

namespace MyTools.Views
{
    public partial class MainModulesView : UserControl
    {
        private Point _convertQueueDragStartPoint;
        private ConvertQueueItem _convertQueueDragItem;
        private MainViewModel _viewModel;

        public MainModulesView()
        {
            InitializeComponent();
            DataContextChanged += MainModulesView_DataContextChanged;
            Unloaded += MainModulesView_Unloaded;
        }

        private PasswordBox CurrentSqlPasswordBox => FindVisualChildByName<PasswordBox>(SqlExportHost, "SqlPasswordBox");

        private DataGrid CurrentSqlQueryResultGrid => FindVisualChildByName<DataGrid>(SqlExportHost, "SqlQueryResultGrid");

        private void MainModulesView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged -= ViewModel_OnPropertyChanged;
            }

            _viewModel = e.NewValue as MainViewModel;
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged += ViewModel_OnPropertyChanged;
                SyncSqlPasswordBox(_viewModel);
            }
        }

        private void MainModulesView_Unloaded(object sender, RoutedEventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged -= ViewModel_OnPropertyChanged;
                _viewModel = null;
            }
        }

        private void ViewModel_OnPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MainViewModel.SqlPassword) && sender is MainViewModel viewModel)
            {
                SyncSqlPasswordBox(viewModel);
            }
        }

        private void SyncSqlPasswordBox(MainViewModel viewModel)
        {
            var passwordBox = CurrentSqlPasswordBox;
            if (passwordBox != null && passwordBox.Password != (viewModel.SqlPassword ?? string.Empty))
            {
                passwordBox.Password = viewModel.SqlPassword ?? string.Empty;
            }
        }

        private void SqlPasswordBox_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel && sender is PasswordBox passwordBox)
            {
                passwordBox.Password = viewModel.SqlPassword ?? string.Empty;
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
                var passwordBox = CurrentSqlPasswordBox;
                if (passwordBox != null)
                {
                    passwordBox.Password = password;
                }
            }
        }

        private void ImageAdjustmentSlider_OnCommit(object sender, RoutedEventArgs e)
        {
            if (sender is Slider slider)
            {
                slider.GetBindingExpression(Slider.ValueProperty)?.UpdateSource();
            }
        }

        private void ImageFolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is ImageFolderNode node && DataContext is MainViewModel vm)
            {
                vm.OnImageFolderTreeSelected(node);
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
                var resultGrid = CurrentSqlQueryResultGrid;
                var text = BuildSqlQueryResultClipboardText(viewModel.SqlQueryResult, resultGrid);
                if (string.IsNullOrWhiteSpace(text))
                {
                    MessageBox.Show("没有可复制的查询结果。", "复制结果", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Clipboard.SetText(text);
                viewModel.QueryStatusMessage = resultGrid != null && resultGrid.SelectedCells.Count > 0
                    ? $"已复制选中结果，共 {resultGrid.SelectedCells.Count} 个单元格。"
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

            var grid = sender as DataGrid ?? FindVisualChildByName<DataGrid>(ConvertHost, "ConvertQueueGrid");
            if (grid == null)
            {
                return;
            }

            DragDrop.DoDragDrop(grid, _convertQueueDragItem, DragDropEffects.Move);
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

        private void FileHashResult_OnClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && !string.IsNullOrWhiteSpace(vm.FileHashResult))
            {
                try
                {
                    Clipboard.SetText(vm.FileHashResult);
                    vm.FileHashStatusMessage = "已复制到剪贴板。";
                }
                catch
                {
                }
            }
        }

        private void FileVerify_OnDragOver(object sender, DragEventArgs e)
        {
            var ok = e.Data != null && (
                e.Data.GetDataPresent(DataFormats.FileDrop) ||
                e.Data.GetDataPresent(DataFormats.UnicodeText) ||
                e.Data.GetDataPresent(DataFormats.Text));
            e.Effects = ok ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private async void FileVerify_OnDrop(object sender, DragEventArgs e)
        {
            e.Handled = true;
            var vm = DataContext as MainViewModel;
            if (vm == null)
            {
                AppLogService.Warning("FileVerify Drop: DataContext is not MainViewModel.");
                return;
            }

            try
            {
                var fmts = e.Data?.GetFormats() ?? new string[0];
                AppLogService.Information("FileVerify Drop fired. formats=[{Fmts}]", string.Join(",", fmts));

                var paths = new List<string>();
                if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    if (files != null && files.Length > 0)
                    {
                        paths.AddRange(files.Where(File.Exists));
                    }
                }

                if (paths.Count == 0 && e.Data != null)
                {
                    foreach (var fmt in new[] { DataFormats.UnicodeText, DataFormats.Text })
                    {
                        if (!e.Data.GetDataPresent(fmt))
                        {
                            continue;
                        }

                        var value = e.Data.GetData(fmt) as string;
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }

                        var first = value.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim('"', ' ');
                        if (File.Exists(first))
                        {
                            paths.Add(first);
                            break;
                        }
                    }
                }

                paths = paths
                    .Where(File.Exists)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (paths.Count == 0)
                {
                    vm.FileHashStatusMessage = "未取到可用文件路径（格式：" + string.Join(",", fmts) + "）。";
                    AppLogService.Warning("FileVerify Drop: no usable file path. formats=[{Fmts}]", string.Join(",", fmts));
                    return;
                }

                vm.FileHashStatusMessage = paths.Count == 1
                    ? "已接收：" + Path.GetFileName(paths[0])
                    : $"已接收 {paths.Count} 个文件。";
                await vm.VerifyFromPathsAsync(paths);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "FileVerify Drop failed");
                vm.FileHashStatusMessage = "拖放校验失败：" + ex.Message;
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

        private static T FindVisualChildByName<T>(DependencyObject current, string name) where T : FrameworkElement
        {
            if (current == null)
            {
                return null;
            }

            var childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(current);
            for (var index = 0; index < childCount; index++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(current, index);
                if (child is T match && string.Equals(match.Name, name, StringComparison.Ordinal))
                {
                    return match;
                }

                var descendant = FindVisualChildByName<T>(child, name);
                if (descendant != null)
                {
                    return descendant;
                }
            }

            return null;
        }
    }
}
