using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public MainModulesView()
        {
            InitializeComponent();
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
