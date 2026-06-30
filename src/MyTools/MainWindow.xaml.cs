using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using Hardcodet.Wpf.TaskbarNotification;
using MyTools.Services;
using MyTools.Shared;
using MyTools.ViewModels;
using MyTools.Views;

namespace MyTools
{
    public partial class MainWindow : Window
    {
        private const int WM_HOTKEY = 0x0312;
        private HwndSource _windowSource;
        private Point _convertQueueDragStartPoint;
        private ConvertQueueItem _convertQueueDragItem;
        private IDisposable _trayIcon;
        private DispatcherTimer _trayTimer;

        public long InitializeComponentElapsedMilliseconds { get; private set; }

        public MainWindow()
        {
            var initializeStopwatch = System.Diagnostics.Stopwatch.StartNew();
            InitializeComponent();
            initializeStopwatch.Stop();
            InitializeComponentElapsedMilliseconds = initializeStopwatch.ElapsedMilliseconds;
            AppLogService.InformationIfInitialized("MainWindow InitializeComponent completed in {ElapsedMs} ms", InitializeComponentElapsedMilliseconds);
            Loaded += (s, e) => ScheduleDeferredStartupUi();
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.PropertyChanged += ViewModel_OnPropertyChanged;
            }
        }

        private void MenuAbout_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "阿君的工具\nMyTools - 个人实用工具集\n基于 .NET Framework 4.8 + WPF",
                "关于",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void MenuOpenAgents_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var candidates = new[]
                {
                    Path.Combine(baseDirectory, "AGENTS.md"),
                    Path.GetFullPath(Path.Combine(baseDirectory, @"..\..\..\..\..\AGENTS.md")),
                    Path.Combine(Directory.GetCurrentDirectory(), "AGENTS.md")
                };

                var path = candidates
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault(File.Exists);
                if (path == null)
                {
                    return;
                }

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Open AGENTS.md failed: {Msg}", ex.Message);
            }
        }

        private void InitializeTrayIcon()
        {
            if (_trayIcon != null)
            {
                return;
            }

            var trayIcon = new TaskbarIcon
            {
                ToolTipText = "阿君的工具（双击恢复窗口）",
                Visibility = Visibility.Visible,
                MenuActivation = PopupActivationMode.LeftOrRightClick,
                ContextMenu = CreateTrayContextMenu()
            };

            if (DataContext is MainViewModel viewModel)
            {
                trayIcon.DoubleClickCommand = viewModel.RestoreCommand;
            }

            var iconInfo = Application.GetResourceStream(new Uri("pack://application:,,,/Resources/AppIcon.ico", UriKind.Absolute));
            if (iconInfo?.Stream != null)
            {
                using (var icon = new System.Drawing.Icon(iconInfo.Stream))
                {
                    trayIcon.Icon = (System.Drawing.Icon)icon.Clone();
                }
            }

            _trayIcon = trayIcon;
        }

        private void ScheduleDeferredStartupUi()
        {
            _trayTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(10)
            };
            _trayTimer.Tick += (sender, args) =>
            {
                _trayTimer.Stop();
                CosturaBootstrap.EnsureInitialized();
                InitializeTrayIcon();
            };
            _trayTimer.Start();
        }

        private void EnsureDeferredApplicationResourcesLoaded()
        {
            if (DeferredUiResourceService.IsLoaded)
            {
                return;
            }

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                DeferredUiResourceService.EnsureLoaded();
                AppLogService.InformationIfInitialized("Deferred UI resources loaded in {ElapsedMs} ms", stopwatch.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading deferred UI resources failed.");
            }
        }

        private ContextMenu CreateTrayContextMenu()
        {
            var menu = new ContextMenu { DataContext = DataContext };
            if (!(DataContext is MainViewModel viewModel))
            {
                return menu;
            }

            menu.Items.Add(new MenuItem { Header = "显示窗口", Command = viewModel.RestoreCommand });

            var autoStartItem = new MenuItem { Command = viewModel.ToggleAutoStartCommand };
            var autoStartStyle = new Style(typeof(MenuItem));
            autoStartStyle.Setters.Add(new Setter(HeaderedItemsControl.HeaderProperty, "开启 开机自启"));
            var autoStartTrigger = new DataTrigger
            {
                Binding = new Binding(nameof(MainViewModel.IsAutoStartEnabled)),
                Value = true
            };
            autoStartTrigger.Setters.Add(new Setter(HeaderedItemsControl.HeaderProperty, "关闭 开机自启"));
            autoStartStyle.Triggers.Add(autoStartTrigger);
            autoStartItem.Style = autoStartStyle;
            menu.Items.Add(autoStartItem);

            menu.Items.Add(new MenuItem { Header = "打开日志目录", Command = viewModel.OpenLogFolderCommand });
            menu.Items.Add(new Separator());
            menu.Items.Add(new MenuItem { Header = "退出程序", Command = viewModel.ExitCommand });
            return menu;
        }

        private void LoadHomeContent()
        {
            var template = TryFindResource("HomeViewTemplate") as DataTemplate;
            if (template != null && !ReferenceEquals(HomeHost.ContentTemplate, template))
            {
                HomeHost.ContentTemplate = template;
            }
        }

        private void ImageAdjustmentSlider_OnCommit(object sender, RoutedEventArgs e)
        {
            if (sender is Slider slider)
            {
                slider.GetBindingExpression(Slider.ValueProperty)?.UpdateSource();
            }
        }

        private void ViewModel_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is MainViewModel navigationVm && e.PropertyName == nameof(MainViewModel.CurrentModule))
            {
                if (string.Equals(navigationVm.CurrentModule, "Home", StringComparison.Ordinal))
                {
                    if (DeferredUiResourceService.IsLoaded)
                    {
                        LoadHomeContent();
                    }
                }
                else
                {
                    EnsureDeferredApplicationResourcesLoaded();
                }
            }

            if (sender is MainViewModel moduleVm && e.PropertyName == nameof(MainViewModel.CurrentModule)
                && moduleVm.CurrentModule != "VideoViewer")
            {
                CurrentVideoViewerPage?.PauseVideoViewer();
            }

            if (sender is MainViewModel && e.PropertyName == nameof(MainViewModel.VideoViewerSource))
            {
                CurrentVideoViewerPage?.ResetForSourceChange();
            }

            if (sender is MainViewModel speedVm && e.PropertyName == nameof(MainViewModel.VideoViewerSpeedRatio))
            {
                CurrentVideoViewerPage?.ApplyVideoSpeed(speedVm.VideoViewerSpeedRatio);
            }        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var handle = new WindowInteropHelper(this).Handle;
            HotkeyService.Initialize(handle);
            _windowSource = HwndSource.FromHwnd(handle);
            _windowSource?.AddHook(WndProc);
            EnableDragDropForElevatedProcess(handle);
            if (DataContext is MainViewModel vm)
            {
                vm.ReRegisterHotkey();
                vm.ScheduleStartupHotkeyRegistration();
            }
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
            if (msg == WM_HOTKEY)
            {
                var id = (int)wParam;
                if (id == HotkeyService.ScreenshotHotkeyId)
                {
                    if (DataContext is MainViewModel vm) _ = vm.TriggerScreenshotAsync();
                    handled = true;
                }
                else if (id == HotkeyService.VideoRecordHotkeyId)
                {
                    if (DataContext is MainViewModel vm && vm.StartVideoRecordingCommand.CanExecute(null))
                        vm.StartVideoRecordingCommand.Execute(null);
                    handled = true;
                }
                else if (id == HotkeyService.AudioRecordHotkeyId)
                {
                    if (DataContext is MainViewModel vm && vm.ToggleAudioRecordingCommand.CanExecute(null))
                        vm.ToggleAudioRecordingCommand.Execute(null);
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (TryHandleVideoViewerShortcut(e))
            {
                return;
            }

            if (!(DataContext is MainViewModel vm)) return;
            var isCapturing = vm.IsCapturingHotkey || vm.IsCapturingVideoRecordHotkey || vm.IsCapturingAudioRecordHotkey;
            if (!isCapturing) return;

            var key = ResolveHotkeyKey(e);

            if (key == Key.LeftShift || key == Key.RightShift ||
                key == Key.LeftCtrl  || key == Key.RightCtrl  ||
                key == Key.LeftAlt   || key == Key.RightAlt   ||
                key == Key.LWin      || key == Key.RWin       ||
                key == Key.Escape)
            {
                if (key == Key.Escape)
                {
                    if (vm.IsCapturingVideoRecordHotkey) vm.IsCapturingVideoRecordHotkey = false;
                    else if (vm.IsCapturingAudioRecordHotkey) vm.IsCapturingAudioRecordHotkey = false;
                    else vm.IsCapturingHotkey = false;
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

            if (fsModifiers == 0) { e.Handled = true; return; }

            var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            if (vk == 0) { e.Handled = true; return; }

            if (vm.IsCapturingVideoRecordHotkey) vm.ApplyPendingVideoRecordHotkey(fsModifiers, vk);
            else if (vm.IsCapturingAudioRecordHotkey) vm.ApplyPendingAudioRecordHotkey(fsModifiers, vk);
            else vm.ApplyPendingHotkey(fsModifiers, vk);
            e.Handled = true;
        }

        private bool TryHandleVideoViewerShortcut(KeyEventArgs e)
        {
            return CurrentVideoViewerPage?.HandleShortcut(e) == true;
        }

        private VideoViewerPage CurrentVideoViewerPage => FindVisualChild<VideoViewerPage>(MainModulesHost);
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
                CurrentVideoViewerPage?.StopForDroppedMedia();
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

            var grid = sender as DataGrid ?? FindVisualChildByName<DataGrid>(MainModulesHost, "ConvertQueueGrid");
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

        private static T FindVisualChild<T>(DependencyObject current) where T : DependencyObject
        {
            if (current == null)
            {
                return null;
            }

            var childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(current);
            for (var index = 0; index < childCount; index++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(current, index);
                if (child is T match)
                {
                    return match;
                }

                var descendant = FindVisualChild<T>(child);
                if (descendant != null)
                {
                    return descendant;
                }
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

        protected override void OnClosed(EventArgs e)
        {
            CurrentVideoViewerPage?.StopVideoViewer();
            _trayTimer?.Stop();
            _trayTimer = null;
            _windowSource?.RemoveHook(WndProc);
            _windowSource = null;
            HotkeyService.UnregisterAll();
            _trayIcon?.Dispose();
            _trayIcon = null;
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
