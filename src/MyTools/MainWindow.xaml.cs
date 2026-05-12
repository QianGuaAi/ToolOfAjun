using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using MahApps.Metro.Controls;
using MyTools.Services;
using MyTools.ViewModels;

namespace MyTools
{
    public partial class MainWindow : MetroWindow
    {
        private const int WM_HOTKEY = 0x0312;
        private HwndSource _windowSource;

        public MainWindow()
        {
            InitializeComponent();
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

        private void ViewModel_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
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
            if (DataContext is MainViewModel vm)
                vm.ReRegisterHotkey();
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

        private void Window_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = TryGetDroppedFolders(e, out _) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (TryGetDroppedFolders(e, out var folders) && DataContext is MainViewModel viewModel)
            {
                await viewModel.AddCodexProfileFoldersAsync(folders);
            }

            e.Handled = true;
        }

        /// <summary>
        /// 同时支持文件夹拖入与文件（config.toml / auth.json）拖入。
        /// 文件会被映射为其父目录，再由 ViewModel 校验该目录是否同时包含两个必需文件。
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

        protected override void OnClosed(EventArgs e)
        {
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
