using System;
using System.ComponentModel;
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
            HwndSource.FromHwnd(handle)?.AddHook(WndProc);
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

            var key = e.Key == Key.System ? e.SystemKey : e.Key;

            if (key == Key.LeftShift || key == Key.RightShift ||
                key == Key.LeftCtrl  || key == Key.RightCtrl  ||
                key == Key.LeftAlt   || key == Key.RightAlt   ||
                key == Key.LWin      || key == Key.RWin       ||
                key == Key.Escape)
            {
                if (key == Key.Escape) vm.IsCapturingHotkey = false;
                return;
            }

            var modifiers = Keyboard.Modifiers;
            uint fsModifiers = 0;
            if ((modifiers & ModifierKeys.Control) != 0) fsModifiers |= 0x0002;
            if ((modifiers & ModifierKeys.Shift)   != 0) fsModifiers |= 0x0004;
            if ((modifiers & ModifierKeys.Alt)     != 0) fsModifiers |= 0x0001;

            if (fsModifiers == 0) return;

            var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
            vm.ApplyPendingHotkey(fsModifiers, vk);
            e.Handled = true;
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
