using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using MyTools.ViewModels;

namespace MyTools
{
    public partial class MainWindow : MetroWindow
    {
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
