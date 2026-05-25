using System.Windows;
using System.Windows.Controls;
using MyTools.ViewModels;

namespace MyTools.Views
{
    public partial class FrpView : UserControl
    {
        public FrpView()
        {
            InitializeComponent();
        }

        private void FrpTokenBox_OnPasswordChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is FrpViewModel viewModel && sender is PasswordBox box)
            {
                viewModel.FrpToken = box.Password;
                if (FrpTokenVisibleBox.Text != box.Password)
                {
                    FrpTokenVisibleBox.Text = box.Password;
                }
            }
        }

        private void FrpView_OnLoaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is FrpViewModel viewModel && FrpTokenBox.Password != (viewModel.FrpToken ?? string.Empty))
            {
                FrpTokenBox.Password = viewModel.FrpToken ?? string.Empty;
                FrpTokenVisibleBox.Text = viewModel.FrpToken ?? string.Empty;
            }
        }

        private void ShowFrpTokenButton_OnChecked(object sender, RoutedEventArgs e)
        {
            FrpTokenVisibleBox.Text = FrpTokenBox.Password;
            FrpTokenBox.Visibility = Visibility.Collapsed;
            FrpTokenMaskText.Visibility = Visibility.Collapsed;
            FrpTokenVisibleBox.Visibility = Visibility.Visible;
            ShowFrpTokenButton.Content = "隐藏Token";
        }

        private void ShowFrpTokenButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            FrpTokenBox.Password = FrpTokenVisibleBox.Text ?? string.Empty;
            FrpTokenVisibleBox.Visibility = Visibility.Collapsed;
            FrpTokenMaskText.Visibility = Visibility.Visible;
            FrpTokenBox.Visibility = Visibility.Visible;
            ShowFrpTokenButton.Content = "显示Token";
        }
    }
}
