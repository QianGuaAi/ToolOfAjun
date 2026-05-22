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
            }
        }

        private void FrpView_OnLoaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is FrpViewModel viewModel && FrpTokenBox.Password != (viewModel.FrpToken ?? string.Empty))
            {
                FrpTokenBox.Password = viewModel.FrpToken ?? string.Empty;
            }
        }
    }
}
