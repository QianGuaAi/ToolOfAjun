using System.Windows;

namespace MyTools.Views
{
    public partial class PasswordInputDialog : Window
    {
        public PasswordInputDialog(string title, string message)
        {
            InitializeComponent();
            Title = string.IsNullOrWhiteSpace(title) ? "输入口令" : title;
            MessageText.Text = message ?? string.Empty;
            Loaded += (s, e) => PasswordBox.Focus();
        }

        public string Password { get; private set; }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Password = PasswordBox.Password ?? string.Empty;
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
