using System;
using System.Windows;

namespace MyTools.Views
{
    public partial class NewScheduleDialog : Window
    {
        public int SelectedYear { get; private set; }
        public int SelectedMonth { get; private set; }
        public string VersionName { get; private set; } = "v1";

        public NewScheduleDialog()
        {
            InitializeComponent();
            var now = DateTime.Now;
            for (int y = now.Year - 1; y <= now.Year + 2; y++) YearBox.Items.Add(y);
            for (int m = 1; m <= 12; m++) MonthBox.Items.Add(m);
            YearBox.SelectedItem = now.Year;
            MonthBox.SelectedItem = now.Month;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (YearBox.SelectedItem == null || MonthBox.SelectedItem == null) return;
            SelectedYear = (int)YearBox.SelectedItem;
            SelectedMonth = (int)MonthBox.SelectedItem;
            VersionName = string.IsNullOrWhiteSpace(VersionBox.Text) ? "v1" : VersionBox.Text.Trim();
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
