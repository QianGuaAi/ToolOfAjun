using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace MyTools.ViewModels
{
    public class ConflictLevelToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ConflictLevel level)
            {
                return level == ConflictLevel.Hard
                    ? new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36))
                    : new SolidColorBrush(Color.FromRgb(0xFF, 0xAB, 0x00));
            }
            return new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
