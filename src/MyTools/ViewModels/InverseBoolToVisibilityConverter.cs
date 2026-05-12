using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MyTools.ViewModels
{
    /// <summary>true → Collapsed；false → Visible。用于"非"布尔到可见性的绑定。</summary>
    public class InverseBoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var flag = value is bool b && b;
            return flag ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is Visibility v && v != Visibility.Visible;
        }
    }
}
