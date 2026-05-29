using System;
using System.Globalization;
using System.Windows.Data;

namespace MyTools.ViewModels
{
    public class ConflictLevelToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ConflictLevel level)
            {
                return level == ConflictLevel.Hard ? "硬" : "软";
            }
            return "硬";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
