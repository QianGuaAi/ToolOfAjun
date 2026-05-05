using System;
using System.Globalization;
using System.Net.NetworkInformation;
using System.Windows.Data;
using System.Windows.Media;

namespace MyTools.ViewModels
{
    public class StatusToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is OperationalStatus status)
            {
                switch (status)
                {
                    case OperationalStatus.Up:
                        return Brushes.Green;
                    case OperationalStatus.Down:
                        return Brushes.Gray;
                    case OperationalStatus.Testing:
                        return Brushes.Orange;
                    default:
                        return Brushes.Red;
                }
            }
            return Brushes.Black;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
