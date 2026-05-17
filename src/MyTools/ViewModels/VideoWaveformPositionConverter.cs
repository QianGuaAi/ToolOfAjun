using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MyTools.ViewModels
{
    public class VideoWaveformPositionConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 3)
            {
                return new Thickness(0);
            }

            var width = ToDouble(values[0]);
            var positionSeconds = ToDouble(values[1]);
            var durationSeconds = ToDouble(values[2]);
            if (width <= 0 || durationSeconds <= 0)
            {
                return new Thickness(0);
            }

            var ratio = Math.Max(0, Math.Min(1, positionSeconds / durationSeconds));
            return new Thickness(Math.Max(0, (width - 2) * ratio), 0, 0, 0);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        private static double ToDouble(object value)
        {
            if (value == null || value == DependencyProperty.UnsetValue)
            {
                return 0;
            }

            try
            {
                return System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }
    }

    public class VideoWaveformRangeConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 4)
            {
                if (IsWidthRequest(parameter))
                {
                    return 0d;
                }

                return new Thickness(0);
            }

            var width = ToDouble(values[0]);
            var startSeconds = ToDouble(values[1]);
            var endSeconds = ToDouble(values[2]);
            var durationSeconds = ToDouble(values[3]);
            if (width <= 0 || durationSeconds <= 0 || startSeconds < 0 || endSeconds <= startSeconds)
            {
                if (IsWidthRequest(parameter))
                {
                    return 0d;
                }

                return new Thickness(0);
            }

            var leftRatio = Math.Max(0, Math.Min(1, startSeconds / durationSeconds));
            var rightRatio = Math.Max(0, Math.Min(1, endSeconds / durationSeconds));
            var left = width * Math.Min(leftRatio, rightRatio);
            var right = width * Math.Max(leftRatio, rightRatio);

            if (IsHandleRequest(parameter, "StartHandle"))
            {
                return new Thickness(ClampHandleLeft(left, width), 0, 0, 0);
            }

            if (IsHandleRequest(parameter, "EndHandle"))
            {
                return new Thickness(ClampHandleLeft(right, width), 0, 0, 0);
            }

            if (IsWidthRequest(parameter))
            {
                return Math.Max(0, right - left);
            }

            return new Thickness(left, 0, 0, 0);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        private static bool IsWidthRequest(object parameter)
        {
            return string.Equals(parameter as string, "Width", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHandleRequest(object parameter, string name)
        {
            return string.Equals(parameter as string, name, StringComparison.OrdinalIgnoreCase);
        }

        private static double ClampHandleLeft(double center, double width)
        {
            return Math.Max(0, Math.Min(Math.Max(0, width - 10), center - 5));
        }

        private static double ToDouble(object value)
        {
            if (value == null || value == DependencyProperty.UnsetValue)
            {
                return 0;
            }

            try
            {
                return System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }
    }
}
