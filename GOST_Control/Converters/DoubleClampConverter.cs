using Avalonia.Data.Converters;
using System;
using System.Globalization;
using System.Linq;

namespace GOST_Control
{
    public class DoubleClampConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString() ?? "0";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string input = value?.ToString()?.Replace('.', ',') ?? "0";

            if (string.IsNullOrWhiteSpace(input))
                return 0.0;

            var parts = input.Split(',');
            if (parts.Length > 2)
            {
                input = parts[0] + "," + string.Join("", parts.Skip(1));
            }

            if (double.TryParse(input, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double result))
            {
                return result;
            }

            return 0.0;
        }
    }
}