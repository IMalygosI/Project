using System;
using Avalonia.Data.Converters;
using System.Globalization;

namespace GOST_Control
{
    public class DoubleToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString() ?? "14"; // Значение по умолчанию
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string strValue && double.TryParse(strValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                return result;
            return 14.0; // Значение по умолчанию
        }
    }
}