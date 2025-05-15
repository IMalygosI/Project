using Avalonia.Data.Converters;
using System;
using System.Globalization;

namespace GOST_Control
{
    public class DoubleClampConverter : IValueConverter
    {
        private const double DefaultMaxValue = 132;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.ToString() ?? string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string input = value?.ToString()?.Replace(',', '.') ?? string.Empty;

            if (string.IsNullOrWhiteSpace(input))
                return null;

            // Убираем лишние точки
            int firstDotIndex = input.IndexOf('.');
            if (firstDotIndex != -1)
            {
                string beforeDot = input.Substring(0, firstDotIndex + 1);
                string afterDot = input.Substring(firstDotIndex + 1).Replace(".", string.Empty);
                input = beforeDot + afterDot;
            }

            if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                result = Math.Round(result, 2);

                // Получаем максимальное значение из параметра
                double max = DefaultMaxValue;
                if (parameter != null)
                {
                    if (double.TryParse(parameter.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedMax))
                    {
                        max = parsedMax;
                    }
                }

                // Ограничиваем значение максимальным значением
                return result > max ? max : result;
            }

            return null;
        }
    }
}
