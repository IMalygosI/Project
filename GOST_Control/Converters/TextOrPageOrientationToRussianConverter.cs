using Avalonia.Data.Converters;
using System;
using System.Globalization;


namespace GOST_Control
{
    /// <summary>
    /// Класс конвертаций данных с комбобокса
    /// </summary>
    public class TextOrPageOrientationToRussianConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string? category = parameter as string;

            return category switch
            {
                "Alignment" => value?.ToString() switch
                {
                    "Left" => "По левому краю",
                    "Center" => "По центру",
                    "Right" => "По правому краю",
                    "Both" => "По ширине",
                    _ => "По левому краю"
                },
                "PageOrientation" => value?.ToString() switch
                {
                    "Portrait" => "Книжная",
                    "Landscape" => "Альбомная",
                    _ => "Неизвестная ориентация"
                },
                "Position" => value?.ToString() switch
                {
                    "Top" => "Вверху",
                    "Bottom" => "Внизу",
                    _ => "Не указано"
                },
                "Bool" => value?.ToString() switch
                {
                    "True" => "Да",
                    "False" => "Нет",
                    _ => "Нет"
                },
                _ => string.Empty
            };
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
