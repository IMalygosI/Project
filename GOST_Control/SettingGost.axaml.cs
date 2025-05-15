using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using TextBox = Avalonia.Controls.TextBox;

namespace GOST_Control;

/// <summary>
/// Класс настройки параметров ГОСТа
/// </summary>
public partial class SettingGost : Window
{
    // Текущие настройки ГОСТа
    public Gost CurrentGost { get; set; }
    // Путь к файлу JSON
    private const string JsonFilePath = "gosts.json";

    /// <summary>
    /// Загрузка данных из Json
    /// </summary>
    public SettingGost()
    {
        InitializeComponent();

        TextDoc.Background = Brushes.White;

        // Загрузка JSON
        if (File.Exists(JsonFilePath))
        {
            var json = File.ReadAllText(JsonFilePath);
            var options = new JsonSerializerOptions
            {
                Converters = { new JsonStringEnumConverter() }
            };
            try
            {
                var gosts = JsonSerializer.Deserialize<List<Gost>>(json, options);
                CurrentGost = gosts?.FirstOrDefault() ?? new Gost();
            }
            catch
            {
                CurrentGost = new Gost();
            }
        }
        else
        {
            CurrentGost = new Gost();
        }

        DataContext = this;

        InitializePaperSizeUI();
    }

    /// <summary>
    /// Сохранение ГОСТа
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Save(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            // Настройки сериализации
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping, // Для кириллицы
                NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.Strict,  // Обработка чисел
                Converters = { new JsonStringEnumConverter() }  // Конвертер для Enum типов
            };

            var gosts = new List<Gost> { CurrentGost };
            string json = JsonSerializer.Serialize(gosts, options);

            // Записываем в файл JSON
            File.WriteAllText(JsonFilePath, json, System.Text.Encoding.UTF8);
            Close(true); // Закрытие окна или дополнительная логика после сохранения
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при сохранении: {ex.Message}");
        }
    }

    /// <summary>
    /// Открывает настройки простого текста
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_TextDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisTextDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;

        TextDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Острывает настройки Формата документа
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_PoleDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisMaketDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;

        MaketDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Открывает настройки заголовков
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_ZagolovkiDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisZagolovkiDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;

        ZagolovDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Открывает настройки заголовков
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Dop_ZagolovkiDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisDopZagolovkiDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;

        DopZagolovDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Для ввода заголовков
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnRequiredSectionsLostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            var normalized = string.Join(", ", textBox.Text.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s)));
            textBox.Text = normalized;
            CurrentGost.RequiredSections = normalized;
        }
    }

    /// <summary>
    /// Отображает настройку оглавлений
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_OglavleniaDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisOglavlenieDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;

        OglavleniaDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Отображает настройку картинок
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_ImageCheckDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisImageCheckDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;

        ImageCheckDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Отображение подписей таблиц
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_TablePodCheckDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {

        IsVisTablePodCheckDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTableCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;

        TablePodCheckDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TextInTableCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Отображение текст что в таблицах
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_TextInTableCheckDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisTableCheckDoc.IsVisible = true;

        IsVisListBulletCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;

        TextInTableCheckDoc.Background = Brushes.White;

        TextInBulletCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// Проверка списков
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_TextInBulletCheckDoc(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        IsVisListBulletCheckDoc.IsVisible = true;

        IsVisTableCheckDoc.IsVisible = false;
        IsVisTablePodCheckDoc.IsVisible = false;
        IsVisImageCheckDoc.IsVisible = false;
        IsVisOglavlenieDoc.IsVisible = false;
        IsVisDopZagolovkiDoc.IsVisible = false;
        IsVisMaketDoc.IsVisible = false;
        IsVisTextDoc.IsVisible = false;
        IsVisZagolovkiDoc.IsVisible = false;

        TextInBulletCheckDoc.Background = Brushes.White;

        TextInTableCheckDoc.Background = Brushes.Wheat;
        TablePodCheckDoc.Background = Brushes.Wheat;
        ImageCheckDoc.Background = Brushes.Wheat;
        OglavleniaDoc.Background = Brushes.Wheat;
        DopZagolovDoc.Background = Brushes.Wheat;
        ZagolovDoc.Background = Brushes.Wheat;
        TextDoc.Background = Brushes.Wheat;
        MaketDoc.Background = Brushes.Wheat;
    }

    /// <summary>
    /// блокируем ввод не цифр
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnTextInput_NumberOnly(object sender, Avalonia.Input.TextInputEventArgs e)
    {
        var textBox = sender as TextBox;

        // Проверка на вводимые символы (цифры или точка)
        if (e.Text.Any(c => !char.IsDigit(c) && c != '.' && c != ','))
        {
            e.Handled = true;  // Блокируем ввод, если это не цифра или точка
            return;
        }

        var currentText = textBox.Text;

        // Блокировка ввода второй точки
        if (currentText.Contains(".") && e.Text == "." && !currentText.EndsWith("."))
        {
            e.Handled = true;  // Блокируем, если точка уже есть
            return;
        }

        // Запрещаем удаление точки, если после неё есть цифры
        if (e.Text.Length == 0 && currentText.EndsWith("."))
        {
            e.Handled = true;  // Запрещаем удаление точки
            return;
        }
    }

    /// <summary>
    /// Макс значение для междустрочного интеврала
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnLineSpacingBoxLostFocus(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text.Trim();

            // Если строка пустая — 0
            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            // Замена запятых на точки для унификации
            inputText = inputText.Replace(',', '.');

            // Удаление всех лишних точек после первой
            int firstDotIndex = inputText.IndexOf('.');
            if (firstDotIndex != -1)
            {
                string beforeDot = inputText.Substring(0, firstDotIndex + 1);
                string afterDot = inputText.Substring(firstDotIndex + 1).Replace(".", string.Empty);
                inputText = beforeDot + afterDot;
            }

            if (double.TryParse(inputText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double value))
            {
                if (value > 132)
                    value = 132;

                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    /// <summary>
    /// Макс значение для стандартных полей
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnLineSpacingGenericLostFocus(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace(',', '.');

            int firstDotIndex = inputText.IndexOf('.');
            if (firstDotIndex != -1)
            {
                string beforeDot = inputText.Substring(0, firstDotIndex + 1);
                string afterDot = inputText.Substring(firstDotIndex + 1).Replace(".", string.Empty);
                inputText = beforeDot + afterDot;
            }

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value > 55.68)
                    value = 55.68;

                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    /// <summary>
    /// Макс значение для отступов листа слева и справа
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnMarginLeftRightLostFocus(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace(',', '.');

            int firstDotIndex = inputText.IndexOf('.');
            if (firstDotIndex != -1)
            {
                string beforeDot = inputText.Substring(0, firstDotIndex + 1);
                string afterDot = inputText.Substring(firstDotIndex + 1).Replace(".", string.Empty);
                inputText = beforeDot + afterDot;
            }

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                double maxValue = 19.73;  // Максимальное значение для отступа слева и справа
                if (value > maxValue)
                    value = maxValue;

                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    /// <summary>
    /// Макс значение для отступов листа сверху и снизу
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnMarginTopBottomLostFocus(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace(',', '.');

            int firstDotIndex = inputText.IndexOf('.');
            if (firstDotIndex != -1)
            {
                string beforeDot = inputText.Substring(0, firstDotIndex + 1);
                string afterDot = inputText.Substring(firstDotIndex + 1).Replace(".", string.Empty);
                inputText = beforeDot + afterDot;
            }

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                double maxValue = 29.44;  // Максимальное значение для отступа сверху и снизу
                if (value > maxValue)
                    value = maxValue;

                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    /// <summary>
    /// Обновление комбобокса и значений размера листа
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnPageSizeSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PageSizeComboBox.SelectedItem is string selectedPage)
        {
            if (selectedPage == "Особое")
            {
                WidthTextBox.IsEnabled = true;
                HeightTextBox.IsEnabled = true;
                WidthTextBox.Text = string.Empty;
                HeightTextBox.Text = string.Empty;
                CurrentGost.PaperSize = "Особое";
            }
            else
            {
                var pageSize = PageSizesRazmer.AllSizes.FirstOrDefault(size => size.Name == selectedPage);
                if (pageSize.Width != 0 && pageSize.Height != 0)
                {
                    WidthTextBox.Text = pageSize.Width.ToString("F2");
                    HeightTextBox.Text = pageSize.Height.ToString("F2");
                    CurrentGost.PaperSize = selectedPage;
                    CurrentGost.PaperWidthMm = pageSize.Width;
                    CurrentGost.PaperHeightMm = pageSize.Height;
                }

                WidthTextBox.IsEnabled = false;
                HeightTextBox.IsEnabled = false;
            }
        }
    }

    /// <summary>
    /// Загрузка данных размера
    /// </summary>
    private void InitializePaperSizeUI()
    {
        double widthCm = (CurrentGost.PaperWidthMm ?? 0);
        double heightCm = (CurrentGost.PaperHeightMm ?? 0);

        var matched = PageSizesRazmer.AllSizes.FirstOrDefault(p => Math.Abs(p.Width - widthCm) < 0.1 && Math.Abs(p.Height - heightCm) < 0.1);

        if (matched.Name != null)
        {
            foreach (var item in PageSizeComboBox.Items.OfType<ComboBoxItem>())
            {
                if ((string)item.Content == matched.Name)
                {
                    PageSizeComboBox.SelectedItem = item;
                    break;
                }
            }

            WidthTextBox.Text = matched.Width.ToString("F2");
            HeightTextBox.Text = matched.Height.ToString("F2");
            WidthTextBox.IsEnabled = false;
            HeightTextBox.IsEnabled = false;
        }
        else
        {
            // Особый формат
            foreach (var item in PageSizeComboBox.Items.OfType<ComboBoxItem>())
            {
                if ((string)item.Content == "Особое")
                {
                    PageSizeComboBox.SelectedItem = item;
                    break;
                }
            }

            WidthTextBox.Text = widthCm.ToString("F2");
            HeightTextBox.Text = heightCm.ToString("F2");
            WidthTextBox.IsEnabled = true;
            HeightTextBox.IsEnabled = true;
        }
    }

    private void OnDimensionLostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            // Заменяем запятую на точку (для корректного парсинга)
            inputText = inputText.Replace(',', '.');

            // Очищаем лишние точки
            int firstDotIndex = inputText.IndexOf('.');
            if (firstDotIndex != -1)
            {
                string beforeDot = inputText.Substring(0, firstDotIndex + 1);
                string afterDot = inputText.Substring(firstDotIndex + 1).Replace(".", string.Empty);
                inputText = beforeDot + afterDot;
            }

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value <= 0)
                {
                    value = 0;
                }
                else if (value > 118.9)
                {
                    value = 118.9;
                }

                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.InvariantCulture);

                // Принудительно установить "Особое", если пользователь вручную вводит значения
                if (PageSizeComboBox.SelectedItem is string selectedPage && selectedPage != "Особое")
                {
                    PageSizeComboBox.SelectedItem = "Особое";
                }

                // Сохраняем в CurrentGost
                if (textBox == WidthTextBox)
                {
                    CurrentGost.PaperWidthMm = value / 10;
                }
                else if (textBox == HeightTextBox)
                {
                    CurrentGost.PaperHeightMm = value / 10;
                }
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    /// <summary>
    /// Выход из окна настроек
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Setting_Out(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        Close(false);
    }


}