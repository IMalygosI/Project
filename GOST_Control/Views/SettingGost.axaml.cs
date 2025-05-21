using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
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
using System.Threading.Tasks;
using TextBox = Avalonia.Controls.TextBox;

namespace GOST_Control;

/// <summary>
/// ����� ��������� ���������� �����
/// </summary>
public partial class SettingGost : Window
{
    // ������� ��������� �����
    public Gost CurrentGost { get; set; }
    // ���� � ����� JSON
    private const string JsonFilePath = "gosts_modified.json"; // ���������� ����
    private readonly JsonGostService _gostService;

    /// <summary>
    /// �������� ������ �� Json
    /// </summary>
    public SettingGost()
    {
        InitializeComponent();
        TextDoc.Background = Brushes.White;

        // �������������� ������
        _gostService = new JsonGostService("GOST_Control.Resources.gosts.json", JsonFilePath);

        LoadGostAsync().ConfigureAwait(false);

        DataContext = this;
        InitializePaperSizeUI();
    }

    private async Task LoadGostAsync()
    {
        try
        {
            var gosts = await _gostService.GetAllGostsAsync();
            CurrentGost = gosts.FirstOrDefault() ?? new Gost();
        }
        catch
        {
            CurrentGost = new Gost();
        }
    }

    /// <summary>
    /// ���������� �����
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void Button_Click_Save(object? sender, RoutedEventArgs e)
    {
        try
        {
            await _gostService.AddOrUpdateGostAsync(CurrentGost);
            Close(true);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"������ ��� ����������: {ex.Message}");
        }
    }

    /// <summary>
    /// ��������� ��������� �������� ������
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
    /// ��������� ��������� ������� ���������
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
    /// ��������� ��������� ����������
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
    /// ��������� ��������� ����������
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
    /// ��� ����� ����������
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
    /// ���������� ��������� ����������
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
    /// ���������� ��������� ��������
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
    /// ����������� �������� ������
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
    /// ����������� ����� ��� � ��������
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
    /// �������� �������
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
    /// ��������� ���� �� ����
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnTextInput_NumberOnly(object sender, Avalonia.Input.TextInputEventArgs e)
    {
        var textBox = sender as TextBox;
        var currentText = textBox.Text ?? "";

        // ��������� �����, �����, ������� � Backspace
        if (!char.IsDigit(e.Text[0]) && e.Text != "." && e.Text != "," && e.Text != "\b")
        {
            e.Handled = true;
            return;
        }

        // ���������� ������ �����/�������
        if ((e.Text == "." || e.Text == ",") && (currentText.Contains('.') || currentText.Contains(',')))
        {
            e.Handled = true;
        }
    }

    /// <summary>
    /// ���� �������� ��� �������������� ���������
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnLineSpacingBoxLostFocus(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text.Trim();

            // ���� ������ ������ � 0
            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            // ������ ������� �� ����� ��� ����������
            inputText = inputText.Replace(',', '.');

            // �������� ���� ������ ����� ����� ������
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
    /// ���� �������� ��� �������� ����� ����� � ������
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ','); // ������ � �������

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }


    /// <summary>
    /// ���� �������� ��� �������� ����� ������ � �����
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
                double maxValue = 29.44;  // ������������ �������� ��� ������� ������ � �����
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
    /// ���������� ���������� � �������� ������� �����
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnPageSizeSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PageSizeComboBox.SelectedItem is string selectedPage)
        {
            if (selectedPage == "������")
            {
                WidthTextBox.IsEnabled = true;
                HeightTextBox.IsEnabled = true;
                WidthTextBox.Text = string.Empty;
                HeightTextBox.Text = string.Empty;
                CurrentGost.PaperSize = "������";
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
    /// �������� ������ �������
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
            // ������ ������
            foreach (var item in PageSizeComboBox.Items.OfType<ComboBoxItem>())
            {
                if ((string)item.Content == "������")
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

            // �������� ������� �� ����� (��� ����������� ��������)
            inputText = inputText.Replace(',', '.');

            // ������� ������ �����
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

                // ������������� ���������� "������", ���� ������������ ������� ������ ��������
                if (PageSizeComboBox.SelectedItem is string selectedPage && selectedPage != "������")
                {
                    PageSizeComboBox.SelectedItem = "������";
                }

                // ��������� � CurrentGost
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
    /// ����� �� ���� ��������
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_Setting_Out(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        Close(false);
    }


    private void OnFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.TextIndentOrOutdent = "���";
                CurrentGost.FirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.FirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.TextIndentOrOutdent = "���";
                }
                else if (CurrentGost.TextIndentOrOutdent == "���")
                {
                    CurrentGost.TextIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.FirstLineIndent = 0;
                CurrentGost.TextIndentOrOutdent = "���";
            }
        }
    }

    private void OnIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.TextIndentOrOutdent == "���")
        {
            CurrentGost.FirstLineIndent = 0;
        }
        else if (CurrentGost.FirstLineIndent == 0)
        {
            CurrentGost.FirstLineIndent = 1.27;
        }

    }

    private string GetDefaultLineSpacingValue()
    {
        return CurrentGost.LineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private void OnFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.TextIndentOrOutdent == "���")
                {
                    CurrentGost.TextIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.TextIndentOrOutdent = "���";
                }
            }
        }
    }

    private void OnLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultLineSpacingValue();
                CurrentGost.LineSpacingValue = double.Parse(GetDefaultLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinLineSpacingValue();
                double maxValue = 132;

                // ���� �������� ������ ������������ ��� �������� ���� - ������������
                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.LineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.LineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }


    private double GetMinLineSpacingValue()
    {
        return CurrentGost.LineSpacingType switch
        {
            "���������" => 0.5, // ����������� ��������� (� ��)
            "�����" => 0.03,    // ����������� ������ �������� (� ��)
            "�������" => 0.03,  // ����������� �������� (� ��)
            _ => 0.03           // �������� �� ��������� (� ��)
        };
    }

    private void OnLineSpacingGenericLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ',');

            var parts = inputText.Split(',');
            if (parts.Length > 2)
            {
                inputText = parts[0] + "," + string.Join("", parts.Skip(1));
            }

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double maxValue = 55.87; // �� ���������
                double minValue = 0;

                if (textBox.Name == "LineSpacingTextBox")
                {
                    maxValue = 132;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    // ������ ��� ������ � ��������� ����������
    private void OnHeaderIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.HeaderIndentOrOutdent == "���")
        {
            CurrentGost.HeaderFirstLineIndent = 0;
        }
        else if (CurrentGost.HeaderFirstLineIndent == 0)
        {
            CurrentGost.HeaderFirstLineIndent = 1.27;
        }
    }

    private void OnHeaderFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.HeaderIndentOrOutdent = "���";
                CurrentGost.HeaderFirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.HeaderFirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.HeaderIndentOrOutdent = "���";
                }
                else if (CurrentGost.HeaderIndentOrOutdent == "���")
                {
                    CurrentGost.HeaderIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.HeaderFirstLineIndent = 0;
                CurrentGost.HeaderIndentOrOutdent = "���";
            }
        }
    }

    private void OnHeaderFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.HeaderIndentOrOutdent == "���")
                {
                    CurrentGost.HeaderIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.HeaderIndentOrOutdent = "���";
                }
            }
        }
    }

    // ������ ��� ������ � �������������� ����������� ����������
    private string GetDefaultHeaderLineSpacingValue()
    {
        return CurrentGost.HeaderLineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinHeaderLineSpacingValue()
    {
        return CurrentGost.HeaderLineSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnHeaderLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultHeaderLineSpacingValue();
                CurrentGost.HeaderLineSpacingValue = double.Parse(GetDefaultHeaderLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinHeaderLineSpacingValue();
                double maxValue = 132;

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.HeaderLineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultHeaderLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.HeaderLineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }





    // ======================= ������ ��� ������ � ��������� �������������� ���������� =======================

    private void OnAdditionalHeaderIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.AdditionalHeaderIndentOrOutdent == "���")
        {
            CurrentGost.AdditionalHeaderFirstLineIndent = 0;
        }
        else if (CurrentGost.AdditionalHeaderFirstLineIndent == 0)
        {
            CurrentGost.AdditionalHeaderFirstLineIndent = 1.27;
        }
    }

    private void OnAdditionalHeaderFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.AdditionalHeaderIndentOrOutdent = "���";
                CurrentGost.AdditionalHeaderFirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.AdditionalHeaderFirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.AdditionalHeaderIndentOrOutdent = "���";
                }
                else if (CurrentGost.AdditionalHeaderIndentOrOutdent == "���")
                {
                    CurrentGost.AdditionalHeaderIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.AdditionalHeaderFirstLineIndent = 0;
                CurrentGost.AdditionalHeaderIndentOrOutdent = "���";
            }
        }
    }

    private void OnAdditionalHeaderFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.AdditionalHeaderIndentOrOutdent == "���")
                {
                    CurrentGost.AdditionalHeaderIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.AdditionalHeaderIndentOrOutdent = "���";
                }
            }
        }
    }

    // ������ ��� ������ � �������������� ����������� �������������� ����������
    private string GetDefaultAdditionalHeaderLineSpacingValue()
    {
        return CurrentGost.AdditionalHeaderLineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinAdditionalHeaderLineSpacingValue()
    {
        return CurrentGost.AdditionalHeaderLineSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnAdditionalHeaderLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultAdditionalHeaderLineSpacingValue();
                CurrentGost.AdditionalHeaderLineSpacingValue = double.Parse(GetDefaultAdditionalHeaderLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinAdditionalHeaderLineSpacingValue();
                double maxValue = 132;

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.AdditionalHeaderLineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultAdditionalHeaderLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.AdditionalHeaderLineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }

    private void OnAdditionalHeaderMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                // ��������� ��������������� �������� � ����������� �� ����, ����� ��� ����
                if (textBox.Name.Contains("Left"))
                {
                    CurrentGost.AdditionalHeaderIndentLeft = value;
                }
                else if (textBox.Name.Contains("Right"))
                {
                    CurrentGost.AdditionalHeaderIndentRight = value;
                }
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }



    // ======================= ������ ��� ������ � ��������� ���������� =======================

    private void OnTocIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.TocIndentOrOutdent == "���")
        {
            CurrentGost.TocFirstLineIndent = 0;
        }
        else if (CurrentGost.TocFirstLineIndent == 0)
        {
            CurrentGost.TocFirstLineIndent = 1.27;
        }
    }

    private void OnTocFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.TocIndentOrOutdent = "���";
                CurrentGost.TocFirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.TocFirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.TocIndentOrOutdent = "���";
                }
                else if (CurrentGost.TocIndentOrOutdent == "���")
                {
                    CurrentGost.TocIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.TocFirstLineIndent = 0;
                CurrentGost.TocIndentOrOutdent = "���";
            }
        }
    }

    private void OnTocFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.TocIndentOrOutdent == "���")
                {
                    CurrentGost.TocIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.TocIndentOrOutdent = "���";
                }
            }
        }
    }

    // ������ ��� ������ � �������������� ����������� ����������
    private string GetDefaultTocLineSpacingValue()
    {
        return CurrentGost.TocLineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinTocLineSpacingValue()
    {
        return CurrentGost.TocLineSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnTocLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultTocLineSpacingValue();
                CurrentGost.TocLineSpacing = double.Parse(GetDefaultTocLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinTocLineSpacingValue();
                double maxValue = 132;

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.TocLineSpacing = value;
            }
            else
            {
                string defaultValue = GetDefaultTocLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.TocLineSpacing = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }

    private void OnTocMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                // ��������� ��������������� �������� � ����������� �� ����, ����� ��� ����
                if (textBox.Name.Contains("Left"))
                {
                    CurrentGost.TocIndentLeft = value;
                }
                else if (textBox.Name.Contains("Right"))
                {
                    CurrentGost.TocIndentRight = value;
                }
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }





    // ======================= ������ ��� ������ � ��������� �������� � ������������ =======================

    private void OnImageCaptionIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.ImageCaptionIndentOrOutdent == "���")
        {
            CurrentGost.ImageCaptionFirstLineIndent = 0;
        }
        else if (CurrentGost.ImageCaptionFirstLineIndent == 0)
        {
            CurrentGost.ImageCaptionFirstLineIndent = 1.27;
        }
    }

    private void OnImageCaptionFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.ImageCaptionIndentOrOutdent = "���";
                CurrentGost.ImageCaptionFirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.ImageCaptionFirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.ImageCaptionIndentOrOutdent = "���";
                }
                else if (CurrentGost.ImageCaptionIndentOrOutdent == "���")
                {
                    CurrentGost.ImageCaptionIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.ImageCaptionFirstLineIndent = 0;
                CurrentGost.ImageCaptionIndentOrOutdent = "���";
            }
        }
    }

    private void OnImageCaptionFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.ImageCaptionIndentOrOutdent == "���")
                {
                    CurrentGost.ImageCaptionIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.ImageCaptionIndentOrOutdent = "���";
                }
            }
        }
    }

    // ������ ��� ������ � �������������� ����������� �������� � ������������
    private string GetDefaultImageCaptionLineSpacingValue()
    {
        return CurrentGost.ImageCaptionLineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinImageCaptionLineSpacingValue()
    {
        return CurrentGost.ImageCaptionLineSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnImageCaptionLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultImageCaptionLineSpacingValue();
                CurrentGost.ImageCaptionLineSpacingValue = double.Parse(GetDefaultImageCaptionLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinImageCaptionLineSpacingValue();
                double maxValue = 132;

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.ImageCaptionLineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultImageCaptionLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.ImageCaptionLineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }

    private void OnImageCaptionMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                // ��������� ��������������� �������� � ����������� �� ����, ����� ��� ����
                if (textBox.Name.Contains("Left"))
                {
                    CurrentGost.ImageCaptionIndentLeft = value;
                }
                else if (textBox.Name.Contains("Right"))
                {
                    CurrentGost.ImageCaptionIndentRight = value;
                }
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    // ======================= ������ ��� ������ � ��������� �������� ��� ��������� =======================

    private void OnTableCaptionIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.TableCaptionIndentOrOutdent == "���")
        {
            CurrentGost.TableCaptionFirstLineIndent = 0;
        }
        else if (CurrentGost.TableCaptionFirstLineIndent == 0)
        {
            CurrentGost.TableCaptionFirstLineIndent = 1.27;
        }
    }

    private void OnTableCaptionFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.TableCaptionIndentOrOutdent = "���";
                CurrentGost.TableCaptionFirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.TableCaptionFirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.TableCaptionIndentOrOutdent = "���";
                }
                else if (CurrentGost.TableCaptionIndentOrOutdent == "���")
                {
                    CurrentGost.TableCaptionIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.TableCaptionFirstLineIndent = 0;
                CurrentGost.TableCaptionIndentOrOutdent = "���";
            }
        }
    }

    private void OnTableCaptionFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.TableCaptionIndentOrOutdent == "���")
                {
                    CurrentGost.TableCaptionIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.TableCaptionIndentOrOutdent = "���";
                }
            }
        }
    }

    // ������ ��� ������ � �������������� ����������� �������� ��� ���������
    private string GetDefaultTableCaptionLineSpacingValue()
    {
        return CurrentGost.TableCaptionLineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinTableCaptionLineSpacingValue()
    {
        return CurrentGost.TableCaptionLineSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnTableCaptionLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultTableCaptionLineSpacingValue();
                CurrentGost.TableCaptionLineSpacingValue = double.Parse(GetDefaultTableCaptionLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinTableCaptionLineSpacingValue();
                double maxValue = 132;

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.TableCaptionLineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultTableCaptionLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.TableCaptionLineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }

    private void OnTableCaptionMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                // ��������� ��������������� �������� � ����������� �� ����, ����� ��� ����
                if (textBox.Name.Contains("Left"))
                {
                    CurrentGost.TableCaptionIndentLeft = value;
                }
                else if (textBox.Name.Contains("Right"))
                {
                    CurrentGost.TableCaptionIndentRight = value;
                }
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }

    // ======================= ������ ��� ������ � ��������� ������ =======================

    private void OnTableIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (CurrentGost.TableIndentOrOutdent == "���")
        {
            CurrentGost.TableFirstLineIndent = 0;
        }
        else if (CurrentGost.TableFirstLineIndent == 0)
        {
            CurrentGost.TableFirstLineIndent = 1.27;
        }
    }

    private void OnTableFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                CurrentGost.TableIndentOrOutdent = "���";
                CurrentGost.TableFirstLineIndent = 0;
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.TableFirstLineIndent = value;

                if (value == 0)
                {
                    CurrentGost.TableIndentOrOutdent = "���";
                }
                else if (CurrentGost.TableIndentOrOutdent == "���")
                {
                    CurrentGost.TableIndentOrOutdent = "������";
                }
            }
            else
            {
                textBox.Text = "0";
                CurrentGost.TableFirstLineIndent = 0;
                CurrentGost.TableIndentOrOutdent = "���";
            }
        }
    }

    private void OnTableFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && CurrentGost.TableIndentOrOutdent == "���")
                {
                    CurrentGost.TableIndentOrOutdent = "������";
                }
                else if (value == 0)
                {
                    CurrentGost.TableIndentOrOutdent = "���";
                }
            }
        }
    }

    private string GetDefaultTableLineSpacingValue()
    {
        return CurrentGost.TableSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinTableLineSpacingValue()
    {
        return CurrentGost.TableSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnTableLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultTableLineSpacingValue();
                CurrentGost.TableLineSpacingValue = double.Parse(GetDefaultTableLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = GetMinTableLineSpacingValue();
                double maxValue = 132;

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, maxValue);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.TableLineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultTableLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.TableLineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }

    private void OnTableMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                // ��������� ��������������� �������� � ����������� �� ����, ����� ��� ����
                if (textBox.Name.Contains("Left"))
                {
                    CurrentGost.TableIndentLeft = value;
                }
                else if (textBox.Name.Contains("Right"))
                {
                    CurrentGost.TableIndentRight = value;
                }
            }
            else
            {
                textBox.Text = "0";
            }
        }
    }


    // ======================= ������ ��� ������ � ��������� ������� =======================

    private void OnBulletIndentTypeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox comboBox && comboBox.SelectedItem != null)
        {
            string indentType = comboBox.SelectedItem.ToString();

            // ���������� ������� ������ �� ����� ����������
            string level = comboBox.Name.Replace("IndentOrOutdent", "");

            // �������� �������� ��� �������� ������
            var indentOrOutdentProperty = typeof(Gost).GetProperty($"{level}IndentOrOutdent");
            var firstLineIndentProperty = typeof(Gost).GetProperty($"{level}FirstLineIndent");

            if (indentOrOutdentProperty == null || firstLineIndentProperty == null) return;

            // ������������� ��� �������
            indentOrOutdentProperty.SetValue(CurrentGost, indentType);

            // �������� ������� �������� �������
            double currentIndent = (double)firstLineIndentProperty.GetValue(CurrentGost);

            if (indentType == "���")
            {
                firstLineIndentProperty.SetValue(CurrentGost, 0.0);
            }
            else if (currentIndent == 0)
            {
                firstLineIndentProperty.SetValue(CurrentGost, 1.27);
            }
        }
    }



    private void OnBulletFirstLineIndentLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string propertyPrefix = textBox.Name.Replace("FirstLineIndentTextBox", "")
                                      .Replace("IndentTextBox", "");

            string inputText = textBox.Text?.Trim() ?? "";

            var indentProperty = typeof(Gost).GetProperty($"{propertyPrefix}Indent");
            var indentOrOutdentProperty = typeof(Gost).GetProperty($"{propertyPrefix}IndentOrOutdent");

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                indentOrOutdentProperty?.SetValue(CurrentGost, "���");
                indentProperty?.SetValue(CurrentGost, 0.0); 
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, 0, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                indentProperty?.SetValue(CurrentGost, value);

                if (value == 0)
                {
                    indentOrOutdentProperty?.SetValue(CurrentGost, "���");
                }
                else if ((string)(indentOrOutdentProperty?.GetValue(CurrentGost) ?? "���") == "���")
                {
                    indentOrOutdentProperty?.SetValue(CurrentGost, "������");
                }
            }
            else
            {
                textBox.Text = "0";
                indentProperty?.SetValue(CurrentGost, 0.0); 
                indentOrOutdentProperty?.SetValue(CurrentGost, "���");
            }
        }
    }

    private void OnBulletFirstLineIndentChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is TextBox textBox && textBox.IsFocused)
        {
            string propertyPrefix = textBox.Name.Replace("FirstLineIndentTextBox", "")
                                      .Replace("IndentTextBox", "");

            string inputText = textBox.Text?.Trim() ?? "";

            if (!string.IsNullOrEmpty(inputText) &&
                double.TryParse(inputText.Replace('.', ','), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                if (value != 0 && (string)(typeof(Gost).GetProperty($"{propertyPrefix}IndentOrOutdent")?.GetValue(CurrentGost) ?? "���") == "���")
                {
                    typeof(Gost).GetProperty($"{propertyPrefix}IndentOrOutdent")?.SetValue(CurrentGost, "������");
                }
                else if (value == 0)
                {
                    typeof(Gost).GetProperty($"{propertyPrefix}IndentOrOutdent")?.SetValue(CurrentGost, "���");
                }
            }
        }
    }

    // ������ ��� ������ � �������������� ����������� �������
    private string GetDefaultBulletLineSpacingValue()
    {
        return CurrentGost.BulletLineSpacingType switch
        {
            "���������" => "1,0",
            "�����" => "0,5",
            "�������" => "0,5",
            _ => "0"
        };
    }

    private double GetMinBulletLineSpacingValue()
    {
        return CurrentGost.BulletLineSpacingType switch
        {
            "���������" => 0.5,
            "�����" => 0.03,
            "�������" => 0.03,
            _ => 0.03
        };
    }

    private void OnBulletLineSpacingLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string inputText = textBox.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = GetDefaultBulletLineSpacingValue();
                CurrentGost.BulletLineSpacingValue = double.Parse(GetDefaultBulletLineSpacingValue(), CultureInfo.GetCultureInfo("ru-RU"));
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                double minValue = CurrentGost.BulletLineSpacingType switch
                {
                    "���������" => 0.5,
                    "�����" => 0.03,
                    "�������" => 0.03,
                    _ => 0.03
                };

                if (value < minValue)
                {
                    value = minValue;
                }

                value = Math.Clamp(value, minValue, 132);
                value = Math.Round(value, 2);

                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));
                CurrentGost.BulletLineSpacingValue = value;
            }
            else
            {
                string defaultValue = GetDefaultBulletLineSpacingValue();
                textBox.Text = defaultValue;
                CurrentGost.BulletLineSpacingValue = double.Parse(defaultValue, CultureInfo.GetCultureInfo("ru-RU"));
            }
        }
    }

    // ����� ����� ��� ��������� �������� �����/������ ��� ���� ������� �������
    private void OnBulletMarginLeftRightLostFocus(object? sender, RoutedEventArgs e)
    {
        if (sender is TextBox textBox)
        {
            string propertyName = textBox.Name.Replace("TextBox", "");
            string inputText = textBox.Text?.Trim() ?? "";

            var property = typeof(Gost).GetProperty(propertyName);
            if (property == null) return;

            Type propertyType = property.PropertyType;
            bool isNullableDouble = propertyType == typeof(double?);

            if (string.IsNullOrWhiteSpace(inputText))
            {
                textBox.Text = "0";
                // ������������� ���������� ��� ��������
                if (isNullableDouble)
                    property.SetValue(CurrentGost, (double?)0);
                else
                    property.SetValue(CurrentGost, 0.0);
                return;
            }

            inputText = inputText.Replace('.', ',');

            if (double.TryParse(inputText, NumberStyles.Any, CultureInfo.GetCultureInfo("ru-RU"), out double value))
            {
                value = Math.Clamp(value, -55.87, 55.87);
                value = Math.Round(value, 2);
                textBox.Text = value.ToString("0.##", CultureInfo.GetCultureInfo("ru-RU"));

                // ������������� �������� � ������ ���� ��������
                if (isNullableDouble)
                    property.SetValue(CurrentGost, (double?)value);
                else
                    property.SetValue(CurrentGost, value);
            }
            else
            {
                textBox.Text = "0";
                // ������������� ���������� ��� ��������
                if (isNullableDouble)
                    property.SetValue(CurrentGost, (double?)0);
                else
                    property.SetValue(CurrentGost, 0.0);
            }
        }
    }
}