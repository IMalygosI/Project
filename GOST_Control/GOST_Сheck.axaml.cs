using System;
using System.IO;
using System.Linq;  // Для методов расширений LINQ
using System.Text;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GOST_Control.Context;
using GOST_Control.Models;
using Microsoft.EntityFrameworkCore;

namespace GOST_Control;

public partial class GOST_Сheck : Window
{
    public DbSet<Gost> Gosts { get; set; }
    private readonly string _filePath;

    public GOST_Сheck(string filePath)
    {
        InitializeComponent();
        _filePath = filePath; // Сохраняем путь к файлу
        FilePathTextBlock.Text = $"Путь к файлу: {_filePath}"; // Отображаем путь в UI
    }

    /// <summary>
    /// Метод отвечающий за поиск ГОСТа в базе
    /// </summary>
    /// <param name="gostId"></param>
    /// <returns></returns>
    private async Task<Gost> GetGostByIdAsync(int gostId)
    {
        using (var context = new DimaBaseContext())
        {
            return await context.Gosts
                .FirstOrDefaultAsync(g => g.GostId == gostId);  // Просто ищем ГОСТ по его ID
        }
    }

    public async Task CheckFileForGostAsync(string filePath, int gostId)
    {
        // Получаем ГОСТ из базы данных
        var gost = await GetGostByIdAsync(gostId);
        if (gost == null)
        {
            ErrorControlGost.Text = "ГОСТ не найден в базе данных.";
            ErrorControlGost.Foreground = Brushes.Red; // Красный цвет
            return;
        }
        else
        {
            ErrorControlGost.Text = "ГОСТ найден в базе данных.";
            ErrorControlGost.Foreground = Brushes.Green; // Зеленый цвет
        }

        if (!string.IsNullOrEmpty(gost.FontName) || gost.FontSize.HasValue)
        {
            try
            {
                using (var wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    // Обновляем UI
                    if (wordDoc != null)
                    {
                        ErrorControl.Text = "Удалось открыть документ.";
                        ErrorControl.Foreground = Brushes.Green; // Зеленый цвет
                    }
                    else
                    {
                        ErrorControl.Text = "Не удалось открыть документ.";
                        ErrorControl.Foreground = Brushes.Red; // Красный цвет
                    }

                    // Забираем текст из документа
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Проверки ГОСТа
                    bool FontNameValid = true;
                    bool FontSizeValid = true;
                    bool MarginsValid = true;
                    bool LineSpacingValid = true;
                    bool FirstLineIndentValid = true;
                    bool TextAlignmentValid = true;



                    // Проверка типа шрифта
                    if (!string.IsNullOrEmpty(gost.FontName))
                    {
                        FontNameValid = CheckFontName(gost.FontName, body);
                    }
                    ErrorControlFont.Text = "Тип шрифта соответствует ГОСТу.";
                    ErrorControlFont.Foreground = Brushes.Green; // Зеленый цвет

                    // Проверка размера шрифта
                    if (gost.FontSize.HasValue)
                    {
                        FontSizeValid = CheckFontSize(gost.FontSize, body);
                    }
                    ErrorControlFontSize.Text = "Размер шрифта соответствует госту!";
                    ErrorControlFontSize.Foreground = Brushes.Green; // Зеленый цвет

                    // Проверка полей документа
                    MarginsValid = CheckMargins(gost.MarginTop, gost.MarginBottom, gost.MarginLeft, gost.MarginRight, body);
                    ErrorControlMargins.Text = "Поля документа соответствуют ГОСТу.";
                    ErrorControlMargins.Foreground = Brushes.Green; // Зеленый цвет

                    // Проверка межстрочного интервала
                    if (gost.LineSpacing.HasValue)
                    {
                        LineSpacingValid = CheckLineSpacing(gost.LineSpacing, body);
                    }
                    ErrorControlMnochitel.Text = "Межстрочный интервал соответствует ГОСТу.";
                    ErrorControlMnochitel.Foreground = Brushes.Green; // Зеленый цвет

                    // Проверка отступа  
                    if (gost.FirstLineIndent.HasValue)
                    {
                        FirstLineIndentValid = CheckFirstLineIndent(gost.FirstLineIndent, body);
                    }
                    ErrorControlFirstLineIndent.Text = "Отступ соответствует ГОСТу.";
                    ErrorControlFirstLineIndent.Foreground = Brushes.Green; // Зеленый цвет

                    // Проверка выравнивания текста
                    if (!string.IsNullOrEmpty(gost.TextAlignment))
                    {
                        TextAlignmentValid = CheckTextAlignment(gost.TextAlignment, body);
                    }
                    ErrorControlViravnivanie.Text = "Выравнивание текста соответствует ГОСТу.";
                    ErrorControlViravnivanie.Foreground = Brushes.Green; // Зеленый цвет



                    // Обновляем UI в зависимости от результатов проверки
                    if (FontNameValid && FontSizeValid && MarginsValid && LineSpacingValid && FirstLineIndentValid && TextAlignmentValid)
                    {
                        GostControl.Text = "Документ соответствует ГОСТу.";
                        GostControl.Foreground = Brushes.Green; // Зеленый цвет
                    }
                    else
                    {
                        GostControl.Text = "Документ не соответствует ГОСТу:";
                        GostControl.Foreground = Brushes.Red; // Красный цвет

                        if (!FontNameValid)
                        {
                            ErrorControlFont.Text = "Тип шрифта не соответствует.";
                            ErrorControlFont.Foreground = Brushes.Red; // Красный цвет
                        }
                        if (!FontSizeValid)
                        {
                            ErrorControlFontSize.Text = "Размер шрифта не соответствует.";
                            ErrorControlFontSize.Foreground = Brushes.Red; // Красный цвет
                        }
                        if (!MarginsValid)
                        {
                            ErrorControlMargins.Text = "Поля документа не соответствуют ГОСТу.";
                            ErrorControlMargins.Foreground = Brushes.Red; // Красный цвет
                        }                        
                        if (!LineSpacingValid)
                        {
                            ErrorControlMnochitel.Text = "Межстрочный интервал не соответствует ГОСТу.";
                            ErrorControlMnochitel.Foreground = Brushes.Red; // Красный цвет
                        }
                        if (!FirstLineIndentValid)
                        {
                            ErrorControlFirstLineIndent.Text = "Отступ не соответствует ГОСТу.";
                            ErrorControlFirstLineIndent.Foreground = Brushes.Red; // Красный цвет
                        }
                        if (!TextAlignmentValid)
                        {
                            ErrorControlViravnivanie.Text = "Выравнивание текста не соответствует ГОСТу.";
                            ErrorControlViravnivanie.Foreground = Brushes.Red; // Красный цвет
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GostControl.Text = $"Ошибка при открытии документа! Закройте докуммент!";
                GostControl.Foreground = Brushes.Red; // Красный цвет
            }
        }
    }









    /// <summary>
    /// Проверка на выравнивание 
    /// </summary>
    /// <param name="requiredAlignment"></param>
    /// <param name="body"></param>
    /// <returns></returns>
    private bool CheckTextAlignment(string requiredAlignment, Body body)
    {
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null) continue;

            var justification = paragraphProperties.Justification; // Выравнивание текста
            if (justification == null) continue;

            // Получаем текущее выравнивание
            string currentAlignment;

            if (justification.Val?.Value == JustificationValues.Left)
            {
                currentAlignment = "Left";
            }
            else if (justification.Val?.Value == JustificationValues.Center)
            {
                currentAlignment = "Center";
            }
            else if (justification.Val?.Value == JustificationValues.Right)
            {
                currentAlignment = "Right";
            }
            else if (justification.Val?.Value == JustificationValues.Both)
            {
                currentAlignment = "Justify";
            }
            else
            {
                currentAlignment = "Left";
            }

            // Сравниваем с требуемым выравниванием
            if (currentAlignment != requiredAlignment)
            {
                return false; // Выравнивание не соответствует ГОСТу
            }
        }
        return true; // Всё соответствует ГОСТу
    }

    /// <summary>
    /// Проверка на отступы
    /// </summary>
    /// <param name="requiredFirstLineIndent"></param>
    /// <param name="body"></param>
    /// <returns></returns>
    private bool CheckFirstLineIndent(double? requiredFirstLineIndent, Body body)
    {
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null) continue;

            var indent = paragraphProperties.Indentation;
            if (indent == null) continue;

            if (indent.FirstLine != null)
            {
                double firstLineIndentInCm = double.Parse(indent.FirstLine.Value) / 567.0; // 1 см = 567 twips

                double tolerance = 0.01;

                if (Math.Abs(firstLineIndentInCm - requiredFirstLineIndent.Value) > tolerance)
                {
                    return false; 
                }
            }
        }
        return true; 
    }

    /// <summary>
    /// Проверка межстрочного интервала на соответствие ГОСТу
    /// </summary>
    /// <param name="requiredLineSpacing"></param>
    /// <param name="body"></param>
    /// <returns></returns>
    private bool CheckLineSpacing(double? requiredLineSpacing, Body body)
    {
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null) continue;

            var spacing = paragraphProperties.SpacingBetweenLines;
            if (spacing == null) continue;

            if (spacing.Line != null && spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                double lineSpacing = double.Parse(spacing.Line.Value) / 240.0; // 240 = 1 строка

                if (Math.Abs(lineSpacing - requiredLineSpacing.Value) > 0.01) // Допустимая погрешность
                {
                    return false;
                }
            }
        }
        return true;
    }

    /// <summary>
    /// Метод проверки полей документа
    /// </summary>
    /// <param name="requiredMarginTop">    Верхний отступ </param>
    /// <param name="requiredMarginBottom"> Нижний отступ  </param>
    /// <param name="requiredMarginLeft">   Левый отступ   </param>
    /// <param name="requiredMarginRight">  Правый отступ  </param>
    /// <param name="body">Тело документа</param>
    /// <returns>True, если поля соответствуют ГОСТу, иначе False</returns>
    private bool CheckMargins(double? requiredMarginTop, double? requiredMarginBottom, double? requiredMarginLeft, double? requiredMarginRight, Body body)
    {
        var sectionProperties = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProperties == null) return false;

        var pageMargin = sectionProperties.Elements<PageMargin>().FirstOrDefault();
        if (pageMargin == null) return false;

        // Преобразуем значения полей в сантиметры
        double marginTopInCm = pageMargin.Top.Value / 567.0; // 1 см = 567 twips
        double marginBottomInCm = pageMargin.Bottom.Value / 567.0;
        double marginLeftInCm = pageMargin.Left.Value / 567.0;
        double marginRightInCm = pageMargin.Right.Value / 567.0;

        // Допустимая погрешность (0.01 см) 
        double pogreshnost = 0.01;

        // Проверка значениями из ГОСТа и из Документа
        if (requiredMarginTop.HasValue && Math.Abs(marginTopInCm - requiredMarginTop.Value) > pogreshnost) return false;
        if (requiredMarginBottom.HasValue && Math.Abs(marginBottomInCm - requiredMarginBottom.Value) > pogreshnost) return false;
        if (requiredMarginLeft.HasValue && Math.Abs(marginLeftInCm - requiredMarginLeft.Value) > pogreshnost) return false;
        if (requiredMarginRight.HasValue && Math.Abs(marginRightInCm - requiredMarginRight.Value) > pogreshnost) return false;

        return true; // Поля соответствуют ГОСТу
    }

    /// <summary>
    /// Размер шрифта
    /// </summary>
    /// <param name="requiredFontSize"></param>
    /// <param name="body"></param>
    /// <returns></returns>
    private bool CheckFontSize(double? requiredFontSize, Body body)
    {
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties == null) continue;

                // Получаем размер шрифта
                var fontSizeElement = runProperties.FontSize;
                if (fontSizeElement != null)
                {
                    // Преобразуем значение размера шрифта из строки в double
                    if (double.TryParse(fontSizeElement.Val.Value, out double fontSize))
                    {
                        double fontSizeInPoints = fontSize / 2;
                        // Сравниваем с требуемым размером шрифта
                        if (fontSizeInPoints != requiredFontSize.Value)
                        {
                            return false; // Размер шрифта не соответствует ГОСТу
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }
        return true;
    }

    /// <summary>
    /// Смотрим наименование шрифта
    /// </summary>
    /// <param name="requiredFontName"></param>
    /// <param name="body"></param>
    /// <returns></returns>
    private bool CheckFontName(string requiredFontName, Body body)
    {
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties == null) continue;

                // Получаем имя шрифта
                var fontNameElement = runProperties.RunFonts;
                string fontName = fontNameElement?.Ascii?.Value;

                if (fontName != null && fontName != requiredFontName)
                {
                    return false; // Тип шрифта не соответствует ГОСТу
                }
            }
        }
        return true; // Всё соответствует ГОСТу
    }

    /// <summary>
    /// Кнопка проверки на соответствие ГОСТу
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void Button_Click_SelectFile(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            // Указываем номер ГОСТа
            int gostId = 1; 

            // Проверяем файл на соответствие ГОСТу
            await CheckFileForGostAsync(_filePath, gostId);
        }
        catch (Exception ex)
        {
            GostControl.Text = $"Ошибка: {ex.Message}";
        }
    }

    /// <summary>
    /// Кнопка выхода
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Button_Click_LogOut(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        MainWindow mainWindow = new MainWindow();
        mainWindow.Show();
        Close();
    }
}
