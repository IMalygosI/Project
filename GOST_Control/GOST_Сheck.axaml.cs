using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GOST_Control
{
    public partial class GOST_Сheck : Window
    {
        private readonly string _filePath;
        private JsonGostService _gostService;
        private readonly Task _initializationTask;

        public GOST_Сheck()
        {
            InitializeComponent();
        }

        public GOST_Сheck(string filePath)
        {
            InitializeComponent();

            _filePath = filePath;

            _initializationTask = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            try
            {
                _gostService = await Task.Run(() =>
                {
                    return new JsonGostService("gosts.json");
                });

                Dispatcher.UIThread.Post(() => { FilePathTextBlock.Text = $"Путь к файлу: {_filePath}"; });
            }
            catch (Exception ex)
            {
                Dispatcher.UIThread.Post(() => { FilePathTextBlock.Text = $"Ошибка загрузки: {ex.Message}"; });
                Console.WriteLine($"Load error: {ex}");
            }
        }

        /// <summary>
        /// Метод отвечающий за поиск ГОСТа в JSON файле
        /// </summary>
        /// <param name="gostId">ID ГОСТа для поиска</param>
        /// <returns>Найденный ГОСТ или null</returns>
        private async Task<Gost> GetGostByIdAsync(int gostId)
        {
            return await _gostService.GetGostByIdAsync(gostId);
        }

        /// <summary>
        /// Основной метод проверки документа на соответствие ГОСТу
        /// </summary>
        /// <param name="filePath">Путь к проверяемому файлу</param>
        /// <param name="gostId">ID ГОСТа для проверки</param>
        public async Task CheckFileForGostAsync(string filePath, int gostId)
        {
            await _initializationTask;

            if (_gostService == null)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlGost.Text = "Сервис не инициализирован";
                    ErrorControlGost.Foreground = Brushes.Red;
                });
                return;
            }

            // Получаем ГОСТ из JSON файла
            var gost = await GetGostByIdAsync(gostId);

            if (gost == null)
            {
                ErrorControlGost.Text = "ГОСТ не найден в JSON файле.";
                ErrorControlGost.Foreground = Brushes.Red;
                return;
            }
            else
            {
                ErrorControlGost.Text = "ГОСТ найден в JSON файле.";
                ErrorControlGost.Foreground = Brushes.Green;
            }

            if (!string.IsNullOrEmpty(gost.FontName) || gost.FontSize.HasValue)
            {
                try
                {
                    using (var wordDoc = WordprocessingDocument.Open(filePath, false))
                    {
                        if (wordDoc != null)
                        {
                            ErrorControl.Text = "Удалось открыть документ.";
                            ErrorControl.Foreground = Brushes.Green;
                        }
                        else
                        {
                            ErrorControl.Text = "Не удалось открыть документ.";
                            ErrorControl.Foreground = Brushes.Red;
                        }

                        var body = wordDoc.MainDocumentPart.Document.Body;

                        // Флаги результатов проверок
                        bool fontNameValid = true;
                        bool fontSizeValid = true;
                        bool marginsValid = true;
                        bool lineSpacingValid = true;
                        bool firstLineIndentValid = true;
                        bool textAlignmentValid = true;
                        bool pageNumberingValid = true;

                        // Проверка типа шрифта (игнорируя заголовки разделов)
                        if (!string.IsNullOrEmpty(gost.FontName))
                        {
                            fontNameValid = CheckFontName(gost.FontName, body, gost);
                            ErrorControlFont.Text = fontNameValid ? "Тип шрифта соответствует ГОСТу." : "Тип шрифта не соответствует.";
                            ErrorControlFont.Foreground = fontNameValid ? Brushes.Green : Brushes.Red;
                        }

                        // Проверка размера шрифта (игнорируя заголовки разделов)
                        if (gost.FontSize.HasValue)
                        {
                            fontSizeValid = CheckFontSize(gost.FontSize.Value, body, gost);
                            ErrorControlFontSize.Text = fontSizeValid ? "Размер шрифта соответствует ГОСТу!" : "Размер шрифта не соответствует!";
                            ErrorControlFontSize.Foreground = fontSizeValid ? Brushes.Green : Brushes.Red;
                        }

                        // Проверка полей документа
                        if (gost.MarginTop.HasValue || gost.MarginBottom.HasValue ||
                            gost.MarginLeft.HasValue || gost.MarginRight.HasValue)
                        {
                            marginsValid = CheckMargins(gost.MarginTop, gost.MarginBottom, gost.MarginLeft, gost.MarginRight, body);
                            ErrorControlMargins.Text = marginsValid ? "Поля документа соответствуют ГОСТу." : "Поля документа не соответствуют ГОСТу.";
                            ErrorControlMargins.Foreground = marginsValid ? Brushes.Green : Brushes.Red;
                        }

                        // Проверка межстрочного интервала
                        if (gost.LineSpacing.HasValue)
                        {
                            lineSpacingValid = CheckLineSpacing(gost.LineSpacing.Value, body);
                            ErrorControlMnochitel.Text = lineSpacingValid ? "Межстрочный интервал соответствует ГОСТу." : "Межстрочный интервал не соответствует ГОСТу.";
                            ErrorControlMnochitel.Foreground = lineSpacingValid ? Brushes.Green : Brushes.Red;
                        }

                        // Проверка отступа первой строки
                        if (gost.FirstLineIndent.HasValue)
                        {
                            firstLineIndentValid = CheckFirstLineIndent(gost.FirstLineIndent.Value, body);
                            ErrorControlFirstLineIndent.Text = firstLineIndentValid ? "Отступ соответствует ГОСТу." : "Отступ не соответствует ГОСТу.";
                            ErrorControlFirstLineIndent.Foreground = firstLineIndentValid ? Brushes.Green : Brushes.Red;
                        }

                        // Проверка выравнивания текста (игнорируя заголовки разделов)
                        if (!string.IsNullOrEmpty(gost.TextAlignment))
                        {
                            textAlignmentValid = CheckTextAlignment(gost.TextAlignment, body, gost);
                            ErrorControlViravnivanie.Text = textAlignmentValid ? "Выравнивание текста соответствует ГОСТу." : "Выравнивание текста не соответствует ГОСТу.";
                            ErrorControlViravnivanie.Foreground = textAlignmentValid ? Brushes.Green : Brushes.Red;
                        }

                        // Проверка нумерации страниц
                        if (gost.PageNumbering.HasValue)
                        {
                            pageNumberingValid = CheckPageNumbering(wordDoc, gost.PageNumbering.Value);
                            ErrorControlNumberPage.Text = pageNumberingValid ? "Нумерация страниц соответствует ГОСТу." : "Нумерация страниц не соответствует ГОСТу.";
                            ErrorControlNumberPage.Foreground = pageNumberingValid ? Brushes.Green : Brushes.Red;
                        }
                        else
                        {
                            ErrorControlNumberPage.Text = "Нумерация страниц не требуется.";
                            ErrorControlNumberPage.Foreground = Brushes.Gray;
                        }

                        // Проверка обязательных разделов (Введение, Заключение)
                        bool sectionsValid = true;
                        if (!string.IsNullOrEmpty(gost.RequiredSections))
                        {
                            sectionsValid = CheckRequiredSections(gost, body);
                        }

                        // Общий результат проверки
                        if (fontNameValid && fontSizeValid && marginsValid && lineSpacingValid &&
                            firstLineIndentValid && textAlignmentValid && pageNumberingValid && sectionsValid)
                        {
                            GostControl.Text = "Документ соответствует ГОСТу.";
                            GostControl.Foreground = Brushes.Green;
                        }
                        else
                        {
                            GostControl.Text = "Документ не соответствует ГОСТу:";
                            GostControl.Foreground = Brushes.Red;
                        }
                    }
                }
                catch (Exception ex)
                {
                    GostControl.Text = $"Ошибка при открытии документа! Закройте документ!";
                    GostControl.Foreground = Brushes.Red;
                }
            }
        }


        /// <summary>
        /// Проверка обязательных разделов (Введение, Заключение и т.д.)
        /// </summary>
        private bool CheckRequiredSections(Gost gost, Body body)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections))
                return true; // Если разделы не требуются, проверка пройдена

            var requiredSections = gost.RequiredSections.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            if (!requiredSections.Any())
                return true;

            bool allSectionsFound = true;
            bool allSectionsValid = true;
            var missingSections = new List<string>();
            var invalidSections = new List<string>();

            foreach (var section in requiredSections)
            {
                bool sectionFound = false;
                bool sectionValid = true;

                // Ищем параграфы, содержащие название раздела
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    var text = paragraph.InnerText.Trim();
                    if (text.Equals(section, StringComparison.OrdinalIgnoreCase))
                    {
                        sectionFound = true;

                        // Проверяем шрифт заголовка
                        if (gost.HeaderFontSize.HasValue)
                        {
                            var run = paragraph.Elements<Run>().FirstOrDefault();
                            if (run != null)
                            {
                                var fontSize = run.RunProperties?.FontSize;
                                if (fontSize != null)
                                {
                                    double fontSizeValue = double.Parse(fontSize.Val.Value) / 2;
                                    if (Math.Abs(fontSizeValue - gost.HeaderFontSize.Value) > 0.1)
                                    {
                                        invalidSections.Add($"{section} (неверный размер шрифта)");
                                        sectionValid = false;
                                        break;
                                    }
                                }
                                else
                                {
                                    invalidSections.Add($"{section} (отсутствует размер шрифта)");
                                    sectionValid = false;
                                    break;
                                }
                            }
                            else
                            {
                                invalidSections.Add($"{section} (отсутствует Run)");
                                sectionValid = false;
                                break;
                            }
                        }

                        // Проверяем выравнивание заголовка
                        if (!string.IsNullOrEmpty(gost.HeaderAlignment))
                        {
                            var justification = paragraph.ParagraphProperties?.Justification;
                            var currentAlignment = GetAlignmentString(justification);

                            if (currentAlignment != gost.HeaderAlignment)
                            {
                                invalidSections.Add($"{section} (неверное выравнивание)");
                                sectionValid = false;
                                break;
                            }
                        }
                    }
                }

                if (!sectionFound)
                {
                    missingSections.Add(section);
                    allSectionsFound = false;
                }
                else if (!sectionValid)
                {
                    allSectionsValid = false;
                }
            }

            // Обновляем интерфейс
            Dispatcher.UIThread.Post(() =>
            {
                if (!allSectionsFound)
                {
                    ErrorControlSections.Text = $"Не найдены разделы: {string.Join(", ", missingSections)}";
                    ErrorControlSections.Foreground = Brushes.Red;
                }
                else if (!allSectionsValid)
                {
                    ErrorControlSections.Text = $"Найдены, но не соответствуют: {string.Join(", ", invalidSections)}";
                    ErrorControlSections.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlSections.Text = "Все обязательные разделы найдены и соответствуют требованиям";
                    ErrorControlSections.Foreground = Brushes.Green;
                }
            });

            return allSectionsFound && allSectionsValid;
        }

        // Вспомогательный метод для получения строкового представления выравнивания
        private string GetAlignmentString(Justification justification)
        {
            if (justification == null) return "По левому краю";

            string alignment;

            if (justification.Val?.Value == JustificationValues.Left)
            {
                alignment = "Left";
            }
            else if (justification.Val?.Value == JustificationValues.Center)
            {
                alignment = "Center";
            }
            else if (justification.Val?.Value == JustificationValues.Right)
            {
                alignment = "Right";
            }
            else if (justification.Val?.Value == JustificationValues.Both)
            {
                alignment = "Both";
            }
            else
            {
                alignment = "Left";
            }

            return alignment;
        }


        /// <summary>
        /// Получает список всех заголовков разделов из документа
        /// </summary>
        private List<Paragraph> GetAllSectionHeaders(Body body, Gost gost)
        {
            var requiredSections = GetRequiredSectionsList(gost);
            var headers = new List<Paragraph>();

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var text = paragraph.InnerText.Trim();
                if (requiredSections.Any(s => text.Equals(s, StringComparison.OrdinalIgnoreCase)))
                {
                    headers.Add(paragraph);
                }
            }
            return headers;
        }

        /// <summary>
        /// Проверка типа шрифта (полностью исключая заголовки)
        /// </summary>
        private bool CheckFontName(string requiredFontName, Body body, Gost gost)
        {
            var headers = GetAllSectionHeaders(body, gost);

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем все заголовки
                if (headers.Contains(paragraph))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    var fontName = run.RunProperties?.RunFonts?.Ascii?.Value;
                    if (fontName != null && fontName != requiredFontName)
                    {
                        Console.WriteLine($"Ошибка шрифта в тексте: {paragraph.InnerText}");
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Проверка размера шрифта (полностью исключая заголовки)
        /// </summary>
        private bool CheckFontSize(double requiredFontSize, Body body, Gost gost)
        {
            var headers = GetAllSectionHeaders(body, gost);

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем все заголовки
                if (headers.Contains(paragraph))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    var fontSize = run.RunProperties?.FontSize;
                    if (fontSize == null) continue;

                    if (double.TryParse(fontSize.Val.Value, out double fontSizeValue))
                    {
                        double fontSizeInPoints = fontSizeValue / 2;
                        if (Math.Abs(fontSizeInPoints - requiredFontSize) > 0.1)
                        {
                            Console.WriteLine($"Ошибка размера шрифта ({fontSizeInPoints} вместо {requiredFontSize}) в тексте: {paragraph.InnerText}");
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Проверка выравнивания текста (полностью исключая заголовки)
        /// </summary>
        private bool CheckTextAlignment(string requiredAlignment, Body body, Gost gost)
        {
            var headers = GetAllSectionHeaders(body, gost);

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем все заголовки
                if (headers.Contains(paragraph))
                    continue;

                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                    continue;

                var justification = paragraphProperties.Justification;
                if (justification == null)
                    continue;

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
                    currentAlignment = "Both";
                }
                else
                {
                    currentAlignment = "Left";
                }

                if (currentAlignment != requiredAlignment)
                {
                    Console.WriteLine($"Ошибка выравнивания ({currentAlignment} вместо {requiredAlignment}) в тексте: {paragraph.InnerText}");
                    return false;
                }
            }
            return true;
        }



        /// <summary>
        /// Получает список обязательных разделов из строки
        /// </summary>
        private List<string> GetRequiredSectionsList(Gost gost)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections))
                return new List<string>();

            return gost.RequiredSections.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        /// <summary>
        /// Проверяет, является ли параграф заголовком раздела
        /// </summary>
        private bool IsSectionHeader(Paragraph paragraph, List<string> requiredSections)
        {
            if (!requiredSections.Any())
                return false;

            var text = paragraph.InnerText;
            return requiredSections.Any(section => text.Contains(section));
        }

        /// <summary>
        /// Проверка нумерации страниц
        /// </summary>
        private bool CheckPageNumbering(WordprocessingDocument wordDoc, bool requiredNumbering)
        {
            if (!requiredNumbering) return true;

            // Проверка верхних колонтитулов
            if (wordDoc.MainDocumentPart.HeaderParts != null)
            {
                foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                {
                    if (headerPart.Header.Descendants<SimpleField>()
                        .Any(f => f.Instruction?.Value?.Contains("PAGE") == true))
                    {
                        return true;
                    }
                }
            }

            // Проверка нижних колонтитулов
            if (wordDoc.MainDocumentPart.FooterParts != null)
            {
                foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                {
                    if (footerPart.Footer.Descendants<SimpleField>()
                        .Any(f => f.Instruction?.Value?.Contains("PAGE") == true))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Проверка отступа первой строки
        /// </summary>
        private bool CheckFirstLineIndent(double requiredFirstLineIndent, Body body)
        {
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var indent = paragraph.ParagraphProperties?.Indentation;
                if (indent?.FirstLine == null) continue;

                double firstLineIndentInCm = double.Parse(indent.FirstLine.Value) / 567.0;
                if (Math.Abs(firstLineIndentInCm - requiredFirstLineIndent) > 0.01)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Проверка межстрочного интервала
        /// </summary>
        private bool CheckLineSpacing(double requiredLineSpacing, Body body)
        {
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                if (spacing?.Line == null || spacing.LineRule != LineSpacingRuleValues.Auto)
                    continue;

                double lineSpacing = double.Parse(spacing.Line.Value) / 240.0;
                if (Math.Abs(lineSpacing - requiredLineSpacing) > 0.01)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Проверка полей документа
        /// </summary>
        private bool CheckMargins(
            double? requiredMarginTop, double? requiredMarginBottom,
            double? requiredMarginLeft, double? requiredMarginRight,
            Body body)
        {
            var pageMargin = body.Elements<SectionProperties>()
                .FirstOrDefault()?
                .Elements<PageMargin>()
                .FirstOrDefault();

            if (pageMargin == null) return false;

            // Преобразование в сантиметры (1 см = 567 twips)
            double marginTopInCm = pageMargin.Top.Value / 567.0;
            double marginBottomInCm = pageMargin.Bottom.Value / 567.0;
            double marginLeftInCm = pageMargin.Left.Value / 567.0;
            double marginRightInCm = pageMargin.Right.Value / 567.0;

            // Проверка с погрешностью 0.01 см
            if (requiredMarginTop.HasValue &&
                Math.Abs(marginTopInCm - requiredMarginTop.Value) > 0.01)
                return false;

            if (requiredMarginBottom.HasValue &&
                Math.Abs(marginBottomInCm - requiredMarginBottom.Value) > 0.01)
                return false;

            if (requiredMarginLeft.HasValue &&
                Math.Abs(marginLeftInCm - requiredMarginLeft.Value) > 0.01)
                return false;

            if (requiredMarginRight.HasValue &&
                Math.Abs(marginRightInCm - requiredMarginRight.Value) > 0.01)
                return false;

            return true;
        }

        /// <summary>
        /// Кнопка проверки на соответствие ГОСТу
        /// </summary>
        private async void Button_Click_SelectFile(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            try
            {
                int gostId = 1;
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
        private void Button_Click_LogOut(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            Close();
        }
    }
}