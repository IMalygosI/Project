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
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace GOST_Control
{
    /// <summary>
    /// Класс для проверок ГОСТа на соответствие заданным требованиям
    /// </summary>
    public partial class GOST_Сheck : Window
    {
        private readonly string _filePath; // Путь к файлу документа, который будет проверяться на соответствие ГОСТу
        private JsonGostService _gostService; // Сервис для работы с данными ГОСТов из JSON-файла
        private readonly Task _initializationTask; // Задача инициализации сервиса ГОСТов, запускаемая при создании экземпляра класса

        /// <summary>
        // Конструктор по умолчанию класса GOST_Сheck. Инициализирует компоненты окна.
        /// </summary>
        public GOST_Сheck()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Конструктор класса GOST_Сheck с параметром пути к файлу.
        /// Инициализирует компоненты окна и запускает асинхронную инициализацию сервиса ГОСТов.
        /// </summary>
        /// <param name="filePath"></param>
        public GOST_Сheck(string filePath)
        {
            InitializeComponent();

            _filePath = filePath;

            _initializationTask = InitializeAsync();
        }

        /// <summary>
        /// Асинхронно инициализирует сервис для работы с ГОСТами из JSON-файла.
        /// Обновляет UI с информацией о пути к файлу или ошибкой загрузки.
        /// </summary>
        /// <returns></returns>
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
        /// Основной метод проверки-"вызова проверок" документа на соответствие ГОСТу
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

            // Получение ГОСТ из JSON файла
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

                        bool stylesValid = CheckStyleFonts(wordDoc, gost);
                        if (!stylesValid)
                        {
                            ErrorControlFont.Text = "Ошибка в стилях документа!";
                            ErrorControlFont.Foreground = Brushes.Red;
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
                            pageNumberingValid = CheckPageNumbering(wordDoc, gost.PageNumbering.Value, gost.PageNumberingAlignment, gost.PageNumberingPosition);
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
                return true;

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

                // Ищем параграфы, содержащие название раздела (более гибкое сравнение)
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    var text = paragraph.InnerText.Trim();
                    if (text.IndexOf(section, StringComparison.OrdinalIgnoreCase) >= 0)
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
                                        invalidSections.Add($"{section} (неверный размер шрифта: {fontSizeValue})");
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
                        }

                        // Проверяем выравнивание заголовка
                        if (!string.IsNullOrEmpty(gost.HeaderAlignment))
                        {
                            var justification = paragraph.ParagraphProperties?.Justification;
                            var currentAlignment = GetAlignmentString(justification);

                            if (currentAlignment != gost.HeaderAlignment)
                            {
                                invalidSections.Add($"{section} (неверное выравнивание: {currentAlignment})");
                                sectionValid = false;
                                break;
                            }
                        }

                        // Если нашли раздел и он соответствует требованиям, переходим к следующему
                        if (sectionValid) break;
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


        /// <summary>
        /// Проверка типа шрифта (исключая заголовки)
        /// </summary>
        private bool CheckFontName(string requiredFontName, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                if (headerTexts.Contains(paragraph.InnerText.Trim()) || IsEmptyParagraph(paragraph))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    var fontName = run.RunProperties?.RunFonts?.Ascii?.Value;
                    if (fontName != null && fontName != requiredFontName)
                    {
                        Dispatcher.UIThread.Post(() => {
                            ErrorControlFont.Text = "Неверный шрифт в основном тексте";
                            ErrorControlFont.Foreground = Brushes.Red;
                        });
                        isValid = false;
                        break;
                    }
                }
                if (!isValid) break;
            }
            return isValid;
        }


        /// <summary>
        /// Проверка размера шрифта (исключая заголовки)
        /// </summary>
        private bool CheckFontSize(double requiredFontSize, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                if (headerTexts.Contains(paragraph.InnerText.Trim()) || IsEmptyParagraph(paragraph))
                    continue;

                var text = paragraph.InnerText.Trim();
                if (headerTexts.Contains(text)) continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    var fontSize = run.RunProperties?.FontSize;
                    if (fontSize != null)
                    {
                        double fontSizeValue = double.Parse(fontSize.Val.Value) / 2;
                        if (Math.Abs(fontSizeValue - requiredFontSize) > 0.1)
                        {
                            Dispatcher.UIThread.Post(() => {
                                ErrorControlFontSize.Text = "Ошибка: неверный размер шрифта";
                                ErrorControlFontSize.Foreground = Brushes.Red;
                            });
                            isValid = false;
                            break;
                        }
                    }
                }
                if (!isValid) break;
            }
            return isValid;
        }


        /// <summary>
        /// Проверка выравнивания текста (исключая заголовки)
        /// </summary>
        private bool CheckTextAlignment(string requiredAlignment, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем: 1) заголовки, 2) пустые параграфы
                if (headerTexts.Contains(paragraph.InnerText.Trim()) || IsEmptyParagraph(paragraph))
                    continue;

                var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);

                if (currentAlignment != requiredAlignment)
                {
                    Dispatcher.UIThread.Post(() => {
                        ErrorControlViravnivanie.Text = $"Ошибка: выравнивание {currentAlignment} (требуется {requiredAlignment})";
                        ErrorControlViravnivanie.Foreground = Brushes.Red;
                    });
                    isValid = false;
                    break;
                }
            }
            return isValid;
        }

        /// <summary>
        /// Получает тексты заголовков из тела документа на основе обязательных разделов ГОСТа
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private HashSet<string> GetHeaderTexts(Body body, Gost gost)
        {
            var requiredSections = GetRequiredSectionsList(gost);
            var headerTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var text = paragraph.InnerText.Trim();
                foreach (var section in requiredSections)
                {
                    if (text.IndexOf(section, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        headerTexts.Add(text);
                        break;
                    }
                }
            }
            return headerTexts;
        }

        /// <summary>
        /// Проверяет соответствие шрифтов в стилях документа требованиям ГОСТа
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckStyleFonts(WordprocessingDocument doc, Gost gost)
        {
            var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart == null) return true;

            foreach (var style in stylesPart.Styles.Elements<Style>())
            {
                var justification = style.StyleParagraphProperties?.Justification;
                if (justification != null)
                {
                    string alignment = GetAlignmentString(justification);
                    string requiredAlignment = style.Type == StyleValues.Paragraph ?
                        gost.TextAlignment : gost.HeaderAlignment;

                    if (alignment != requiredAlignment)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Преобразует объект выравнивания в строковое представление
        /// </summary>
        /// <param name="justification"></param>
        /// <returns></returns>
        private string GetAlignmentString(Justification justification)
        {
            if (justification == null) return "Left";

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
        /// Проверка нумерации страниц и её расположения
        /// </summary>
        /// <summary>
        /// Проверка нумерации страниц и её расположения
        /// </summary>
        private bool CheckPageNumbering(WordprocessingDocument wordDoc, bool requiredNumbering,
                                      string requiredAlignment, string requiredPosition)
        {
            if (!requiredNumbering) return true;

            bool hasCorrectNumbering = false;
            bool hasExtraNumbering = false;
            string actualCorrectPosition = "";
            string actualCorrectAlignment = "";
            List<string> extraNumberings = new List<string>();

            // Проверяем все колонтитулы на наличие номеров страниц
            if (wordDoc.MainDocumentPart.HeaderParts != null)
            {
                foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                {
                    foreach (var paragraph in headerPart.Header.Elements<Paragraph>())
                    {
                        var pageField = paragraph.Descendants<SimpleField>()
                            .FirstOrDefault(f => f.Instruction?.Value?.Contains("PAGE") == true);

                        if (pageField != null)
                        {
                            var justification = paragraph.ParagraphProperties?.Justification;
                            string alignment = GetAlignmentString(justification);
                            string position = "Top";

                            // Проверяем, соответствует ли текущая нумерация требованиям
                            bool positionMatch = string.IsNullOrEmpty(requiredPosition) ||
                                               position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase);
                            bool alignmentMatch = string.IsNullOrEmpty(requiredAlignment) ||
                                                alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase);

                            if (positionMatch && alignmentMatch)
                            {
                                hasCorrectNumbering = true;
                                actualCorrectPosition = position;
                                actualCorrectAlignment = alignment;
                            }
                            else
                            {
                                hasExtraNumbering = true;
                                extraNumberings.Add($"{position}, {alignment}");
                            }
                        }
                    }
                }
            }

            if (wordDoc.MainDocumentPart.FooterParts != null)
            {
                foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                {
                    foreach (var paragraph in footerPart.Footer.Elements<Paragraph>())
                    {
                        var pageField = paragraph.Descendants<SimpleField>()
                            .FirstOrDefault(f => f.Instruction?.Value?.Contains("PAGE") == true);

                        if (pageField != null)
                        {
                            var justification = paragraph.ParagraphProperties?.Justification;
                            string alignment = GetAlignmentString(justification);
                            string position = "Bottom";

                            // Проверяем, соответствует ли текущая нумерация требованиям
                            bool positionMatch = string.IsNullOrEmpty(requiredPosition) ||
                                               position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase);
                            bool alignmentMatch = string.IsNullOrEmpty(requiredAlignment) ||
                                                alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase);

                            if (positionMatch && alignmentMatch)
                            {
                                hasCorrectNumbering = true;
                                actualCorrectPosition = position;
                                actualCorrectAlignment = alignment;
                            }
                            else
                            {
                                hasExtraNumbering = true;
                                extraNumberings.Add($"{position}, {alignment}");
                            }
                        }
                    }
                }
            }

            // Формируем сообщение для пользователя
            if (!hasCorrectNumbering && !hasExtraNumbering)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlNumberPage.Text = "Нумерация страниц отсутствует";
                    ErrorControlNumberPage.Foreground = Brushes.Red;
                });
                return false;
            }

            if (hasCorrectNumbering && !hasExtraNumbering)
            {
                Dispatcher.UIThread.Post(() => {
                    string message = string.IsNullOrEmpty(requiredAlignment) && string.IsNullOrEmpty(requiredPosition)
                        ? "Нумерация страниц присутствует"
                        : $"Нумерация соответствует ({actualCorrectPosition}, {actualCorrectAlignment})";

                    ErrorControlNumberPage.Text = message;
                    ErrorControlNumberPage.Foreground = Brushes.Green;
                });
                return true;
            }

            if (hasCorrectNumbering && hasExtraNumbering)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlNumberPage.Text = $"Нумерация соответствует ({actualCorrectPosition}, {actualCorrectAlignment}), " +
                                                $"но есть лишняя нумерация в: {string.Join("; ", extraNumberings)}";
                    ErrorControlNumberPage.Foreground = Brushes.Red;
                });
                return false;
            }

            if (!hasCorrectNumbering && hasExtraNumbering)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlNumberPage.Text = $"Не найдена требуемая нумерация, но есть лишняя в: {string.Join("; ", extraNumberings)}";
                    ErrorControlNumberPage.Foreground = Brushes.Red;
                });
                return false;
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
        /// Проверяет, является ли параграф пустым. 
        /// Для того чтобы пустые места в документе не вызывали ошибок!
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsEmptyParagraph(Paragraph paragraph)
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                foreach (var text in run.Elements<Text>())
                {
                    if (!string.IsNullOrWhiteSpace(text.Text))
                        return false;
                }

                // Проверяем специальные символы (например, разрывы строк)
                if (run.Elements<Break>().Any() || run.Elements<TabChar>().Any())
                    return false;
            }
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