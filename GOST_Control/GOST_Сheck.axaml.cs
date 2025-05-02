using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GOST_Control;
using Xceed.Words.NET;
using Avalonia.Layout;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using System.Diagnostics;
using DocumentFormat.OpenXml;

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
                        bool paperSizeValid = true;
                        bool orientationValid = true;
                        bool tocValid = true;
                        bool bulletedListsValid = true;
                        bool sectionsValid = true;

                        // Список для хранения ошибок
                        var errors = new List<string>();

                        // Проверка типа шрифта (игнорируя заголовки разделов)
                        if (!string.IsNullOrEmpty(gost.FontName))
                        {
                            fontNameValid = CheckFontName(gost.FontName, body, gost);
                            ErrorControlFont.Text = fontNameValid ? "Тип шрифта соответствует ГОСТу." : "Тип шрифта не соответствует.";
                            ErrorControlFont.Foreground = fontNameValid ? Brushes.Green : Brushes.Red;
                            if (!fontNameValid) errors.Add("Тип шрифта не соответствует ГОСТу");
                        }

                        // Проверка размера шрифта (игнорируя заголовки разделов)
                        if (gost.FontSize.HasValue)
                        {
                            fontSizeValid = CheckFontSize(gost.FontSize.Value, body, gost);
                            ErrorControlFontSize.Text = fontSizeValid ? "Размер шрифта соответствует ГОСТу!" : "Размер шрифта не соответствует!";
                            ErrorControlFontSize.Foreground = fontSizeValid ? Brushes.Green : Brushes.Red;
                            if (!fontSizeValid) errors.Add("Размер шрифта не соответствует ГОСТу");
                        }

                        // Проверка полей документа
                        if (gost.MarginTop.HasValue || gost.MarginBottom.HasValue || gost.MarginLeft.HasValue || gost.MarginRight.HasValue)
                        {
                            marginsValid = CheckMargins(gost.MarginTop, gost.MarginBottom, gost.MarginLeft, gost.MarginRight, body);
                            ErrorControlMargins.Text = marginsValid ? "Поля документа соответствуют ГОСТу." : "Поля документа не соответствуют ГОСТу.";
                            ErrorControlMargins.Foreground = marginsValid ? Brushes.Green : Brushes.Red;
                            if (!marginsValid) errors.Add("Поля документа не соответствуют ГОСТу");
                        }

                        // Проверка межстрочного интервала
                        if (gost.LineSpacing.HasValue)
                        {
                            lineSpacingValid = CheckLineSpacing(gost.LineSpacing.Value, body, gost);
                            ErrorControlMnochitel.Text = lineSpacingValid ? "Межстрочный интервал соответствует ГОСТу." : "Межстрочный интервал не соответствует ГОСТу.";
                            ErrorControlMnochitel.Foreground = lineSpacingValid ? Brushes.Green : Brushes.Red;
                            if (!lineSpacingValid) errors.Add("Межстрочный интервал не соответствует ГОСТу");
                        }
                                
                        // Проверка отступа первой строки
                        if (gost.FirstLineIndent.HasValue)
                        {
                            firstLineIndentValid = CheckFirstLineIndent(gost.FirstLineIndent.Value, body, gost); // Добавлен параметр gost
                            ErrorControlFirstLineIndent.Text = firstLineIndentValid ? "Отступ соответствует ГОСТу." : "Отступ не соответствует ГОСТу.";
                            ErrorControlFirstLineIndent.Foreground = firstLineIndentValid ? Brushes.Green : Brushes.Red;
                            if (!firstLineIndentValid) errors.Add("Отступ первой строки не соответствует ГОСТу");
                        }

                        // Проверка выравнивания текста (игнорируя заголовки разделов)
                        if (!string.IsNullOrEmpty(gost.TextAlignment))
                        {
                            textAlignmentValid = CheckTextAlignment(gost.TextAlignment, body, gost);
                            ErrorControlViravnivanie.Text = textAlignmentValid ? "Выравнивание текста соответствует ГОСТу." : "Выравнивание текста не соответствует ГОСТу.";
                            ErrorControlViravnivanie.Foreground = textAlignmentValid ? Brushes.Green : Brushes.Red;
                            if (!textAlignmentValid) errors.Add("Выравнивание текста не соответствует ГОСТу");
                        }

                        // Проверка нумерации страниц
                        if (gost.PageNumbering.HasValue)
                        {
                            pageNumberingValid = CheckPageNumbering(wordDoc, gost.PageNumbering.Value, gost.PageNumberingAlignment, gost.PageNumberingPosition);
                            ErrorControlNumberPage.Text = pageNumberingValid ? "Нумерация страниц соответствует ГОСТу." : "Нумерация страниц не соответствует ГОСТу.";
                            ErrorControlNumberPage.Foreground = pageNumberingValid ? Brushes.Green : Brushes.Red;
                            if (!pageNumberingValid) errors.Add("Нумерация страниц не соответствует ГОСТу");
                        }
                        else
                        {
                            ErrorControlNumberPage.Text = "Нумерация страниц не требуется.";
                            ErrorControlNumberPage.Foreground = Brushes.Gray;
                        }

                        // Проверка обязательных разделов (Введение, Заключение)
                        if (!string.IsNullOrEmpty(gost.RequiredSections))
                        {
                            sectionsValid = CheckRequiredSections(gost, body);
                            if (!sectionsValid) errors.Add("Отсутствуют обязательные разделы");
                        }

                        // Проверка формата
                        if (gost.PaperWidthMm.HasValue && gost.PaperHeightMm.HasValue)
                        {
                            paperSizeValid = CheckPaperSize(wordDoc, gost);
                            if (!paperSizeValid) errors.Add("Размер бумаги не соответствует ГОСТу");
                        }

                        // Проверка Ориентации
                        if (!string.IsNullOrEmpty(gost.PageOrientation))
                        {
                            orientationValid = CheckPageOrientation(wordDoc, gost);
                            if (!orientationValid) errors.Add("Ориентация страницы не соответствует ГОСТу");
                        }

                        // Оглавление
                        if (gost.RequireTOC.HasValue && gost.RequireTOC.Value)
                        {
                            tocValid = CheckTableOfContents(wordDoc, gost);
                            if (!tocValid) errors.Add("Оглавление не соответствует ГОСТу");
                        }

                        // Проверка маркированных списков
                        if (gost.RequireBulletedLists.HasValue && gost.RequireBulletedLists.Value)
                        {
                            bulletedListsValid = CheckBulletedLists(body, gost);
                            if (!bulletedListsValid) errors.Add("Маркированные списки не соответствуют ГОСТу");
                        }

                        // Общий результат проверки
                        if (fontNameValid && fontSizeValid && marginsValid && lineSpacingValid && firstLineIndentValid && textAlignmentValid && pageNumberingValid && 
                            sectionsValid && paperSizeValid && orientationValid && tocValid && bulletedListsValid )
                        {
                            GostControl.Text = "Документ соответствует ГОСТу.";
                            GostControl.Foreground = Brushes.Green;
                        }
                        else
                        {
                            GostControl.Text = "Документ не соответствует ГОСТу:";
                            GostControl.Foreground = Brushes.Red;

                            // Создаем документ с ошибками
                            await CreateErrorReportDocument(wordDoc, gost, errors, filePath);
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
        /// Проверка маркированных списков с выделением ошибок
        /// </summary>
        private bool CheckBulletedLists(Body body, Gost gost)
        {
            if (!(gost.RequireBulletedLists ?? false))
            {
                UpdateBulletedListsUI(new List<string>(), true, false);
                return true;
            }

            var errors = new List<string>();
            bool hasLists = false;
            bool listsValid = true;
            var listLevels = new Dictionary<int, int>();
            bool hasMultiLevelLists = false;

            // Сначала определяем, есть ли в документе многоуровневые списки
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                if (!IsListItem(paragraph) || IsEmptyParagraph(paragraph))
                    continue;

                int level = GetListLevel(paragraph, gost);
                if (level > 1)
                {
                    hasMultiLevelLists = true;
                    break;
                }
            }

            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                if (!IsListItem(paragraph) || IsEmptyParagraph(paragraph))
                    continue;

                hasLists = true;
                bool paragraphHasError = false;
                var runsWithText = paragraph.Elements<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)).ToList();

                // Определяем уровень списка
                int listLevel = GetListLevel(paragraph, gost);
                if (!listLevels.ContainsKey(listLevel))
                    listLevels[listLevel] = 0;
                listLevels[listLevel]++;

                // Проверка межстрочного интервала
                if (gost.BulletLineSpacing.HasValue)
                {
                    var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    if (spacing?.Line == null || Math.Abs(double.Parse(spacing.Line.Value) / 240.0 - gost.BulletLineSpacing.Value) > 0.01)
                    {
                        errors.Add($"Межстрочный интервал не соответствует {gost.BulletLineSpacing.Value}");
                        paragraphHasError = true;
                    }
                }

                // Проверка отступов
                double? requiredIndent = null;

                // Если есть многоуровневые списки ИЛИ явно указаны отступы для уровней
                if (hasMultiLevelLists ||
                    (listLevel == 1 && gost.ListLevel1Indent.HasValue) ||
                    (listLevel == 2 && gost.ListLevel2Indent.HasValue) ||
                    (listLevel == 3 && gost.ListLevel3Indent.HasValue))
                {
                    requiredIndent = listLevel switch
                    {
                        1 => gost.ListLevel1Indent ?? gost.ListHangingIndent,
                        2 => gost.ListLevel2Indent ?? gost.ListHangingIndent,
                        3 => gost.ListLevel3Indent ?? gost.ListHangingIndent,
                        _ => gost.ListHangingIndent
                    };
                }
                else
                {
                    // Для простых списков используем общий отступ
                    requiredIndent = gost.ListHangingIndent;
                }

                if (requiredIndent.HasValue)
                {
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    bool indentValid = false;

                    if (indent?.Hanging != null && Math.Abs(double.Parse(indent.Hanging.Value) / 567.0 - requiredIndent.Value) <= 0.05)
                    {
                        indentValid = true;
                    }
                    else if (indent?.FirstLine != null && Math.Abs(double.Parse(indent.FirstLine.Value) / 567.0 - requiredIndent.Value) <= 0.05)
                    {
                        indentValid = true;
                    }

                    if (!indentValid)
                    {
                        errors.Add($"Отступ не соответствует {requiredIndent.Value} см");
                        paragraphHasError = true;
                    }
                }

                // Проверка формата нумерации только для нумерованных списков
                if (IsNumberedList(paragraph))
                {
                    // Проверяем формат только если есть многоуровневые списки ИЛИ указаны форматы для уровней
                    if (hasMultiLevelLists ||
                        (listLevel == 1 && !string.IsNullOrEmpty(gost.ListLevel1NumberFormat)) ||
                        (listLevel == 2 && !string.IsNullOrEmpty(gost.ListLevel2NumberFormat)) ||
                        (listLevel == 3 && !string.IsNullOrEmpty(gost.ListLevel3NumberFormat)))
                    {
                        string? requiredFormat = listLevel switch
                        {
                            1 => gost.ListLevel1NumberFormat,
                            2 => gost.ListLevel2NumberFormat,
                            3 => gost.ListLevel3NumberFormat,
                            _ => null
                        };

                        if (!string.IsNullOrEmpty(requiredFormat))
                        {
                            var firstRunText = runsWithText.FirstOrDefault()?.InnerText.Trim();
                            if (firstRunText != null && !CheckNumberFormat(firstRunText, requiredFormat))
                            {
                                errors.Add($"Неверный формат нумерации '{firstRunText}' (требуется '{requiredFormat}')");
                                paragraphHasError = true;
                            }
                        }
                    }
                }

                // Проверка шрифта и размера
                foreach (var run in runsWithText)
                {
                    bool runHasError = false;

                    if (!string.IsNullOrEmpty(gost.BulletFontName))
                    {
                        var font = run.RunProperties?.RunFonts?.Ascii?.Value;
                        if (font != null && font != gost.BulletFontName)
                        {
                            errors.Add($"Шрифт списка '{font}' вместо '{gost.BulletFontName}'");
                            runHasError = true;
                        }
                    }

                    if (gost.BulletFontSize.HasValue)
                    {
                        var fontSize = run.RunProperties?.FontSize;
                        if (fontSize != null)
                        {
                            double actualSize = double.Parse(fontSize.Val.Value) / 2;
                            if (Math.Abs(actualSize - gost.BulletFontSize.Value) > 0.1)
                            {
                                errors.Add($"Размер шрифта списка {actualSize}pt вместо {gost.BulletFontSize.Value}pt");
                                runHasError = true;
                            }
                        }
                    }

                    if (runHasError)
                    {
                        paragraphHasError = true;
                        HighlightRun(run);
                    }
                }

                if (paragraphHasError)
                {
                    listsValid = false;
                    HighlightParagraph(paragraph);
                }
            }

            // Проверяем наличие списков разных уровней только если в документе есть многоуровневые списки
            if (hasMultiLevelLists)
            {
                if (gost.ListLevel1Indent.HasValue && !listLevels.ContainsKey(1))
                {
                    errors.Add("Отсутствуют списки 1-го уровня");
                    listsValid = false;
                }

                if (gost.ListLevel2Indent.HasValue && !listLevels.ContainsKey(2))
                {
                    errors.Add("Отсутствуют списки 2-го уровня");
                    listsValid = false;
                }

                if (gost.ListLevel3Indent.HasValue && !listLevels.ContainsKey(3))
                {
                    errors.Add("Отсутствуют списки 3-го уровня");
                    listsValid = false;
                }
            }

            if (!hasLists && gost.RequireBulletedLists == true)
            {
                errors.Add("Отсутствуют обязательные маркированные списки");
                listsValid = false;
            }

            UpdateBulletedListsUI(errors.Distinct().ToList(), listsValid, hasLists);
            return listsValid;
        }

        // Вспомогательные методы для выделения
        private int GetListLevel(Paragraph paragraph, Gost gost)
        {
            var numberingProps = paragraph.ParagraphProperties?.NumberingProperties;
            if (numberingProps?.NumberingLevelReference?.Val?.Value != null)
            {
                return numberingProps.NumberingLevelReference.Val.Value + 1; // Уровни обычно 0-based
            }

            // Эвристика для определения уровня по отступу
            var indent = paragraph.ParagraphProperties?.Indentation;
            if (indent?.Left != null)
            {
                double leftIndent = double.Parse(indent.Left.Value) / 567.0; // в см

                if (gost.ListLevel3Indent.HasValue && leftIndent >= gost.ListLevel3Indent.Value - 0.5)
                    return 3;
                if (gost.ListLevel2Indent.HasValue && leftIndent >= gost.ListLevel2Indent.Value - 0.5)
                    return 2;
            }

            return 1; // По умолчанию считаем первым уровнем
        }

        private bool IsNumberedList(Paragraph paragraph)
        {
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun == null) return false;

            var text = firstRun.InnerText.Trim();

            // Проверяем форматы нумерации: 1., 1), a., a), I., и т.д.
            return Regex.IsMatch(text, @"^(\d+[\.\)]|[a-z]\)|[A-Z]\.|I+\.|V+\.|X+\.)");
        }

        private bool CheckNumberFormat(string text, string requiredFormat)
        {
            // Простая проверка соответствия формата
            if (requiredFormat.EndsWith(".") && text.EndsWith("."))
                return true;
            if (requiredFormat.EndsWith(")") && text.EndsWith(")"))
                return true;

            // Более сложные проверки могут быть добавлены здесь

            return false;
        }
        private void HighlightParagraph(Paragraph p)
        {
            foreach (var run in p.Elements<Run>())
            {
                HighlightRun(run);
            }
        }

        private void HighlightRun(Run run, IBrush? highlightColor = null)
        {
            run.RunProperties ??= new RunProperties();
            // Удаляем все существующие выделения
            run.RunProperties.RemoveAllChildren<Highlight>();
            // Добавляем красное выделение фона
            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
            run.RunProperties.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Color>();
        }

        private void HighlightTocErrorsInReport(WordprocessingDocument doc, Gost gost)
        {
            double requiredSpacing = gost.LineSpacing ?? 1.5;
            string requiredFont = gost.FontName ?? "Times New Roman";
            double requiredSize = gost.FontSize ?? 14.0;

            // Находим автоматическое оглавление
            var tocField = doc.MainDocumentPart.Document.Body.Descendants<FieldCode>().FirstOrDefault(f => f.Text.Contains(" TOC ") || f.Text.Contains("TOC \\"));

            if (tocField == null) return;

            var tocContainer = tocField.Ancestors<Paragraph>().FirstOrDefault()?.Parent;
            if (tocContainer == null) return;

            bool hasSpacingError = false;

            // Сначала проверяем межстрочный интервал во всем оглавлении
            foreach (var para in tocContainer.Descendants<Paragraph>())
            {
                if (IsEmptyParagraph(para)) continue;

                var spacing = para.ParagraphProperties?.SpacingBetweenLines;
                if (spacing?.Line != null)
                {
                    double actualSpacing = double.Parse(spacing.Line.Value) / 240.0;
                    if (Math.Abs(actualSpacing - requiredSpacing) > 0.01)
                    {
                        hasSpacingError = true;
                        break;
                    }
                }
                else
                {
                    hasSpacingError = true;
                    break;
                }
            }

            // Обрабатываем все параграфы
            foreach (var para in tocContainer.Descendants<Paragraph>())
            {
                if (IsEmptyParagraph(para)) continue;

                // Если есть ошибка интервала - выделяем весь параграф
                if (hasSpacingError)
                {
                    HighlightEntireParagraph(para);
                }
                else
                {
                    // Иначе проверяем шрифт и размер для каждого Run
                    foreach (var run in para.Descendants<Run>())
                    {
                        if (!string.IsNullOrWhiteSpace(run.InnerText))
                        {
                            bool hasFontError = false;
                            var font = run.RunProperties?.RunFonts?.Ascii?.Value;
                            var fontSize = run.RunProperties?.FontSize?.Val?.Value;

                            if ((font != null && font != requiredFont) ||
                                (fontSize != null && Math.Abs(double.Parse(fontSize) / 2 - requiredSize) > 0.1))
                            {
                                hasFontError = true;
                            }

                            if (hasFontError)
                            {
                                HighlightRun(run);
                            }
                        }
                    }
                }
            }

            doc.MainDocumentPart.Document.Save();
        }

        private void HighlightEntireParagraph(Paragraph para)
        {
            foreach (var run in para.Descendants<Run>())
            {
                HighlightRun(run);
            }
        }

        private bool IsEmptyParagraph(Paragraph p)
        {
            return !p.Descendants<Run>().Any(r => !string.IsNullOrWhiteSpace(r.InnerText));
        }

        /// <summary>
        /// Проверка оглавления
        /// </summary>
        private bool CheckTableOfContents(WordprocessingDocument doc, Gost gost)
        {
            // НОВОЕ: Приоритет TocFontName > FontName
            string requiredFont = gost.TocFontName ?? gost.FontName ?? "Times New Roman";

            // НОВОЕ: Приоритет TocFontSize > FontSize
            double requiredSize = gost.TocFontSize ?? gost.FontSize ?? 14.0;

            // НОВОЕ: Приоритет TocLineSpacing > LineSpacing
            double requiredSpacing = gost.TocLineSpacing ?? gost.LineSpacing ?? 1.5;

            bool hasErrors = false;
            var errorDetails = new List<string>();

            var tocField = doc.MainDocumentPart.Document.Body
                .Descendants<FieldCode>()
                .FirstOrDefault(f => f.Text.Contains(" TOC ") || f.Text.Contains("TOC \\"));

            if (tocField == null)
            {
                ShowTocError("Автоматическое оглавление не найдено! Создайте через 'Ссылки → Оглавление'");
                return false;
            }

            var tocContainer = tocField.Ancestors<Paragraph>().FirstOrDefault()?.Parent;
            if (tocContainer == null)
            {
                ShowTocError("Не удалось определить границы оглавления");
                return false;
            }

            foreach (var para in tocContainer.Descendants<Paragraph>())
            {
                if (IsEmptyParagraph(para)) continue;

                var spacing = para.ParagraphProperties?.SpacingBetweenLines;
                double actualSpacing = spacing?.Line != null ? double.Parse(spacing.Line.Value) / 240.0 : 0;
                if (Math.Abs(actualSpacing - requiredSpacing) > 0.01)
                {
                    errorDetails.Add($"Неверный межстрочный интервал: {actualSpacing:0.##} (требуется {requiredSpacing})");
                    hasErrors = true;
                }

                foreach (var run in para.Descendants<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)))
                {
                    bool runHasError = false;

                    var font = run.RunProperties?.RunFonts?.Ascii?.Value;
                    if (font != requiredFont)
                    {
                        errorDetails.Add($"Неверный шрифт: '{font}' (требуется '{requiredFont}')");
                        runHasError = true;
                    }

                    var size = run.RunProperties?.FontSize?.Val?.Value;
                    if (size != null)
                    {
                        double actualSize = double.Parse(size) / 2;
                        if (Math.Abs(actualSize - requiredSize) > 0.1)
                        {
                            errorDetails.Add($"Неверный размер шрифта: {actualSize:0.##}pt (требуется {requiredSize}pt)");
                            runHasError = true;
                        }
                    }

                    if (runHasError) hasErrors = true;
                }
            }

            if (hasErrors)
            {
                string errorMessage = $"Ошибки в оглавлении:\n{string.Join("\n", errorDetails.Distinct().Take(3))}";
                if (errorDetails.Count > 3) errorMessage += $"\n...и ещё {errorDetails.Count - 3} ошибок";
                ShowTocError(errorMessage);
            }
            else
            {
                ShowTocSuccess("Оглавление соответствует требованиям ГОСТ");
            }

            return !hasErrors;
        }

        // Вспомогательные методы
        private void ShowTocError(string message)
        {
            Dispatcher.UIThread.Post(() => {
                ErrorControlTOC.Text = message;
                ErrorControlTOC.Foreground = Brushes.Red;
            });
        }

        private void ShowTocSuccess(string message)
        {
            Dispatcher.UIThread.Post(() => {
                ErrorControlTOC.Text = message;
                ErrorControlTOC.Foreground = Brushes.Green;
            });
        }

        private async Task CreateErrorReportDocument(WordprocessingDocument originalDoc, Gost gost, List<string> errors, string originalFilePath)
        {
            // 1. Сначала показываем диалог подтверждения
            var mainWindow = (Application.Current.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;

            bool confirmSave = await ShowConfirmationDialog(mainWindow,
                "Сохранить отчет об ошибках",
                "Документ не соответствует ГОСТу. Хотите сохранить отчет с выделенными ошибками?");

            if (!confirmSave) return;

            // 2. Создаем временную копию с правами на запись
            string tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.docx");
            File.Copy(originalFilePath, tempPath, true);

            // 3. Настройка диалога сохранения
            var saveDialog = new SaveFileDialog
            {
                Title = "Сохранить отчет об ошибках",
                InitialFileName = $"Ошибки_{Path.GetFileNameWithoutExtension(originalFilePath)}",
                DefaultExtension = ".docx",
                Filters = new List<FileDialogFilter> { new() { Name = "Word Documents", Extensions = { "docx" } } }
            };

            var saveResult = await saveDialog.ShowAsync(mainWindow);
            if (string.IsNullOrEmpty(saveResult)) return;

            // 4. Обработка документа
            using (var errorDoc = WordprocessingDocument.Open(tempPath, true))
            {
                var body = errorDoc.MainDocumentPart.Document.Body;

                // Выделение ошибок в оглавлении
                HighlightTocErrorsInReport(errorDoc, gost);

                // Выделение обязательных разделов
                HighlightRequiredSections(body, gost, errors);

                // Выделение ошибок форматирования
                HighlightFormattingErrors(body, gost, errors);

                // Выделение обычного текста с ошибками
                HighlightErrorText(body, errors);

                //  Выделение ошибок в Маркированных списках
                CheckBulletedLists(body, gost);

                // Выделение ошибок в нумерации страниц
                if (gost.PageNumbering.HasValue)
                {
                    CheckPageNumbering(errorDoc, gost.PageNumbering.Value, gost.PageNumberingAlignment, gost.PageNumberingPosition);
                }

                errorDoc.MainDocumentPart.Document.Save();
            }

            // 5. Сохранение и открытие
            try
            {
                File.Copy(tempPath, saveResult, true);
                Process.Start(new ProcessStartInfo(saveResult) { UseShellExecute = true });
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }

        /// <summary>
        /// Новый метод для выделения ошибок форматирования
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <param name="errors"></param>
        private void HighlightFormattingErrors(Body body, Gost gost, List<string> errors)
        {
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                if (IsEmptyParagraph(paragraph)) continue;

                // Пропускаем специальные элементы
                if (IsTocParagraph(paragraph) || IsHeaderParagraph(paragraph, gost) || IsListItem(paragraph))
                    continue;

                // Проверка только для обычного текста
                bool hasError = false;

                // 1. Проверка межстрочного интервала
                if (gost.LineSpacing.HasValue)
                {
                    var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    if (spacing?.Line == null ||
                        Math.Abs(double.Parse(spacing.Line.Value) / 240.0 - gost.LineSpacing.Value) > 0.01)
                    {
                        hasError = true;
                        errors.Add($"Неверный межстрочный интервал в тексте: '{paragraph.InnerText.Trim()}'");
                    }
                }

                // 2. Проверка отступа первой строки
                if (gost.FirstLineIndent.HasValue)
                {
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    if (indent?.FirstLine == null ||
                        Math.Abs(double.Parse(indent.FirstLine.Value) / 567.0 - gost.FirstLineIndent.Value) > 0.05)
                    {
                        hasError = true;
                        errors.Add($"Неверный отступ в тексте: '{paragraph.InnerText.Trim()}'");
                    }
                }

                // 3. Проверка выравнивания
                if (!string.IsNullOrEmpty(gost.TextAlignment))
                {
                    var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);
                    if (currentAlignment != gost.TextAlignment)
                    {
                        hasError = true;
                        errors.Add($"Неверное выравнивание текста: '{paragraph.InnerText.Trim()}'");
                    }
                }

                // Выделяем весь параграф при наличии ошибок
                if (hasError)
                {
                    HighlightWholeParagraph(paragraph);
                }
            }
        }

        /// <summary>
        /// Метод для проверки, является ли параграф заголовком
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool IsHeaderParagraph(Paragraph paragraph, Gost gost)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections)) return false;

            var requiredSections = gost.RequiredSections.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s));

            var paragraphText = paragraph.InnerText.Trim();

            return requiredSections.Any(section =>
                paragraphText.IndexOf(section, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        /// <summary>
        /// Улучшенный метод выделения текста с ошибками
        /// </summary>
        /// <param name="body"></param>
        /// <param name="errors"></param>
        private void HighlightErrorText(Body body, List<string> errors)
        {
            foreach (var error in errors)
            {
                // Ищем текст после последнего '|' как текст для поиска
                var errorParts = error.Split('|');
                if (errorParts.Length == 0) continue;

                var errorText = errorParts.Last().Trim();
                if (string.IsNullOrEmpty(errorText)) continue;

                // Ищем во всех параграфах
                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    if (paragraph.InnerText.Contains(errorText))
                    {
                        // Выделяем весь параграф, если нашли совпадение
                        HighlightParagraph(paragraph);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Метод для выделения обязательных разделов и их ошибок форматирования
        /// </summary>
        private void HighlightRequiredSections(Body body, Gost gost, List<string> errors)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections))
                return;

            var requiredSections = gost.RequiredSections.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            foreach (var section in requiredSections)
            {
                bool sectionFound = false;
                bool hasFormattingErrors = false;

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    // Пропускаем специальные элементы
                    if (ShouldSkipHeaderCheck(paragraph))
                        continue;

                    var paragraphText = paragraph.InnerText.Trim();

                    // Точное сравнение с учетом возможных номеров (например "1 Введение")
                    if (IsSectionMatch(paragraphText, section))
                    {
                        sectionFound = true;
                        hasFormattingErrors = CheckHeaderFormatting(paragraph, section, gost, errors);

                        if (hasFormattingErrors)
                        {
                            HighlightParagraph(paragraph);
                        }
                        break;
                    }
                }

                if (!sectionFound)
                {
                    errors.Add($"Отсутствует обязательный раздел: '{section}'");
                }
            }
        }

        private bool ShouldSkipHeaderCheck(Paragraph paragraph)
        {
            // 1. Пропускаем пустые параграфы
            if (string.IsNullOrWhiteSpace(paragraph.InnerText))
                return true;

            // 2. Пропускаем оглавление
            if (IsTocParagraph(paragraph))
                return true;

            // 3. Пропускаем элементы списков
            if (IsListItem(paragraph))
                return true;

            // 4. Пропускаем таблицы
            if (paragraph.Ancestors<Table>().Any())
                return true;

            // 5. Пропускаем специальные стили
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(style) && style.Contains("TOC"))
                return true;

            return false;
        }

        private bool IsTocParagraph(Paragraph paragraph)
        {
            // 1. Проверка полей оглавления
            if (paragraph.Descendants<FieldCode>().Any(f =>
                f.Text.Contains(" TOC ") || f.Text.Contains("TOC \\")))
                return true;

            // 2. Проверка по характерным признакам
            string text = paragraph.InnerText;
            return text.Contains(".........") ||          // Точечные заполнители
                   text.Contains("\t") ||                 // Табуляция
                   Regex.IsMatch(text, @"\.{3,}\s*\d+$"); // Многоточие с номером страницы
        }

        private bool IsSectionMatch(string paragraphText, string section)
        {
            // Удаляем возможные номера (например "1 Введение" -> "Введение")
            string cleanText = Regex.Replace(paragraphText, @"^\d+\s*", "").Trim();
            return cleanText.Equals(section, StringComparison.OrdinalIgnoreCase);
        }

        private bool CheckHeaderFormatting(Paragraph paragraph, string section, Gost gost, List<string> errors)
        {
            bool hasErrors = false;
            var runs = paragraph.Descendants<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)).ToList();

            // Проверка шрифта
            if (!string.IsNullOrEmpty(gost.HeaderFontName))
            {
                foreach (var run in runs)
                {
                    var font = run.RunProperties?.RunFonts?.Ascii?.Value;
                    if (font != null && font != gost.HeaderFontName)
                    {
                        errors.Add($"Раздел '{section}': неверный шрифт ({font} вместо {gost.HeaderFontName})");
                        HighlightRun(run);
                        hasErrors = true;
                    }
                }
            }

            // Проверка размера шрифта
            if (gost.HeaderFontSize.HasValue)
            {
                foreach (var run in runs)
                {
                    var size = run.RunProperties?.FontSize?.Val?.Value;
                    if (size != null)
                    {
                        double actualSize = double.Parse(size) / 2;
                        if (Math.Abs(actualSize - gost.HeaderFontSize.Value) > 0.1)
                        {
                            errors.Add($"Раздел '{section}': неверный размер шрифта ({actualSize}pt вместо {gost.HeaderFontSize.Value}pt)");
                            HighlightRun(run);
                            hasErrors = true;
                        }
                    }
                }
            }

            // Проверка выравнивания
            if (!string.IsNullOrEmpty(gost.HeaderAlignment))
            {
                var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);
                if (currentAlignment != gost.HeaderAlignment)
                {
                    errors.Add($"Раздел '{section}': неверное выравнивание ({currentAlignment} вместо {gost.HeaderAlignment})");
                    hasErrors = true;
                }
            }

            return hasErrors;
        }

        private void UpdateBulletedListsUI(List<string> errors, bool listsValid, bool hasLists)
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (errors.Any())
                {
                    ErrorControlBulletedLists.Text = "Проблемы в списках:\n" + string.Join("\n", errors.Distinct());
                    ErrorControlBulletedLists.Foreground = Brushes.Red;
                }
                else if (!hasLists)
                {
                    ErrorControlBulletedLists.Text = "В документе отсутствуют обязательные списки";
                    ErrorControlBulletedLists.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlBulletedLists.Text = "Списки соответствуют ГОСТу";
                    ErrorControlBulletedLists.Foreground = Brushes.Green;
                }
            });
        }

        private void HighlightWholeParagraph(Paragraph paragraph)
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                if (run.RunProperties == null)
                {
                    run.RunProperties = new RunProperties();
                }
                run.RunProperties.Append(new Highlight() { Val = HighlightColorValues.Red });
            }
        }

        /// <summary>
        /// Проверка обязательных разделов (Введение, Заключение и т.д.)
        /// </summary>
        private bool CheckRequiredSections(Gost gost, Body body)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections))
                return true;

            var requiredSections = GetRequiredSectionsList(gost);
            bool allSectionsFound = true;
            bool allSectionsValid = true;
            var missingSections = new List<string>();
            var invalidSections = new List<string>();

            foreach (var section in requiredSections)
            {
                bool sectionFound = false;
                bool sectionValid = true;

                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    var text = paragraph.InnerText.Trim();
                    if (text.IndexOf(section, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        sectionFound = true;

                        // НОВАЯ ПРОВЕРКА: Шрифт заголовка
                        if (!string.IsNullOrEmpty(gost.HeaderFontName))
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                var font = run.RunProperties?.RunFonts?.Ascii?.Value;
                                if (font != null && font != gost.HeaderFontName)
                                {
                                    invalidSections.Add($"{section} (неверный шрифт: {font})");
                                    sectionValid = false;
                                    break;
                                }
                            }
                        }

                        // Проверка размера шрифта заголовка
                        if (gost.HeaderFontSize.HasValue &&
                            !CheckHeaderFontSize(paragraph, gost.HeaderFontSize.Value))
                        {
                            invalidSections.Add($"{section} (неверный размер шрифта)");
                            sectionValid = false;
                            break;
                        }

                        // Проверка выравнивания заголовка
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
        /// Проверяет, что основной текст заголовка имеет правильный размер шрифта
        /// (разрешает небольшие форматирования перед основным текстом)
        /// </summary>
        private bool CheckHeaderFontSize(Paragraph paragraph, double requiredSize)
        {
            if (paragraph == null) return false;

            var runs = paragraph.Elements<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)).ToList();

            if (!runs.Any()) return false;

            int totalLength = 0;
            int correctLength = 0;

            foreach (var run in runs)
            {
                var textLength = run.InnerText.Trim().Length;
                totalLength += textLength;

                var fontSize = run.RunProperties?.FontSize;
                if (fontSize != null)
                {
                    double fontSizeValue = double.Parse(fontSize.Val.Value) / 2;
                    if (Math.Abs(fontSizeValue - requiredSize) <= 0.1)
                    {
                        correctLength += textLength;
                    }
                }
            }

            return (double)correctLength / totalLength >= 0.8;
        }

        /// <summary>
        /// Проверка типа шрифта (только для обычных абзацев)
        /// </summary>
        private bool CheckFontName(string requiredFontName, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;
            var errors = new List<string>();

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                if (ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    if (ShouldSkipRun(run)) continue;

                    var fontName = run.RunProperties?.RunFonts?.Ascii?.Value;
                    if (fontName != null && fontName != requiredFontName)
                    {
                        errors.Add($"Найден неверный шрифт: '{fontName}' (ожидался '{requiredFontName}') в тексте: '{run.InnerText.Trim()}'");
                        isValid = false;
                    }
                }
            }

            if (!isValid)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlFont.Text = "Ошибки в шрифте основного текста:\n" + string.Join("\n", errors.Take(3));
                    if (errors.Count > 3) ErrorControlFont.Text += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlFont.Foreground = Brushes.Red;
                });
            }
            else
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlFont.Text = "Шрифт основного текста соответствует ГОСТу";
                    ErrorControlFont.Foreground = Brushes.Green;
                });
            }

            return isValid;
        }

        /// <summary>
        /// Проверка размера шрифта (только для обычных абзацев)
        /// </summary>
        private bool CheckFontSize(double requiredFontSize, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;
            var errors = new List<string>();

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                if (ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    if (ShouldSkipRun(run)) continue;

                    var fontSize = run.RunProperties?.FontSize;
                    if (fontSize != null)
                    {
                        double actualSize = double.Parse(fontSize.Val.Value) / 2;
                        if (Math.Abs(actualSize - requiredFontSize) > 0.1)
                        {
                            errors.Add($"Неверный размер: {actualSize}pt (ожидался {requiredFontSize}pt) в тексте: '{run.InnerText.Trim()}'");
                            isValid = false;
                        }
                    }
                }
            }

            if (!isValid)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlFontSize.Text = "Ошибки в размере шрифта:\n" + string.Join("\n", errors.Take(3));
                    if (errors.Count > 3) ErrorControlFontSize.Text += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlFontSize.Foreground = Brushes.Red;
                });
            }
            else
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlFontSize.Text = "Размер шрифта соответствует ГОСТу";
                    ErrorControlFontSize.Foreground = Brushes.Green;
                });
            }

            return isValid;
        }

        /// <summary>
        /// Определяет, нужно ли пропускать параграф при проверке основного текста
        /// </summary>
        private bool ShouldSkipParagraph(Paragraph paragraph, HashSet<string> headerTexts, Gost gost)
        {
            // Пропускаем заголовки
            if (headerTexts.Contains(paragraph.InnerText.Trim()))
                return true;

            // Пропускаем пустые параграфы
            if (IsEmptyParagraph(paragraph))
                return true;

            // Пропускаем элементы списков
            if (IsListItem(paragraph))
                return true;

            // Пропускаем таблицы
            if (paragraph.Ancestors<Table>().Any())
                return true;

            // Пропускаем специальные стили (если есть)
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(style) && style.Contains("Special"))
                return true;

            return false;
        }

        /// <summary>
        /// Определяет, нужно ли пропускать Run элемент при проверке
        /// </summary>
        private bool ShouldSkipRun(Run run)
        {
            // Пропускаем пустые Run элементы
            if (string.IsNullOrWhiteSpace(run.InnerText))
                return true;

            // Пропускаем специальные символы
            if (run.Elements<Break>().Any() || run.Elements<TabChar>().Any())
                return true;

            return false;
        }

        /// <summary>
        /// Проверка выравнивания текста (исключая заголовки, пустые параграфы и списки)
        /// </summary>
        private bool CheckTextAlignment(string requiredAlignment, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем: 1) заголовки, 2) пустые параграфы, 3) элементы списков
                if (headerTexts.Contains(paragraph.InnerText.Trim()) ||
                    IsEmptyParagraph(paragraph) ||
                    IsListItem(paragraph))
                    continue;

                // Проверяем только параграфы с текстом
                bool hasText = paragraph.Elements<Run>().Any(r => !string.IsNullOrWhiteSpace(r.InnerText));
                if (!hasText) continue;

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
        /// Формат листа "к примеру А4"
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckPaperSize(WordprocessingDocument doc, Gost gost)
        {
            var sectPr = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr == null) return false;

            var pgSz = sectPr.Elements<PageSize>().FirstOrDefault();
            if (pgSz == null) return false;

            // Конвертация в миллиметры
            double widthMm = (pgSz.Width.Value / 1440.0) * 25.4;
            double heightMm = (pgSz.Height.Value / 1440.0) * 25.4;

            // Допустимое отклонение 1 мм
            bool isCorrectSize = Math.Abs(widthMm - gost.PaperWidthMm.Value) <= 1 &&
                                 Math.Abs(heightMm - gost.PaperHeightMm.Value) <= 1;

            Dispatcher.UIThread.Post(() => {
                ErrorControlPaperSize.Text = isCorrectSize ? $"Формат бумаги: {gost.PaperSize} ({widthMm:F1}×{heightMm:F1} мм)" :
                                                                                          $"Требуется {gost.PaperSize} ({gost.PaperWidthMm}×{gost.PaperHeightMm} мм), " +
                                                                                          $"текущий: {widthMm:F1}×{heightMm:F1} мм";

                ErrorControlPaperSize.Foreground = isCorrectSize ? Brushes.Green : Brushes.Red;
            });

            return isCorrectSize;
        }

        /// <summary>
        /// Ориентация листа "Альбомная или Книжная"
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckPageOrientation(WordprocessingDocument doc, Gost gost)
        {
            var sectPr = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr == null) return false;

            var pgSz = sectPr.Elements<PageSize>().FirstOrDefault();
            if (pgSz == null) return false;

            bool isPortrait = pgSz.Orient == null || pgSz.Orient.Value == PageOrientationValues.Portrait;
            bool shouldBePortrait = gost.PageOrientation == "Portrait";

            Dispatcher.UIThread.Post(() => {
                ErrorControlOrientation.Text = (isPortrait == shouldBePortrait) ? $"Ориентация: {(shouldBePortrait ? "Книжная" : "Альбомная")} (соответствует)" :
                                                                                                               $"Ориентация: {(isPortrait ? "Книжная" : "Альбомная")} " +
                                                                                                               $"( Должна быть {(shouldBePortrait ? "Книжная" : "Альбомная")} )";

                ErrorControlOrientation.Foreground = (isPortrait == shouldBePortrait) ? Brushes.Green : Brushes.Red;
            });

            return isPortrait == shouldBePortrait;
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
                    string requiredAlignment = style.Type == StyleValues.Paragraph ? gost.TextAlignment : gost.HeaderAlignment;

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

            return gost.RequiredSections.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
        }

        /// <summary>
        /// Проверка нумерации страниц и её расположения
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="requiredNumbering"></param>
        /// <param name="requiredAlignment"></param>
        /// <param name="requiredPosition"></param>
        /// <returns></returns>
        private bool CheckPageNumbering(WordprocessingDocument wordDoc, bool requiredNumbering, string requiredAlignment, string requiredPosition)
        {
            if (!requiredNumbering)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlNumberPage.Text = "Нумерация страниц не требуется";
                    ErrorControlNumberPage.Foreground = Brushes.Gray;
                });
                return true;
            }

            bool hasCorrectNumbering = false;
            bool hasExtraNumbering = false;
            string actualCorrectPosition = "";
            string actualCorrectAlignment = "";
            List<string> extraNumberings = new List<string>();

            // Функция для выделения номера страницы и связанных элементов
            void HighlightPageNumbering(OpenXmlElement element)
            {
                // Находим все Run элементы, содержащие номер страницы
                var runs = new List<Run>();

                // 1. Сам элемент SimpleField (поле PAGE)
                if (element is SimpleField field)
                {
                    var parentRun = field.Parent as Run;
                    if (parentRun != null) runs.Add(parentRun);
                }

                // 2. Соседние Run элементы (могут содержать отображаемый номер)
                var sibling = element.NextSibling();
                while (sibling != null)
                {
                    if (sibling is Run run && !string.IsNullOrWhiteSpace(run.InnerText))
                    {
                        runs.Add(run);
                        break;
                    }
                    sibling = sibling.NextSibling();
                }

                // 3. Выделяем все найденные Run элементы
                foreach (var run in runs)
                {
                    run.RunProperties ??= new RunProperties();

                    // Удаляем старое выделение, если есть
                    var existingHighlight = run.RunProperties.Elements<Highlight>().FirstOrDefault();
                    if (existingHighlight != null)
                    {
                        run.RunProperties.RemoveChild(existingHighlight);
                    }

                    // Добавляем красное выделение
                    run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });

                    // Также добавляем красный цвет текста для лучшей видимости
                    var color = run.RunProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.Color>().FirstOrDefault();
                    if (color == null)
                    {
                        run.RunProperties.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "FF0000" });
                    }
                    else
                    {
                        color.Val = "FF0000";
                    }
                }
            }

            // Проверяем верхние колонтитулы (headers)
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

                            // Проверяем соответствие требованиям
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
                                HighlightPageNumbering(pageField);
                            }
                        }
                    }
                }
            }

            // Проверяем нижние колонтитулы (footers)
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

                            // Проверяем соответствие требованиям
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
                                HighlightPageNumbering(pageField);
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
        private bool CheckFirstLineIndent(double requiredFirstLineIndent, Body body, Gost gost)
        {
            var headerTexts = GetHeaderTexts(body, gost);
            bool isValid = true;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем: 1) заголовки, 2) пустые параграфы, 3) элементы списков
                if (headerTexts.Contains(paragraph.InnerText.Trim()) ||
                    IsEmptyParagraph(paragraph) ||
                    IsListItem(paragraph))
                    continue;

                // Проверяем только параграфы с текстом
                bool hasText = paragraph.Elements<Run>().Any(r => !string.IsNullOrWhiteSpace(r.InnerText));
                if (!hasText) continue;

                var indent = paragraph.ParagraphProperties?.Indentation;

                // Если отступ явно не задан, считаем его нулевым (что может быть ошибкой)
                if (indent?.FirstLine == null)
                {
                    // Для ГОСТа обычно требуется отступ, поэтому отсутствие отступа - ошибка
                    Dispatcher.UIThread.Post(() => {
                        ErrorControlFirstLineIndent.Text = "Ошибка: отсутствует отступ первой строки";
                        ErrorControlFirstLineIndent.Foreground = Brushes.Red;
                    });
                    isValid = false;
                    break;
                }

                // Преобразуем значение отступа в сантиметры (1 см = 567 twips)
                double firstLineIndentInCm = double.Parse(indent.FirstLine.Value) / 567.0;

                // Проверяем с допуском 0.1 см
                if (Math.Abs(firstLineIndentInCm - requiredFirstLineIndent) > 0.1)
                {
                    Dispatcher.UIThread.Post(() => {
                        ErrorControlFirstLineIndent.Text = $"Ошибка: отступ {firstLineIndentInCm:F2} см (требуется {requiredFirstLineIndent:F2} см)";
                        ErrorControlFirstLineIndent.Foreground = Brushes.Red;
                    });
                    isValid = false;
                    break;
                }
            }

            if (isValid)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlFirstLineIndent.Text = $"Отступ соответствует ГОСТу: {requiredFirstLineIndent:F2} см";
                    ErrorControlFirstLineIndent.Foreground = Brushes.Green;
                });
            }

            return isValid;
        }

        /// <summary>
        /// Проверяет, является ли параграф элементом списка
        /// </summary>
        private bool IsListItem(Paragraph paragraph)
        {
            // Проверяем явные признаки элемента списка
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // Проверяем стили, связанные со списками
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) &&
                (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering")))
                return true;

            // Проверяем наличие маркеров или номеров в тексте
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                var text = firstRun.InnerText.Trim();
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—") ||
                    Regex.IsMatch(text, @"^\d+[\.\)]"))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверка межстрочного интервала
        /// </summary>
        private bool CheckLineSpacing(double requiredLineSpacing, Body body, Gost gost)
        {
            bool isValid = true;
            var headerTexts = GetHeaderTexts(body, gost);

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // Пропускаем: 1) заголовки, 2) пустые параграфы, 3) элементы списков
                if (headerTexts.Contains(paragraph.InnerText.Trim()) ||
                    IsEmptyParagraph(paragraph) ||
                    IsListItem(paragraph))
                    continue;

                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                if (spacing?.Line == null || spacing.LineRule != LineSpacingRuleValues.Auto)
                    continue;

                double lineSpacing = double.Parse(spacing.Line.Value) / 240.0;
                if (Math.Abs(lineSpacing - requiredLineSpacing) > 0.01)
                {
                    isValid = false;
                    HighlightWholeParagraph(paragraph);
                    break;
                }
            }

            Dispatcher.UIThread.Post(() => {
                ErrorControlMnochitel.Text = isValid
                    ? $"Межстрочный интервал соответствует ГОСТу: {requiredLineSpacing}"
                    : "Ошибка в межстрочном интервале основного текста";
                ErrorControlMnochitel.Foreground = isValid ? Brushes.Green : Brushes.Red;
            });

            return isValid;
        }

        /// <summary>
        /// Проверка полей документа
        /// </summary>
        private bool CheckMargins(double? requiredMarginTop, double? requiredMarginBottom, double? requiredMarginLeft, double? requiredMarginRight, Body body)
        {
            var pageMargin = body.Elements<SectionProperties>().FirstOrDefault()?.Elements<PageMargin>().FirstOrDefault();

            if (pageMargin == null) return false;

            // Преобразование в сантиметры (1 см = 567 twips)
            double marginTopInCm = pageMargin.Top.Value / 567.0;
            double marginBottomInCm = pageMargin.Bottom.Value / 567.0;
            double marginLeftInCm = pageMargin.Left.Value / 567.0;
            double marginRightInCm = pageMargin.Right.Value / 567.0;

            // Проверка с погрешностью 0.01 см
            if (requiredMarginTop.HasValue && Math.Abs(marginTopInCm - requiredMarginTop.Value) > 0.01)
                return false;

            if (requiredMarginBottom.HasValue && Math.Abs(marginBottomInCm - requiredMarginBottom.Value) > 0.01)
                return false;

            if (requiredMarginLeft.HasValue && Math.Abs(marginLeftInCm - requiredMarginLeft.Value) > 0.01)
                return false;

            if (requiredMarginRight.HasValue && Math.Abs(marginRightInCm - requiredMarginRight.Value) > 0.01)
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

        /// <summary>
        /// Показывает диалоговое окно подтверждения
        /// </summary>
        private async Task<bool> ShowConfirmationDialog(Window parent, string title, string message)
        {
            var result = false;

            // Создаем основное окно
            var dialog = new Window
            {
                Title = title,
                Width = 550,
                Height = 200,
                MinWidth = 400,
                MinHeight = 180,
                MaxWidth = 550,
                MaxHeight = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                SizeToContent = SizeToContent.Manual,
                WindowState = WindowState.Normal,
                CanResize = false,
                Topmost = true
            };

            // Создаем кнопки
            var yesButton = new Button
            {
                Content = "Да",
                Width = 80,
                Margin = new Avalonia.Thickness(0, 0, 10, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            var noButton = new Button
            {
                Content = "Нет",
                Width = 80,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            // Настраиваем содержимое окна
            dialog.Content = new Avalonia.Controls.Border
            {
                BorderBrush = Brushes.Red,
                BorderThickness = new Avalonia.Thickness(1),
                Padding = new Avalonia.Thickness(20),
                Child = new StackPanel
                {
                    Orientation = Avalonia.Layout.Orientation.Vertical,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Spacing = 20,
                    Children =
            {
                new TextBlock
                {
                    Text = message,
                    TextWrapping = TextWrapping.Wrap,
                    TextAlignment = Avalonia.Media.TextAlignment.Center,
                    FontSize = 14
                },
                new StackPanel
                {
                    Orientation = Avalonia.Layout.Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Spacing = 15,
                    Children = { yesButton, noButton }
                }
            }
                }
            };

            // Назначаем обработчики кнопок
            yesButton.Click += (s, e) => { result = true; dialog.Close(); };
            noButton.Click += (s, e) => { result = false; dialog.Close(); };

            await dialog.ShowDialog(this);
            return result;
        }

    }
}