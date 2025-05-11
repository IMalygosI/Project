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
using DocumentFormat.OpenXml.ExtendedProperties;



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

        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ПРОСТОГО ТЕКСТА =======================
        private const string DefaultTextFont = "Arial";
        private const double DefaultTextSize = 11.0;
        private const string DefaultTextAlignment = "Left";
        private const string DefaultTextLineSpacingType = "Множитель";
        private const double DefaultTextLineSpacingValue = 1.15;
        private const double DefaultTextSpacingBefore = 0.0;
        private const double DefaultTextSpacingAfter = 0.35;
        private const string DefaultTextFirstLineType = "Нет";
        private const double DefaultTextFirstLineIndent = 1.25;
        private const double DefaultTextLeftIndent = 0.0;
        private const double DefaultTextRightIndent = 0.0;

        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ЗАГОЛОВКОВ =======================
        private const string DefaultHeaderFont = "Arial";
        private const double DefaultHeaderSize = 20.0;
        private const string DefaultHeaderAlignment = "Left";
        private const string DefaultHeaderLineSpacingType = "Множитель";
        private const double DefaultHeaderLineSpacingValue = 1.15;
        private const double DefaultHeaderSpacingBefore = 0.85;
        private const double DefaultHeaderSpacingAfter = 0.35;
        private const string DefaultHeaderFirstLineType = "Нет";
        private const double DefaultHeaderFirstLineIndent = 0.0;
        private const double DefaultHeaderLeftIndent = 0.0;
        private const double DefaultHeaderRightIndent = 0.0;

        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ СПИСКОВ =======================
        private const string DefaultListFont = "Arial";
        private const double DefaultListSize = 11.0;
        private const string DefaultListLineSpacingType = "Множитель";
        private const double DefaultListLineSpacingValue = 1.15;
        private const double DefaultListSpacingBefore = 0.0;
        private const double DefaultListSpacingAfter = 0.35;
        private const string DefaultListFirstLineType = "Выступ";
        private const double DefaultListHangingIndent = 0.64;
        private const double DefaultListLeftIndent = 0.62;
        private const double DefaultListRightIndent = 0.0;
        // для многоуровневых
        private const double DefaultListLevel1BulletIndentLeft = 1.87;
        private const double DefaultListLevel2BulletIndentLeft = 2.5;
        private const double DefaultListLevel3BulletIndentLeft = 3.14;
        private const double DefaultListLevel4BulletIndentLeft = 3.77;
        private const double DefaultListLevel5BulletIndentLeft = 4.41;
        private const double DefaultListLevel6BulletIndentLeft = 5.04;
        private const double DefaultListLevel7BulletIndentLeft = 5.68;
        private const double DefaultListLevel8BulletIndentLeft = 6.31;
        private const double DefaultListLevel9BulletIndentLeft = 6.95;
        private const double DefaultListLevel1BulletIndentRight = 0;
        private const double DefaultListLevel2BulletIndentRight = 0;
        private const double DefaultListLevel3BulletIndentRight = 0;
        private const double DefaultListLevel4BulletIndentRight = 0;
        private const double DefaultListLevel5BulletIndentRight = 0;
        private const double DefaultListLevel6BulletIndentRight = 0;
        private const double DefaultListLevel7BulletIndentRight = 0;
        private const double DefaultListLevel8BulletIndentRight = 0;
        private const double DefaultListLevel9BulletIndentRight = 0;
        private const double DefaultListLevel1Indent = 0.64;
        private const double DefaultListLevel2Indent = 0.76;
        private const double DefaultListLevel3Indent = 0.89;
        private const double DefaultListLevel4Indent = 1.14;
        private const double DefaultListLevel5Indent = 1.4;
        private const double DefaultListLevel6Indent = 1.65;
        private const double DefaultListLevel7Indent = 1.91;
        private const double DefaultListLevel8Indent = 2.16;
        private const double DefaultListLevel9Indent = 2.54;
        private const string DefaultListLevel1NumberFormat = "1.";
        private const string DefaultListLevel2NumberFormat = "1.1";
        private const string DefaultListLevel3NumberFormat = "1.1.1";
        private const string DefaultListLevel4NumberFormat = "1.1.1.1";
        private const string DefaultListLevel5NumberFormat = "1.1.1.1.1";
        private const string DefaultListLevel6NumberFormat = "1.1.1.1.1.1";
        private const string DefaultListLevel7NumberFormat = "1.1.1.1.1.1.1";
        private const string DefaultListLevel8NumberFormat = "1.1.1.1.1.1.1.1";
        private const string DefaultListLevel9NumberFormat = "1.1.1.1.1.1.1.1.1";
        private const string DefaultListLevel1IndentOrOutdent = "Выступ";
        private const string DefaultListLevel2IndentOrOutdent = "Выступ";
        private const string DefaultListLevel3IndentOrOutdent = "Выступ";
        private const string DefaultListLevel4IndentOrOutdent = "Выступ";
        private const string DefaultListLevel5IndentOrOutdent = "Выступ";
        private const string DefaultListLevel6IndentOrOutdent = "Выступ";
        private const string DefaultListLevel7IndentOrOutdent = "Выступ";
        private const string DefaultListLevel8IndentOrOutdent = "Выступ";
        private const string DefaultListLevel9IndentOrOutdent = "Выступ";

        /// <summary>
        /// Вспомогательный метод для получения требуемого типа отступа в списках
        /// </summary>
        /// <param name="gost"></param>
        /// <param name="level"></param>
        /// <returns></returns>
        private string GetRequiredIndentType(Gost gost, int level)
        {
            return level switch
            {
                1 => gost.ListLevel1IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                2 => gost.ListLevel2IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                3 => gost.ListLevel3IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                4 => gost.ListLevel4IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                5 => gost.ListLevel5IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                6 => gost.ListLevel6IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                7 => gost.ListLevel7IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                8 => gost.ListLevel8IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                9 => gost.ListLevel9IndentOrOutdent ?? gost.BulletIndentOrOutdent,
                _ => gost.BulletIndentOrOutdent
            };
        }

        /// <summary>
        /// Вспомогательный метод для получения требуемого левого отступа в списках
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        private double GetListLevelIndentLeft(int level)
        {
            switch (level)
            {
                case 1: return DefaultListLevel1BulletIndentLeft;
                case 2: return DefaultListLevel2BulletIndentLeft;
                case 3: return DefaultListLevel3BulletIndentLeft;
                case 4: return DefaultListLevel4BulletIndentLeft;
                case 5: return DefaultListLevel5BulletIndentLeft;
                case 6: return DefaultListLevel6BulletIndentLeft;
                case 7: return DefaultListLevel7BulletIndentLeft;
                case 8: return DefaultListLevel8BulletIndentLeft;
                case 9: return DefaultListLevel9BulletIndentLeft;
                default: return DefaultListLevel1BulletIndentLeft; // по умолчанию уровень 1
            }
        }

        /// <summary>
        /// Вспомогательный метод для получения требуемого правого отступа в списках
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        private double GetListLevelIndentRight(int level)
        {
            switch (level)
            {
                case 1: return DefaultListLevel1BulletIndentRight;
                case 2: return DefaultListLevel2BulletIndentRight;
                case 3: return DefaultListLevel3BulletIndentRight;
                case 4: return DefaultListLevel4BulletIndentRight;
                case 5: return DefaultListLevel5BulletIndentRight;
                case 6: return DefaultListLevel6BulletIndentRight;
                case 7: return DefaultListLevel7BulletIndentRight;
                case 8: return DefaultListLevel8BulletIndentRight;
                case 9: return DefaultListLevel9BulletIndentRight;
                default: return DefaultListLevel1BulletIndentRight; // по умолчанию уровень 1
            }
        }

        /// <summary>
        /// Вспомогательный метод для получения требуемого отступа первой строки в списках
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        private double GetListLevelIndent(int level)
        {
            switch (level)
            {
                case 1: return DefaultListLevel1Indent;
                case 2: return DefaultListLevel2Indent;
                case 3: return DefaultListLevel3Indent;
                case 4: return DefaultListLevel4Indent;
                case 5: return DefaultListLevel5Indent;
                case 6: return DefaultListLevel6Indent;
                case 7: return DefaultListLevel7Indent;
                case 8: return DefaultListLevel8Indent;
                case 9: return DefaultListLevel9Indent;
                default: return DefaultListLevel1Indent; // по умолчанию уровень 1
            }
        }

        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ОГЛАВЛЕНИЯ =======================
        private string DefaultTocFont = "Arial";
        private const double DefaultTocSize = 11.0;
        private const string DefaultTocAlignment = "Left";
        private const string DefaultTocLineSpacingType = "Множитель";
        private const double DefaultTocLineSpacingValue = 1.15;
        private const double DefaultTocSpacingBefore = 0.0;
        private const double DefaultTocSpacingAfter = 0.1;
        private const string DefaultTocFirstLineType = "Нет";
        private const double DefaultTocFirstLineIndent = 0.0;
        private const double DefaultTocLeftIndent = 0.0;
        private const double DefaultTocRightIndent = 0.0;

        /// <summary>
        /// Конструктор по умолчанию класса GOST_Сheck. Инициализирует компоненты окна.
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

        private bool IsTitleParagraph(Paragraph paragraph)
        {
            if (paragraph?.InnerText == null) return false;

            string text = paragraph.InnerText.Trim().ToUpper();
            if (string.IsNullOrEmpty(text)) return false;

            // Ключевые слова для титульного листа
            string[] titleKeywords = {
        "МИНИСТЕРСТВО", "УНИВЕРСИТЕТ", "ИНСТИТУТ", "ФАКУЛЬТЕТ", "КАФЕДРА",
        "ДИСЦИПЛИНА", "КУРСОВАЯ", "ДИПЛОМНАЯ", "РАБОТА", "ПРОЕКТ",
        "РЕФЕРАТ", "ОТЧЕТ", "ВЫПОЛНИЛ", "ПРОВЕРИЛ", "Г.", "ГОД"
    };

            // Проверка, если в тексте содержатся ключевые слова
            return titleKeywords.Any(k => text.Contains(k));
        }

        private void ExtractTitleAndBody(Body body, out List<Paragraph> titlePageParagraphs, out List<Paragraph> bodyParagraphsAfterTitle, out List<Paragraph> allParagraphs)
        {
            titlePageParagraphs = new List<Paragraph>();
            allParagraphs = body.Elements<Paragraph>().ToList();
            int startIndex = 0;

            for (int i = 0; i < allParagraphs.Count; i++)
            {
                var paragraph = allParagraphs[i];
                if (IsTitleParagraph(paragraph))
                {
                    titlePageParagraphs.Add(paragraph);
                }
                else if (paragraph.Descendants<Break>().Any(b => b.Type == BreakValues.Page))
                {
                    startIndex = i + 1;
                    break;
                }
            }

            bodyParagraphsAfterTitle = allParagraphs.Skip(startIndex).ToList();
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
                        bool textIndentsValid = true;
                        bool paragraphSpacingValid = true;
                        bool headerSpacingValid = true;
                        bool tocSpacingValid = true;
                        bool listSpacingValid = true;
                        bool listHangingValid = true;
                        bool headerIndentsValid = true;
                        bool tocIndentsValid = true;
                        bool plainTextLinksValid = true;

                        // Список хранения ошибок
                        var errors = new List<string>();

                        // ======================= ТИТУЛЬНИК =======================
                        List<Paragraph> allParagraphs;
                        ExtractTitleAndBody(body, out var titlePageParagraphs, out var bodyParagraphsAfterTitle, out allParagraphs);

                        //ExtractTitleAndBody(body, out var titlePageParagraphs, out var bodyParagraphsAfterTitle);

                        // ======================= ПРОСТОЙ ТЕКСТ =======================
                        if (!string.IsNullOrEmpty(gost.FontName))
                        {
                            fontNameValid = CheckFontName(gost.FontName, bodyParagraphsAfterTitle, gost, wordDoc);// Проверка типа шрифта
                            ErrorControlFont.Text = fontNameValid ? "Тип шрифта соответствует ГОСТу." : "Тип шрифта не соответствует.";
                            ErrorControlFont.Foreground = fontNameValid ? Brushes.Green : Brushes.Red;
                            if (!fontNameValid) errors.Add("Тип шрифта не соответствует ГОСТу");
                        }

                        if (gost.FontSize.HasValue)
                        {
                            fontSizeValid = CheckFontSize(gost.FontSize.Value, bodyParagraphsAfterTitle, gost, wordDoc);// Проверка размера шрифта 
                            ErrorControlFontSize.Text = fontSizeValid ? "Размер шрифта соответствует ГОСТу!" : "Размер шрифта не соответствует!";
                            ErrorControlFontSize.Foreground = fontSizeValid ? Brushes.Green : Brushes.Red;
                            if (!fontSizeValid) errors.Add("Размер шрифта не соответствует ГОСТу");
                        }

                        if (!string.IsNullOrEmpty(gost.TextAlignment))
                        {
                            textAlignmentValid = CheckTextAlignment(gost.TextAlignment, bodyParagraphsAfterTitle, gost);// Проверка выравнивания текст
                            ErrorControlViravnivanie.Text = textAlignmentValid ? "Выравнивание текста соответствует ГОСТу." : "Выравнивание текста не соответствует ГОСТу.";
                            ErrorControlViravnivanie.Foreground = textAlignmentValid ? Brushes.Green : Brushes.Red;
                            if (!textAlignmentValid) errors.Add("Выравнивание текста не соответствует ГОСТу");
                        }

                        if (gost.LineSpacingValue.HasValue)
                        {
                            lineSpacingValid = CheckLineSpacing(gost.LineSpacingValue ?? 1.5, bodyParagraphsAfterTitle, gost);// Проверка межстрочного интервала
                            ErrorControlMnochitel.Text = lineSpacingValid ? "Межстрочный интервал соответствует ГОСТу." : "Межстрочный интервал не соответствует ГОСТу.";
                            ErrorControlMnochitel.Foreground = lineSpacingValid ? Brushes.Green : Brushes.Red;
                            if (!lineSpacingValid) errors.Add("Межстрочный интервал не соответствует ГОСТу");
                        }

                        paragraphSpacingValid = CheckParagraphSpacing(bodyParagraphsAfterTitle, gost, wordDoc);// Проверка интервалов между абзацами для простого текста
                        if (gost.LineSpacingBefore.HasValue || gost.LineSpacingAfter.HasValue)
                        {
                            if (!paragraphSpacingValid)
                            {
                                errors.Add("Интервалы между абзацами не соответствуют ГОСТу");
                            }
                        }

                        if (gost.FirstLineIndent.HasValue)
                        {
                            firstLineIndentValid = CheckFirstLineIndent(gost.FirstLineIndent.Value, bodyParagraphsAfterTitle, gost);// Проверка отступов у простого текста
                            ErrorControlFirstLineIndent.Text = firstLineIndentValid ? "Отступ соответствует ГОСТу." : "Отступ не соответствует ГОСТу.";
                            ErrorControlFirstLineIndent.Foreground = firstLineIndentValid ? Brushes.Green : Brushes.Red;
                            if (!firstLineIndentValid) errors.Add("Отступ первой строки не соответствует ГОСТу");
                        }

                        // ======================= ПОЛЯ ДОКУМЕНТА =======================
                        if (gost.MarginTop.HasValue || gost.MarginBottom.HasValue || gost.MarginLeft.HasValue || gost.MarginRight.HasValue)
                        {
                            marginsValid = CheckMargins(gost.MarginTop, gost.MarginBottom, gost.MarginLeft, gost.MarginRight, body);// Проверка полей документа
                            ErrorControlMargins.Text = marginsValid ? "Поля документа соответствуют ГОСТу." : "Поля документа не соответствуют ГОСТу.";
                            ErrorControlMargins.Foreground = marginsValid ? Brushes.Green : Brushes.Red;
                            if (!marginsValid) errors.Add("Поля документа не соответствуют ГОСТу");
                        }

                        // ======================= НУМЕРАЦИЯ =======================
                        if (gost.PageNumbering.HasValue)// Проверка нумерации страниц
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

                        // ======================= ЗАГОЛОВКИ =======================
                        if (!string.IsNullOrEmpty(gost.RequiredSections))
                        {
                            sectionsValid = CheckRequiredSections(gost, bodyParagraphsAfterTitle, wordDoc);// Проверка обязательных разделов (Введение, Заключение) ЗАГОЛОВКИ
                            if (!sectionsValid) errors.Add("Отсутствуют обязательные разделы");
                        }

                        if (gost.HeaderLineSpacingBefore.HasValue || gost.HeaderLineSpacingAfter.HasValue)
                        {
                            headerSpacingValid = CheckHeaderParagraphSpacing(wordDoc, bodyParagraphsAfterTitle, gost);// Проверка интервалов для заголовков
                            if (!headerSpacingValid) errors.Add("Интервалы заголовков не соответствуют ГОСТу");
                        }

                        if (gost.HeaderIndentLeft.HasValue || gost.HeaderIndentRight.HasValue || gost.HeaderFirstLineIndent.HasValue)
                        {
                            headerIndentsValid = CheckHeaderIndents(bodyParagraphsAfterTitle, gost);// Проверка отступов для заголовков
                            if (!headerIndentsValid) errors.Add("Отступы заголовков не соответствуют ГОСТу");
                        }

                        // ======================= ФОРМАТ ЛИСТА =======================
                        if (gost.PaperWidthMm.HasValue && gost.PaperHeightMm.HasValue)
                        {
                            paperSizeValid = CheckPaperSize(wordDoc, gost);// Проверка формата 
                            if (!paperSizeValid) errors.Add("Размер бумаги не соответствует ГОСТу");
                        }

                        // ======================= ОРИЕНТАЦИИ ЛИСТА =======================
                        if (!string.IsNullOrEmpty(gost.PageOrientation))
                        {
                            orientationValid = CheckPageOrientation(wordDoc, gost);// Проверка Ориентации документа
                            if (!orientationValid) errors.Add("Ориентация страницы не соответствует ГОСТу");
                        }

                        // ======================= ОГЛАВЛЕНИЯ =======================
                        if (gost.RequireTOC.HasValue && gost.RequireTOC.Value)
                        {
                            tocValid = CheckTableOfContents(wordDoc, gost);// Проверка оглавления
                            if (!tocValid) errors.Add("Оглавление не соответствует ГОСТу");
                        }

                        if (gost.TocIndentLeft.HasValue || gost.TocIndentRight.HasValue || gost.TocFirstLineIndent.HasValue)
                        {
                            tocIndentsValid = CheckTocIndents(wordDoc, gost);// Проверка отступов в оглавлении
                            if (!tocIndentsValid) errors.Add("Отступы Оглавления не соответствуют ГОСТу");
                        }

                        // ======================= СПИСКИ =======================

                        bool hasLists = body.Descendants<Paragraph>().Any(IsListItem);// Проверка списков (если они есть в документе)
                        if (hasLists)
                        {
                            bulletedListsValid = CheckBulletedLists(bodyParagraphsAfterTitle, gost);// Проверка базовых параметров списков
                            if (!bulletedListsValid) errors.Add("Списки не соответствуют ГОСТу");

                            if (gost.BulletLineSpacingBefore.HasValue || gost.BulletLineSpacingAfter.HasValue || gost.BulletLineSpacingValue.HasValue)
                            {
                                listSpacingValid = CheckListParagraphSpacing(bodyParagraphsAfterTitle, gost);// Проверка интервалов списков
                                if (!listSpacingValid) errors.Add("Интервалы списков не соответствуют ГОСТу");
                            }

                            if (gost.BulletIndentLeft.HasValue || gost.BulletIndentRight.HasValue || gost.ListHangingIndent.HasValue ||
                                gost.ListLevel1Indent.HasValue || gost.ListLevel2Indent.HasValue || gost.ListLevel3Indent.HasValue)
                            {
                                listHangingValid = CheckListIndents(bodyParagraphsAfterTitle, gost);// Проверка отступов списков
                                if (!listHangingValid) errors.Add("Отступы списков не соответствуют ГОСТу");
                            }
                        }
                        else
                        {
                            // Если списков нет, просто отмечаем что проверка не требуется
                            Dispatcher.UIThread.Post(() => {
                                ErrorControlBulletedLists.Text = "Списки не обнаружены - проверка не требуется";
                                ErrorControlBulletedLists.Foreground = Brushes.Gray;
                            });
                        }

                        // ======================= ПРОСТОЙ ТЕКСТ - ПРОВЕРКА НЕОФОРМЛЕННЫХ ГИПЕРССЫЛОК =======================
                        plainTextLinksValid = CheckPlainTextLinks(wordDoc); // Проверка гиперссылок
                        if (!plainTextLinksValid)
                        {
                            errors.Add("Гиперссылки оформлены не корректно!");
                        }

                        bool stylesValid = CheckStyleFonts(wordDoc, gost);
                        if (!stylesValid)
                        {
                            ErrorControlFont.Text = "Ошибка в стилях документа!";
                            ErrorControlFont.Foreground = Brushes.Red;
                        }

                        // Общий результат проверки
                        if (fontNameValid && fontSizeValid && marginsValid && lineSpacingValid && firstLineIndentValid && textAlignmentValid && pageNumberingValid && sectionsValid &&
                            paperSizeValid && orientationValid && tocValid && bulletedListsValid && textIndentsValid && paragraphSpacingValid && headerSpacingValid && tocSpacingValid &&
                            listSpacingValid && listHangingValid && headerIndentsValid && tocIndentsValid && plainTextLinksValid && stylesValid)
                        {
                            GostControl.Text = "Документ соответствует ГОСТу.";
                            GostControl.Foreground = Brushes.Green;
                        }
                        else
                        {
                            GostControl.Text = "Документ не соответствует ГОСТу:";
                            GostControl.Foreground = Brushes.Red;

                            // Создаем документ с ошибками
                            await CreateErrorReportDocument(wordDoc, gost, errors, filePath, titlePageParagraphs, bodyParagraphsAfterTitle);
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
        /// Метод что вызывает другие методы выделения ошибок в полях и создает документ с помечеными красным цветом ошибками
        /// </summary>
        /// <param name="originalDoc"></param>
        /// <param name="gost"></param>
        /// <param name="errors"></param>
        /// <param name="originalFilePath"></param>
        /// <param name="titlePageParagraphs"></param>
        /// <param name="bodyParagraphsAfterTitle"></param>
        /// <returns></returns>
        private async Task CreateErrorReportDocument(WordprocessingDocument originalDoc, Gost gost, List<string> errors, string originalFilePath, List<Paragraph> oldTitlePageParagraphs, List<Paragraph> oldBodyParagraphsAfterTitle)
        {
            // 1. Диалог подтверждения
            var mainWindow = (Avalonia.Application.Current.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;

            bool confirmSave = await ShowConfirmationDialog(mainWindow, "Сохранить отчет об ошибках", "Документ не соответствует ГОСТу. Хотите сохранить отчет с выделенными ошибками?");
            if (!confirmSave) return;

            // 2. Создание временного файла
            string tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.docx");
            File.Copy(originalFilePath, tempPath, true);

            // 3. Диалог сохранения
            var saveDialog = new SaveFileDialog
            {
                Title = "Сохранить отчет об ошибках",
                InitialFileName = $"Ошибки_{Path.GetFileNameWithoutExtension(originalFilePath)}",
                DefaultExtension = ".docx",
                Filters = new List<FileDialogFilter> { new() { Name = "Word Documents", Extensions = { "docx" } } }
            };

            var saveResult = await saveDialog.ShowAsync(mainWindow);
            if (string.IsNullOrEmpty(saveResult)) return;

            // 4. Работа с копией документа
            using (var errorDoc = WordprocessingDocument.Open(tempPath, true))
            {
                var body = errorDoc.MainDocumentPart.Document.Body;

                // Повторно извлекаем абзацы уже из нового документа
                ExtractTitleAndBody(body, out var titlePageParagraphs, out var bodyParagraphsAfterTitle, out var allParagraphs);

                // Подсветка ошибок текста
                HighlightTextFormattingErrors(bodyParagraphsAfterTitle, errorDoc, gost, errors);

                // заголовки
                HighlightHeaderErrors(errorDoc, gost, errors);

                // Оглавление
                HighlightTocErrors(errorDoc, gost);

                // списки
                HighlightListErrors(errorDoc, gost, errors);

                if (gost.PageNumbering.HasValue)
                {
                    CheckPageNumbering(errorDoc, gost.PageNumbering.Value, gost.PageNumberingAlignment, gost.PageNumberingPosition);
                }

                // гиперссылки
                HighlightPlainTextLinks(errorDoc);

                // Сохраняем изменения
                errorDoc.MainDocumentPart.Document.Save();
            }

            // 5. Копирование результата и запуск
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
        /// Проверяет на наличие гиперссылок
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>Возвращает true, если гиперссылки оформлены корректно, и false в случае ошибок</returns>
        private bool CheckPlainTextLinks(WordprocessingDocument doc)
        {
            var linkErrors = new List<string>(); // Список для ошибок
            var regex = new Regex(@"https?://[^\s]+", RegexOptions.Compiled);
            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();
            bool isValid = true;
            bool instructionShown = false;

            // Перебираем все параграфы документа
            foreach (var paragraph in paragraphs)
            {
                var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text)); // Сбор текста параграфа
                var matches = regex.Matches(text); // Поиск всех URL-адресов

                foreach (Match match in matches)
                {
                    bool isLinked = paragraph.Descendants<Hyperlink>().Any(h => h.InnerText.Contains(match.Value)); // Проверяем, есть ли гиперссылка

                    if (!isLinked)
                    {
                        linkErrors.Add($"Параграф: '{match.Value}' не оформлен как гиперссылка");
                        isValid = false;
                    }
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (!isValid)
                {
                    string errorMessage = "Ошибки в гиперссылках:\n" + string.Join("\n", linkErrors.Take(5)); 
                    if (linkErrors.Count > 5)
                        errorMessage += $"\n...и ещё {linkErrors.Count - 5} ошибок";
                    ErrorControlLinks.Text = errorMessage;
                    ErrorControlLinks.Foreground = Brushes.Red; 

                    // Проверка на то была ли уже выведена инструкция по установке гиперссылок
                    if (!instructionShown)
                    {
                        ErrorControlLinks.Text += "\nДля того чтобы оформить гиперссылку, выделите текст и нажмите Ctrl+K, затем вставьте нужный URL в поле ссылки.";
                        instructionShown = true; 
                    }
                }
                else
                {
                    ErrorControlLinks.Text = "Все гиперссылки оформлены корректно."; 
                    ErrorControlLinks.Foreground = Brushes.Green; 
                }
            });

            return isValid;
        }

        /// <summary>
        /// Окрашивает гиперссылки красным если допущена ошибка
        /// </summary>
        /// <param name="doc"></param>
        private void HighlightPlainTextLinks(WordprocessingDocument doc)
        {
            var regex = new Regex(@"https?://[^\s]+", RegexOptions.Compiled);
            var body = doc.MainDocumentPart.Document.Body;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
                var matches = regex.Matches(text);

                foreach (Match match in matches)
                {
                    bool isLinked = paragraph.Descendants<Hyperlink>().Any(h => h.InnerText.Contains(match.Value));
                    if (!isLinked)
                    {
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            foreach (var txt in run.Elements<Text>())
                            {
                                if (txt.Text.Contains(match.Value))
                                {
                                    run.RunProperties ??= new RunProperties();
                                    run.RunProperties.Highlight = new Highlight { Val = HighlightColorValues.Red };
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Проверка интервалов для заголовков
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckHeaderParagraphSpacing(WordprocessingDocument doc, List<Paragraph> paragraphs, Gost gost)
        {
            if (!gost.HeaderLineSpacingValue.HasValue && !gost.HeaderLineSpacingBefore.HasValue && !gost.HeaderLineSpacingAfter.HasValue)
                return true;

            bool hasErrors = false;
            var headerTexts = GetHeaderTexts(paragraphs, gost);
            var errors = new List<string>();
            var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
            var styles = stylesPart?.Styles;

            foreach (var paragraph in paragraphs)
            {
                if (!headerTexts.Contains(paragraph.InnerText.Trim()))
                    continue;

                bool paragraphHasErrors = false;
                var paraErrors = new List<string>();
                var explicitSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                Style style = null;
                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

                if (styleId != null && styles != null)
                {
                    style = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                }

                var styleSpacing = style?.StyleParagraphProperties?.SpacingBetweenLines;

                // Проверка межстрочного интервала
                if (gost.HeaderLineSpacingValue.HasValue)
                {
                    double actualSpacing = DefaultHeaderLineSpacingValue;
                    string actualSpacingType = DefaultHeaderLineSpacingType;
                    LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                    // фактическое значение
                    var spacingSource = explicitSpacing ?? styleSpacing;
                    if (spacingSource?.Line != null)
                    {
                        if (spacingSource.LineRule?.Value == LineSpacingRuleValues.Exact)
                        {
                            actualSpacing = double.Parse(spacingSource.Line.Value) / 20.0;
                            actualSpacingType = "Точно";
                            actualRule = LineSpacingRuleValues.Exact;
                        }
                        else if (spacingSource.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                        {
                            actualSpacing = double.Parse(spacingSource.Line.Value) / 20.0;
                            actualSpacingType = "Минимум";
                            actualRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            actualSpacing = double.Parse(spacingSource.Line.Value) / 240.0;
                            actualSpacingType = "Множитель";
                            actualRule = LineSpacingRuleValues.Auto;
                        }
                    }

                    LineSpacingRuleValues requiredRule = (gost.HeaderLineSpacingType ?? DefaultHeaderLineSpacingType) switch
                    {
                        "Минимум" => LineSpacingRuleValues.AtLeast,
                        "Точно" => LineSpacingRuleValues.Exact,
                        _ => LineSpacingRuleValues.Auto
                    };

                    string requiredType = gost.HeaderLineSpacingType ?? DefaultHeaderLineSpacingType;

                    // Проверка типа интервала
                    if (actualRule != requiredRule)
                    {
                        paraErrors.Add($"тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                        paragraphHasErrors = true;
                    }

                    // Проверка значения интервала
                    if (Math.Abs(actualSpacing - (gost.HeaderLineSpacingValue ?? DefaultHeaderLineSpacingValue)) > 0.1)
                    {
                        paraErrors.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {(gost.HeaderLineSpacingValue ?? DefaultHeaderLineSpacingValue):F2})");
                        paragraphHasErrors = true;
                    }
                }

                // Проверка интервалов "Перед" и "После"
                if (gost.HeaderLineSpacingBefore.HasValue || gost.HeaderLineSpacingAfter.HasValue)
                {
                    double actualBefore = DefaultHeaderSpacingBefore;
                    double actualAfter = DefaultHeaderSpacingAfter;

                    if (explicitSpacing?.Before?.Value != null)
                    {
                        actualBefore = ConvertTwipsToPoints(explicitSpacing.Before.Value);
                    }
                    else if (styleSpacing?.Before?.Value != null)
                    {
                        actualBefore = ConvertTwipsToPoints(styleSpacing.Before.Value);
                    }

                    if (gost.HeaderLineSpacingBefore.HasValue &&
                        Math.Abs(actualBefore - gost.HeaderLineSpacingBefore.Value) > 0.1)
                    {
                        paraErrors.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.HeaderLineSpacingBefore.Value:F1} pt)");
                        paragraphHasErrors = true;
                    }

                    if (explicitSpacing?.After?.Value != null)
                    {
                        actualAfter = ConvertTwipsToPoints(explicitSpacing.After.Value);
                    }
                    else if (styleSpacing?.After?.Value != null)
                    {
                        actualAfter = ConvertTwipsToPoints(styleSpacing.After.Value);
                    }

                    if (gost.HeaderLineSpacingAfter.HasValue &&
                        Math.Abs(actualAfter - gost.HeaderLineSpacingAfter.Value) > 0.1)
                    {
                        paraErrors.Add($"интервал после: {actualAfter:F1} pt (требуется {gost.HeaderLineSpacingAfter.Value:F1} pt)");
                        paragraphHasErrors = true;
                    }
                }


                // Проверка выравнивания заголовка
                if (!string.IsNullOrEmpty(gost.HeaderAlignment))
                {
                    var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);
                    string requiredAlignment = gost.HeaderAlignment ?? DefaultHeaderAlignment;

                    if (currentAlignment != requiredAlignment)
                    {
                        paraErrors.Add($"выравнивание: {currentAlignment} (требуется {requiredAlignment})");
                        paragraphHasErrors = true;
                    }
                }

                if (paragraphHasErrors)
                {
                    string headerText = paragraph.InnerText.Trim();
                    errors.Add($"Заголовок '{headerText}': {string.Join(", ", paraErrors)}");
                    hasErrors = true;
                }
            }

            if (hasErrors)
            {
                string errorMessage = $"Ошибки в заголовках:\n{string.Join("\n", errors.Take(3))}";
                if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";
                ShowHeaderError(errorMessage);
            }
            else
            {
                ShowHeaderSuccess("Интервалы и выравнивание в заголовках соответствуют ГОСТу");
            }

            return !hasErrors;
        }

        /// <summary>
        /// Проверяет соответствие отступов заголовков требованиям ГОСТа
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckHeaderIndents(List<Paragraph> paragraphs, Gost gost)
        {
            if (!gost.HeaderIndentLeft.HasValue && !gost.HeaderIndentRight.HasValue && !gost.HeaderFirstLineIndent.HasValue)
            {
                Dispatcher.UIThread.Post(() => {

                    ErrorControlHeaderIndents.Text = "Проверка отступов заголовков не требуется";
                    ErrorControlHeaderIndents.Foreground = Brushes.Gray;

                });
                return true;
            }

            bool isValid = true;
            var errors = new List<string>();
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                if (!headerTexts.Contains(paragraph.InnerText.Trim()))
                    continue;

                var indent = paragraph.ParagraphProperties?.Indentation;
                bool hasError = false;
                var errorDetails = new List<string>();

                // Проверка левого отступа
                if (gost.HeaderIndentLeft.HasValue)
                {
                    double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultHeaderLeftIndent;

                    if (Math.Abs(actualLeft - gost.HeaderIndentLeft.Value) > 0.05)
                    {
                        errorDetails.Add($"левый отступ: {actualLeft:F2} см (требуется {gost.HeaderIndentLeft.Value:F2} см)");
                        hasError = true;
                    }
                }

                // Проверка правого отступа
                if (gost.HeaderIndentRight.HasValue)
                {
                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultHeaderRightIndent;

                    if (Math.Abs(actualRight - gost.HeaderIndentRight.Value) > 0.05)
                    {
                        errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {gost.HeaderIndentRight.Value:F2} см)");
                        hasError = true;
                    }
                }

                // Проверка отступа/выступа первой строки
                if (gost.HeaderFirstLineIndent.HasValue)
                {
                    string currentType = DefaultHeaderFirstLineType;
                    double? currentValue = null;

                    if (indent?.Hanging != null)
                    {
                        currentType = "выступ";
                        currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                    }
                    else if (indent?.FirstLine != null)
                    {
                        currentType = "отступ";
                        currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                    }

                    if (!string.IsNullOrEmpty(gost.HeaderIndentOrOutdent) && gost.HeaderIndentOrOutdent != "нет")
                    {
                        string requiredType = gost.HeaderIndentOrOutdent == "Выступ" ? "выступ" : "отступ";

                        if (currentType != requiredType)
                        {
                            errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                            hasError = true;
                        }
                    }

                    if (currentValue.HasValue)
                    {
                        if (Math.Abs(currentValue.Value - gost.HeaderFirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gost.HeaderFirstLineIndent.Value:F2} см)");
                            hasError = true;
                        }
                    }
                    else if (!string.IsNullOrEmpty(gost.HeaderIndentOrOutdent) && gost.HeaderIndentOrOutdent != "нет")
                    {
                        errorDetails.Add($"отсутствует {gost.HeaderIndentOrOutdent} первой строки");
                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string headerText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                    errors.Add($"Заголовок '{headerText}': {string.Join(", ", errorDetails)}");
                    isValid = false;
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (errors.Any())
                {
                    string errorMessage = $"Ошибки в отступах заголовков:\n{string.Join("\n", errors.Take(3))}";
                    if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlHeaderIndents.Text = errorMessage;
                    ErrorControlHeaderIndents.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlHeaderIndents.Text = "Отступы заголовков соответствуют ГОСТу";
                    ErrorControlHeaderIndents.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Выделение обязательных разделов в заголовках
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <param name="errors"></param>
        private void HighlightHeaderErrors(WordprocessingDocument doc, Gost gost, List<string> errors)
        {
            var body = doc.MainDocumentPart.Document.Body;
            bool hasAnyErrors = false;

            // Получаем все стили документа
            var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
            var styles = stylesPart?.Styles?.Elements<Style>().ToDictionary(s => s.StyleId.Value);

            // Получаем список обязательных заголовков
            var requiredSections = GetRequiredSectionsList(gost);
            var normalizedSections = requiredSections.Select(s => Regex.Replace(s, @"^[\d\.\s]+", "").Trim()).ToList();

            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                var paragraphText = paragraph.InnerText.Trim();
                string cleanText = Regex.Replace(paragraphText, @"^[\d\.\s]+", "").Trim();
                bool isHeader = requiredSections.Contains(paragraphText) || normalizedSections.Contains(cleanText) || IsHeaderByStyle(paragraph, styles);

                if (!isHeader) continue;

                bool hasError = false;
                var errorDetails = new List<string>();
                var indent = paragraph.ParagraphProperties?.Indentation;
                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                var justification = paragraph.ParagraphProperties?.Justification;
                Style style = null;
                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

                if (styleId != null && styles != null && styles.TryGetValue(styleId, out var s))
                {
                    style = s;
                }

                // Проверка шрифта
                if (!string.IsNullOrEmpty(gost.HeaderFontName))
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                        var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? style?.StyleRunProperties?.RunFonts?.Ascii?.Value;

                        if (font != null && !string.Equals(font, gost.HeaderFontName, StringComparison.OrdinalIgnoreCase))
                        {
                            errorDetails.Add($"неверный шрифт: '{font}' (требуется '{gost.HeaderFontName}')");
                            hasError = true;
                            HighlightRun(run);  
                        }
                    }
                }

                // Проверка размера шрифта
                if (gost.HeaderFontSize.HasValue)
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                        string fontSizeVal = run.RunProperties?.FontSize?.Val?.Value ?? style?.StyleRunProperties?.FontSize?.Val?.Value;
                        if (!string.IsNullOrEmpty(fontSizeVal))
                        {
                            if (double.TryParse(fontSizeVal, out var parsedSize))
                            {
                                double size = parsedSize / 2.0;
                                if (Math.Abs(size - gost.HeaderFontSize.Value) > 0.1)
                                {
                                    errorDetails.Add($"неверный размер шрифта: {size:F1} pt (требуется {gost.HeaderFontSize.Value:F1} pt)");
                                    hasError = true;
                                    HighlightRun(run);
                                }
                            }
                        }
                    }
                }

                // Проверка межстрочного интервала
                if (gost.HeaderLineSpacingValue.HasValue || gost.HeaderLineSpacingBefore.HasValue || gost.HeaderLineSpacingAfter.HasValue)
                {
                    var explicitSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    var styleSpacing = style?.StyleParagraphProperties?.SpacingBetweenLines;
                    double actualSpacing = DefaultHeaderLineSpacingValue;
                    string actualSpacingType = DefaultHeaderLineSpacingType;
                    LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                    var spacingSource = explicitSpacing ?? styleSpacing;

                    if (spacingSource?.Line != null)
                    {
                        if (spacingSource.LineRule?.Value == LineSpacingRuleValues.Exact)
                        {
                            actualSpacing = double.Parse(spacingSource.Line.Value) / 20.0;
                            actualSpacingType = "Точно";
                            actualRule = LineSpacingRuleValues.Exact;
                        }
                        else if (spacingSource.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                        {
                            actualSpacing = double.Parse(spacingSource.Line.Value) / 20.0;
                            actualSpacingType = "Минимум";
                            actualRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            actualSpacing = double.Parse(spacingSource.Line.Value) / 240.0;
                            actualSpacingType = "Множитель";
                            actualRule = LineSpacingRuleValues.Auto;
                        }
                    }

                    LineSpacingRuleValues requiredRule = (gost.HeaderLineSpacingType ?? DefaultHeaderLineSpacingType) switch
                    {
                        "Минимум" => LineSpacingRuleValues.AtLeast,
                        "Точно" => LineSpacingRuleValues.Exact,
                        _ => LineSpacingRuleValues.Auto
                    };

                    string requiredType = gost.HeaderLineSpacingType ?? DefaultHeaderLineSpacingType;

                    if (actualRule != requiredRule)
                    {
                        errorDetails.Add($"тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                        hasError = true;
                    }

                    if (Math.Abs(actualSpacing - (gost.HeaderLineSpacingValue ?? DefaultHeaderLineSpacingValue)) > 0.1)
                    {
                        errorDetails.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {(gost.HeaderLineSpacingValue ?? DefaultHeaderLineSpacingValue):F2})");
                        hasError = true;
                    }
                }

                // Проверка выравнивания
                if (!string.IsNullOrEmpty(gost.HeaderAlignment))
                {
                    var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);
                    string requiredAlignment = gost.HeaderAlignment ?? DefaultHeaderAlignment;

                    if (currentAlignment != requiredAlignment)
                    {
                        errorDetails.Add($"выравнивание: {currentAlignment} (требуется {requiredAlignment})");
                        hasError = true;
                    }
                }

                // Проверка отступа первой строки
                if (gost.HeaderFirstLineIndent.HasValue)
                {
                    string currentType = DefaultHeaderFirstLineType;
                    double? currentValue = null;

                    if (indent?.Hanging != null)
                    {
                        currentType = "выступ";
                        currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                    }
                    else if (indent?.FirstLine != null)
                    {
                        currentType = "отступ";
                        currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                    }

                    if (!string.IsNullOrEmpty(gost.HeaderIndentOrOutdent) && gost.HeaderIndentOrOutdent != "нет")
                    {
                        string requiredType = gost.HeaderIndentOrOutdent == "Выступ" ? "выступ" : "отступ";

                        if (currentType != requiredType)
                        {
                            errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                            hasError = true;
                        }
                    }

                    if (currentValue.HasValue)
                    {
                        if (Math.Abs(currentValue.Value - gost.HeaderFirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gost.HeaderFirstLineIndent.Value:F2} см)");
                            hasError = true;
                        }
                    }
                    else if (!string.IsNullOrEmpty(gost.HeaderIndentOrOutdent) && gost.HeaderIndentOrOutdent != "нет")
                    {
                        errorDetails.Add($"отсутствует {gost.HeaderIndentOrOutdent} первой строки");
                        hasError = true;
                    }
                }

                // Проверка левого отступа
                if (gost.HeaderIndentLeft.HasValue)
                {
                    double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultHeaderLeftIndent;

                    if (Math.Abs(actualLeft - gost.HeaderIndentLeft.Value) > 0.05)
                    {
                        errorDetails.Add($"левый отступ: {actualLeft:F2} см (требуется {gost.HeaderIndentLeft.Value:F2} см)");
                        hasError = true;
                    }
                }

                // Проверка правого отступа
                if (gost.HeaderIndentRight.HasValue)
                {
                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultHeaderRightIndent;

                    if (Math.Abs(actualRight - gost.HeaderIndentRight.Value) > 0.05)
                    {
                        errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {gost.HeaderIndentRight.Value:F2} см)");
                        hasError = true;
                    }
                }

                // Проверка интервалов "Перед" и "После"
                if (gost.HeaderLineSpacingBefore.HasValue || gost.HeaderLineSpacingAfter.HasValue)
                {
                    double actualBefore = DefaultHeaderSpacingBefore;
                    double actualAfter = DefaultHeaderSpacingAfter;

                    if (spacing?.Before?.Value != null)
                    {
                        actualBefore = ConvertTwipsToPoints(spacing.Before.Value);
                    }

                    if (gost.HeaderLineSpacingBefore.HasValue &&
                        Math.Abs(actualBefore - gost.HeaderLineSpacingBefore.Value) > 0.1)
                    {
                        errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.HeaderLineSpacingBefore.Value:F1} pt)");
                        hasError = true;
                    }

                    if (spacing?.After?.Value != null)
                    {
                        actualAfter = ConvertTwipsToPoints(spacing.After.Value);
                    }

                    if (gost.HeaderLineSpacingAfter.HasValue &&
                        Math.Abs(actualAfter - gost.HeaderLineSpacingAfter.Value) > 0.1)
                    {
                        errorDetails.Add($"интервал после: {actualAfter:F1} pt (требуется {gost.HeaderLineSpacingAfter.Value:F1} pt)");
                        hasError = true;
                    }
                }


                if (hasError)
                {
                    hasAnyErrors = true;
                    string shortText = paragraphText.Length > 50 ? paragraphText.Substring(0, 47) + "..." : paragraphText;
                    errors.Add($"Заголовок '{shortText}': {string.Join(", ", errorDetails.Distinct())}");

                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (!string.IsNullOrWhiteSpace(run.InnerText))
                        {
                            if (run.RunProperties == null)
                            {
                                run.RunProperties = new RunProperties();
                            }

                            run.RunProperties.RemoveAllChildren<Highlight>();
                            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                        }
                    }
                }
            }

            if (hasAnyErrors)
            {
                doc.MainDocumentPart.Document.Save();
            }
        }

        /// <summary>
        /// Вспомогательный метод для проверки, является ли параграф заголовком по стилю
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="styles"></param>
        /// <returns></returns>
        private bool IsHeaderByStyle(Paragraph paragraph, Dictionary<string, Style> styles)
        {
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId == null || styles == null) return false;

            if (styles.TryGetValue(styleId, out var style))
            {
                return style.Type == StyleValues.Paragraph && (style.StyleName?.Val?.Value?.Contains("Heading") == true || (style.StyleParagraphProperties?.OutlineLevel?.Val?.Value ?? 10) <= 1);
            }
            return false;
        }

        /// <summary>
        /// Проверяет соответствие отступов оглавления требованиям ГОСТа
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckTocIndents(WordprocessingDocument doc, Gost gost)
        {
            // Если параметры отступов не заданы, проверка не требуется
            if (!gost.TocIndentLeft.HasValue && !gost.TocIndentRight.HasValue && !gost.TocFirstLineIndent.HasValue)
            {
                Dispatcher.UIThread.Post(() => {

                    ErrorControlTocIndents.Text = "Проверка отступов оглавления не требуется";
                    ErrorControlTocIndents.Foreground = Brushes.Gray;

                });
                return true;
            }

            bool isValid = true;
            var errors = new List<string>();

            // 2. Поиск оглавления всеми доступными способами
            var body = doc.MainDocumentPart.Document.Body;

            // Основной поиск по полю TOC
            var tocField = body.Descendants<FieldCode>().FirstOrDefault(f => f.Text.Contains(" TOC ") || f.Text.Contains("TOC \\"));

            // Поиск по стилям и характерным признакам
            var tocParagraphs = body.Descendants<Paragraph>().Where(IsTocParagraph).ToList();

            if (tocField == null)
            {
                Dispatcher.UIThread.Post(() => {

                    ErrorControlTocIndents.Text = "Автоматическое оглавление не найдено";
                    ErrorControlTocIndents.Foreground = Brushes.Red;

                });
                return false;
            }

            var tocContainer = tocField.Ancestors<Table>().FirstOrDefault() ?? tocField.Ancestors<Paragraph>().FirstOrDefault()?.Parent;

            if (tocContainer == null) return false;

            foreach (var paragraph in tocContainer.Descendants<Paragraph>())
            {
                if (IsEmptyParagraph(paragraph)) continue;

                var indent = paragraph.ParagraphProperties?.Indentation;
                bool hasError = false;
                var errorDetails = new List<string>();

                // 1. Проверка левого отступа оглавления
                if (gost.TocIndentLeft.HasValue)
                {
                    double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultTocLeftIndent;

                    if (Math.Abs(actualLeft - gost.TocIndentLeft.Value) > 0.05)
                    {
                        errorDetails.Add($"левый отступ: {actualLeft:F2} см (требуется {gost.TocIndentLeft.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 2. Проверка правого отступа оглавления
                if (gost.TocIndentRight.HasValue)
                {
                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTocRightIndent;

                    if (Math.Abs(actualRight - gost.TocIndentRight.Value) > 0.05)
                    {
                        errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {gost.TocIndentRight.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 3. Проверка отступа/выступа первой строки оглавления
                if (gost.TocFirstLineIndent.HasValue)
                {
                    string currentType = "не определен";
                    double? currentValue = null;

                    if (indent?.Hanging != null)
                    {
                        currentType = "выступ";
                        currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                    }
                    else if (indent?.FirstLine != null)
                    {
                        currentType = "отступ";
                        currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                    }

                    // Проверка типа отступа (отступ/выступ)
                    if (!string.IsNullOrEmpty(gost.TocIndentOrOutdent) && gost.TocIndentOrOutdent != "нет")
                    {
                        string requiredType = gost.TocIndentOrOutdent == "Выступ" ? "выступ" : "отступ";

                        if (currentType != requiredType)
                        {
                            errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                            hasError = true;
                        }
                    }

                    // Проверка значения отступа
                    if (currentValue.HasValue)
                    {
                        if (Math.Abs(currentValue.Value - gost.TocFirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gost.TocFirstLineIndent.Value:F2} см)");
                            hasError = true;
                        }
                    }
                    else if (!string.IsNullOrEmpty(gost.TocIndentOrOutdent) && gost.TocIndentOrOutdent != "нет")
                    {
                        errorDetails.Add($"отсутствует {gost.TocIndentOrOutdent} первой строки");
                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string tocText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                    errors.Add($"Оглавление '{tocText}': {string.Join(", ", errorDetails)}");
                    isValid = false;
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (errors.Any())
                {
                    string errorMessage = $"Ошибки в отступах оглавления:\n{string.Join("\n", errors.Take(3))}";
                    if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";

                    ErrorControlTocIndents.Text = errorMessage;
                    ErrorControlTocIndents.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlTocIndents.Text = "Отступы оглавления соответствуют ГОСТу";
                    ErrorControlTocIndents.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Метод проверки, является ли параграф частью оглавления
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsTocParagraph(Paragraph paragraph)
        {
            if (paragraph == null) return false;

            // 1. Проверка по полю TOC (автоматическое оглавление)
            if (paragraph.Descendants<FieldCode>().Any(f => f.Text.Contains(" TOC ", StringComparison.OrdinalIgnoreCase) ||
                                                            f.Text.Contains("TOC \\", StringComparison.OrdinalIgnoreCase)))
                return true;

            // 2. Проверка по стилю
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(style))
            {
                style = style.ToLower();
                if (style.Contains("toc") || style.Contains("contents") || style.Contains("table"))
                    return true;
            }

            // 3. Проверка по характерным признакам
            string text = paragraph.InnerText;
            if (text.Contains(".........") || text.Contains("\t") || Regex.IsMatch(text, @"\.{3,}\s*\d+$")) 
                return true;

            // 4. Проверка по родительскому элементу (таблица оглавления)
            var parentTable = paragraph.Ancestors<Table>().FirstOrDefault();
            if (parentTable != null)
            {
                return parentTable.Descendants<FieldCode>().Any(f => f.Text.Contains(" TOC ", StringComparison.OrdinalIgnoreCase) ||
                                                                     f.Text.Contains("TOC \\", StringComparison.OrdinalIgnoreCase));
            }


            return false;
        }

        /// <summary>
        /// Проверяет соответствие интервалов в списках требованиям ГОСТа
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckListParagraphSpacing(List<Paragraph> paragraphs, Gost gost)
        {
            var lineSpacingTypeNames = new Dictionary<LineSpacingRuleValues, string>
            {
                { LineSpacingRuleValues.Auto, "Множитель" },
                { LineSpacingRuleValues.AtLeast, "Минимум" },
                { LineSpacingRuleValues.Exact, "Точно" }
            };

            bool isValid = true;
            var errors = new List<string>();

            foreach (var paragraph in paragraphs)
            {
                if (!IsListItem(paragraph)) continue;

                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                bool hasError = false;
                var errorDetails = new List<string>();

                // Проверка межстрочного интервала
                if (gost.BulletLineSpacingValue.HasValue)
                {
                    double actualSpacing = DefaultListLineSpacingValue;
                    string actualSpacingType = DefaultListLineSpacingType;
                    LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                    if (spacing?.Line != null)
                    {
                        if (spacing.LineRule?.Value == LineSpacingRuleValues.Exact)
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 20.0;
                            actualSpacingType = "Точно";
                            actualRule = LineSpacingRuleValues.Exact;
                        }
                        else if (spacing.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 20.0;
                            actualSpacingType = "Минимум";
                            actualRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 240.0;
                            actualSpacingType = "Множитель";
                            actualRule = LineSpacingRuleValues.Auto;
                        }
                    }

                    // Определяем требуемый тип интервала
                    LineSpacingRuleValues requiredRule = (gost.BulletLineSpacingType ?? DefaultListLineSpacingType) switch
                    {
                        "Минимум" => LineSpacingRuleValues.AtLeast,
                        "Точно" => LineSpacingRuleValues.Exact,
                        _ => LineSpacingRuleValues.Auto
                    };

                    string requiredType = gost.BulletLineSpacingType ?? DefaultListLineSpacingType;

                    // Проверка типа интервала
                    if (actualRule != requiredRule)
                    {
                        errorDetails.Add($"тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                        hasError = true;
                    }

                    // Проверка значения интервала
                    double requiredSpacingValue = gost.BulletLineSpacingValue ?? DefaultListLineSpacingValue;
                    if (Math.Abs(actualSpacing - requiredSpacingValue) > 0.01)
                    {
                        errorDetails.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {requiredSpacingValue:F2})");
                        hasError = true;
                    }
                }

                // Проверка интервалов перед/после
                if (gost.BulletLineSpacingBefore.HasValue || gost.BulletLineSpacingAfter.HasValue)
                {
                    double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultListSpacingBefore;
                    double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultListSpacingAfter;

                    if (gost.BulletLineSpacingBefore.HasValue && Math.Abs(actualBefore - gost.BulletLineSpacingBefore.Value) > 0.01)
                    {
                        errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.BulletLineSpacingBefore.Value:F1} pt)");
                        hasError = true;
                    }

                    if (gost.BulletLineSpacingAfter.HasValue && Math.Abs(actualAfter - gost.BulletLineSpacingAfter.Value) > 0.01)
                    {
                        errorDetails.Add($"интервал после: {actualAfter:F1} pt (требуется {gost.BulletLineSpacingAfter.Value:F1} pt)");
                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string shortText = GetShortText(paragraph);
                    errors.Add($"Список '{shortText}': {string.Join(", ", errorDetails)}");
                    isValid = false;
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (errors.Any())
                {
                    string errorMessage = $"Ошибки в интервалах списков:\n{string.Join("\n", errors.Take(3))}";
                    if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlListSpacing.Text = errorMessage;
                    ErrorControlListSpacing.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlListSpacing.Text = "Интервалы списков соответствуют ГОСТу";
                    ErrorControlListSpacing.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Метод проверки отступов списков
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckListIndents(List<Paragraph> paragraphs, Gost gost)
        {
            if (!gost.ListHangingIndent.HasValue && !gost.ListLevel1Indent.HasValue && !gost.ListLevel2Indent.HasValue &&
                !gost.ListLevel3Indent.HasValue && !gost.ListLevel4Indent.HasValue && !gost.ListLevel5Indent.HasValue &&
                !gost.ListLevel6Indent.HasValue && !gost.ListLevel7Indent.HasValue && !gost.ListLevel8Indent.HasValue &&
                !gost.ListLevel9Indent.HasValue && !gost.BulletIndentLeft.HasValue && !gost.BulletIndentRight.HasValue)
            {
                Dispatcher.UIThread.Post(() => {

                    ErrorControlListIndents.Text = "Проверка отступов списков не требуется";
                    ErrorControlListIndents.Foreground = Brushes.Gray;

                });
                return true;
            }

            bool isValid = true;
            var errors = new List<string>();

            foreach (var paragraph in paragraphs)
            {
                if (!IsStrictListItem(paragraph)) continue;

                int level = GetListLevel(paragraph, gost);
                var indent = paragraph.ParagraphProperties?.Indentation;
                bool hasError = false;
                var errorDetails = new List<string>();

                // 1. Получаем ТРЕБУЕМЫЕ значения из ГОСТа для текущего уровня
                double? gostRequiredIndent = level switch
                {
                    1 => gost.ListLevel1Indent,
                    2 => gost.ListLevel2Indent,
                    3 => gost.ListLevel3Indent,
                    4 => gost.ListLevel4Indent,
                    5 => gost.ListLevel5Indent,
                    6 => gost.ListLevel6Indent,
                    7 => gost.ListLevel7Indent,
                    8 => gost.ListLevel8Indent,
                    9 => gost.ListLevel9Indent,
                    _ => null
                };

                // Если для уровня нет специфичного требования, используем общее значение
                if (!gostRequiredIndent.HasValue)
                {
                    gostRequiredIndent = gost.ListHangingIndent;
                }

                // 2. Получаем ФАКТИЧЕСКИЕ значения из документа
                string currentType = "не определен";
                double? currentValue = null;

                if (indent?.Hanging != null)
                {
                    currentType = "выступ";
                    currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                }
                else if (indent?.FirstLine != null)
                {
                    currentType = "отступ";
                    currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                }
                else // Если в документе не заданы отступы, используем значения по умолчанию
                {
                    string defaultType = GetRequiredIndentType(gost, level).ToLower();
                    currentType = defaultType;
                    currentValue = GetListLevelIndent(level);
                }

                // 3. Проверяем только если в ГОСТе есть требования для отступов
                if (gostRequiredIndent.HasValue)
                {
                    string requiredType = GetRequiredIndentType(gost, level).ToLower();// Получаем требуемый тип отступа из ГОСТа

                    if (!string.IsNullOrEmpty(requiredType) && currentType != requiredType)
                    {
                        errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                        hasError = true;
                    }

                    if (currentValue.HasValue && Math.Abs(currentValue.Value - gostRequiredIndent.Value) > 0.05)
                    {
                        errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gostRequiredIndent.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 4. Проверка левого отступа 
                if (gost.BulletIndentLeft.HasValue)
                {
                    double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : GetListLevelIndentLeft(level);

                    if (Math.Abs(actualLeft - gost.BulletIndentLeft.Value) > 0.05)
                    {
                        errorDetails.Add($"Левый отступ: {actualLeft:F2} см (требуется {gost.BulletIndentLeft.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 5. Проверка правого отступа 
                if (gost.BulletIndentRight.HasValue)
                {
                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : GetListLevelIndentRight(level);

                    if (Math.Abs(actualRight - gost.BulletIndentRight.Value) > 0.05)
                    {
                        errorDetails.Add($"Правый отступ: {actualRight:F2} см (требуется {gost.BulletIndentRight.Value:F2} см)");
                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string shortText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                    errors.Add($"Список ур. {level} '{shortText}': {string.Join(", ", errorDetails)}");
                    isValid = false;

                    // Выделение ошибки
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (!string.IsNullOrWhiteSpace(run.InnerText))
                        {
                            run.RunProperties ??= new RunProperties();
                            run.RunProperties.RemoveAllChildren<Highlight>();// Удаляем выделение
                            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });// Красное выделение
                        }
                    }
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (errors.Any())
                {
                    string errorMessage = $"Ошибки в отступах списков:\n{string.Join("\n", errors.Take(3))}";
                    if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlListIndents.Text = errorMessage;
                    ErrorControlListIndents.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlListIndents.Text = "Отступы списков соответствуют ГОСТу";
                    ErrorControlListIndents.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Выделяет ошибки в списках на основе существующих проверок
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <param name="errors"></param>
        private void HighlightListErrors(WordprocessingDocument doc, Gost gost, List<string> errors)
        {
            var body = doc.MainDocumentPart.Document.Body;
            bool hasAnyErrors = false;

            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                if (!IsListItem(paragraph)) continue;

                bool hasError = false;
                var errorDetails = new List<string>();
                int level = GetListLevel(paragraph, gost);

                // 1. Проверка формата нумерации (из CheckBulletedLists)
                if (IsNumberedList(paragraph))
                {
                    string? requiredFormat = level switch
                    {
                        1 => gost.ListLevel1NumberFormat,
                        2 => gost.ListLevel2NumberFormat,
                        3 => gost.ListLevel3NumberFormat,
                        4 => gost.ListLevel4NumberFormat,
                        5 => gost.ListLevel5NumberFormat,
                        6 => gost.ListLevel6NumberFormat,
                        7 => gost.ListLevel7NumberFormat,
                        8 => gost.ListLevel8NumberFormat,
                        9 => gost.ListLevel9NumberFormat,
                        _ => null
                    };

                    if (!string.IsNullOrEmpty(requiredFormat))
                    {
                        var firstRunText = paragraph.Elements<Run>().FirstOrDefault()?.InnerText?.Trim();
                        if (firstRunText != null && !CheckNumberFormat(firstRunText, requiredFormat))
                        {
                            errorDetails.Add($"Неверный формат нумерации уровня {level}: '{firstRunText}' (требуется '{requiredFormat}')");
                            hasError = true;
                        }
                    }
                }

                // 2. Проверка шрифта и размера (из CheckBulletedLists)
                if (!string.IsNullOrEmpty(gost.BulletFontName) || gost.BulletFontSize.HasValue)
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                        // Проверка шрифта
                        if (!string.IsNullOrEmpty(gost.BulletFontName))
                        {
                            var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? DefaultListFont;
                            if (font != gost.BulletFontName)
                            {
                                errorDetails.Add($"Неверный шрифт списка: '{font}' (требуется '{gost.BulletFontName}')");
                                hasError = true;
                                break;
                            }
                        }

                        // Проверка размера шрифта
                        if (gost.BulletFontSize.HasValue)
                        {
                            var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value;
                            double actualSize = fontSizeVal != null ? double.Parse(fontSizeVal) / 2 : DefaultListSize;

                            if (Math.Abs(actualSize - gost.BulletFontSize.Value) > 0.1)
                            {
                                errorDetails.Add($"Неверный размер шрифта: {actualSize:F1} pt (требуется {gost.BulletFontSize.Value:F1} pt)");
                                hasError = true;
                                break;
                            }
                        }
                    }
                }

                // 3. Проверка отступов (из CheckListIndents)
                var indent = paragraph.ParagraphProperties?.Indentation;
                if (gost.BulletIndentLeft.HasValue || gost.BulletIndentRight.HasValue || gost.ListHangingIndent.HasValue || (level >= 1 && level <= 9))
                {
                    // Левый отступ
                    if (gost.BulletIndentLeft.HasValue && indent?.Left != null)
                    {
                        double actualLeft = TwipsToCm(double.Parse(indent.Left.Value));
                        if (Math.Abs(actualLeft - gost.BulletIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"Левый отступ списка: {actualLeft:F2} см (требуется {gost.BulletIndentLeft.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // Правый отступ
                    if (gost.BulletIndentRight.HasValue && indent?.Right != null)
                    {
                        double actualRight = TwipsToCm(double.Parse(indent.Right.Value));
                        if (Math.Abs(actualRight - gost.BulletIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"Правый отступ списка: {actualRight:F2} см (требуется {gost.BulletIndentRight.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // Отступ/выступ первой строки
                    if (gost.ListHangingIndent.HasValue || (level >= 1 && level <= 9))
                    {
                        double? requiredIndent = level switch
                        {
                            1 => gost.ListLevel1Indent ?? gost.ListHangingIndent,
                            2 => gost.ListLevel2Indent ?? gost.ListHangingIndent,
                            3 => gost.ListLevel3Indent ?? gost.ListHangingIndent,
                            4 => gost.ListLevel4Indent ?? gost.ListHangingIndent,
                            5 => gost.ListLevel5Indent ?? gost.ListHangingIndent,
                            6 => gost.ListLevel6Indent ?? gost.ListHangingIndent,
                            7 => gost.ListLevel7Indent ?? gost.ListHangingIndent,
                            8 => gost.ListLevel8Indent ?? gost.ListHangingIndent,
                            9 => gost.ListLevel9Indent ?? gost.ListHangingIndent,
                            _ => gost.ListHangingIndent
                        };

                        if (requiredIndent.HasValue)
                        {
                            bool indentValid = false;
                            if (indent?.Hanging != null &&
                                Math.Abs(TwipsToCm(double.Parse(indent.Hanging.Value)) - requiredIndent.Value) <= 0.05)
                            {
                                indentValid = true;
                            }
                            else if (indent?.FirstLine != null &&
                                     Math.Abs(TwipsToCm(double.Parse(indent.FirstLine.Value)) - requiredIndent.Value) <= 0.05)
                            {
                                indentValid = true;
                            }

                            if (!indentValid)
                            {
                                errorDetails.Add($"Неверный отступ/выступ: требуется {requiredIndent.Value:F2} см");
                                hasError = true;
                            }
                        }
                    }
                }

                // 4. Проверка интервалов (из CheckListParagraphSpacing)
                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                if (gost.BulletLineSpacingBefore.HasValue || gost.BulletLineSpacingAfter.HasValue || gost.BulletLineSpacingValue.HasValue)
                {
                    if (gost.BulletLineSpacingBefore.HasValue)
                    {
                        double beforeValue = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultListSpacingBefore;

                        if (Math.Abs(beforeValue - gost.BulletLineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"Интервал перед: {beforeValue:F1} pt (требуется {gost.BulletLineSpacingBefore.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    if (gost.BulletLineSpacingAfter.HasValue)
                    {
                        double afterValue = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultListSpacingAfter;

                        if (Math.Abs(afterValue - gost.BulletLineSpacingAfter.Value) > 0.01)
                        {
                            errorDetails.Add($"Интервал после: {afterValue:F1} pt (требуется {gost.BulletLineSpacingAfter.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    if (gost.BulletLineSpacingValue.HasValue && spacing?.Line != null)
                    {
                        double actualSpacing = spacing.LineRule?.Value == LineSpacingRuleValues.Auto ? double.Parse(spacing.Line.Value) / 240.0 : double.Parse(spacing.Line.Value) / 20.0;

                        if (Math.Abs(actualSpacing - gost.BulletLineSpacingValue.Value) > 0.01)
                        {
                            errorDetails.Add($"Межстрочный интервал: {actualSpacing:F2} (требуется {gost.BulletLineSpacingValue.Value:F2})");
                            hasError = true;
                        }
                    }
                }

                if (hasError)
                {
                    string shortText = GetShortText(paragraph);
                    errors.Add($"Список ур. {level} '{shortText}': {string.Join(", ", errorDetails)}");
                    hasAnyErrors = true;

                    // Выделение всех Run элементов в параграфе
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (!string.IsNullOrWhiteSpace(run.InnerText))
                        {
                            run.RunProperties ??= new RunProperties();
                            run.RunProperties.RemoveAllChildren<Highlight>();
                            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                        }
                    }
                }
            }

            if (hasAnyErrors)
            {
                doc.MainDocumentPart.Document.Save();
            }
        }

        /// <summary>
        /// Строгая проверка определяющая что параграф является элементом списка
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsStrictListItem(Paragraph paragraph)
        {
            // 1. Проверка явных свойств нумерации
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // 2. Проверка стилей списка
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) && (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering")))
                return true;

            // 3. Проверка по содержимому (маркеры или нумерация)
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                string text = firstRun.InnerText?.Trim() ?? "";

                // Маркированные списки
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—"))
                    return true;

                // Нумерованные списки
                if (Regex.IsMatch(text, @"^\d+[\.\)]") ||   // 1. 1) 
                    Regex.IsMatch(text, @"^[a-z]\)") ||     // a) b)
                    Regex.IsMatch(text, @"^[IVXLCDM]+\.", RegexOptions.IgnoreCase))  // I. II.
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверка базовых параметров списков
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckBulletedLists(List<Paragraph> paragraphs, Gost gost)
        {
            var errors = new List<string>();
            bool listsValid = true;

            foreach (var paragraph in paragraphs)
            {
                if (!IsListItem(paragraph)) continue;

                bool paragraphHasError = false;
                var runsWithText = paragraph.Elements<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)).ToList();

                // Проверка формата нумерации
                if (IsNumberedList(paragraph))
                {
                    int level = GetListLevel(paragraph, gost);
                    string? requiredFormat = level switch
                    {
                        1 => gost.ListLevel1NumberFormat,
                        2 => gost.ListLevel2NumberFormat,
                        3 => gost.ListLevel3NumberFormat,
                        4 => gost.ListLevel4NumberFormat,
                        5 => gost.ListLevel5NumberFormat,
                        6 => gost.ListLevel6NumberFormat,
                        7 => gost.ListLevel7NumberFormat,
                        8 => gost.ListLevel8NumberFormat,
                        9 => gost.ListLevel9NumberFormat,
                        _ => null
                    };

                    if (!string.IsNullOrEmpty(requiredFormat))
                    {
                        var firstRunText = runsWithText.FirstOrDefault()?.InnerText?.Trim();
                        if (firstRunText != null && !CheckNumberFormat(firstRunText, requiredFormat))
                        {
                            errors.Add($"Неверный формат нумерации уровня {level}: '{firstRunText}' (требуется '{requiredFormat}')");
                            paragraphHasError = true;
                        }
                    }
                }

                // Проверка типа шрифта
                if (!string.IsNullOrEmpty(gost.BulletFontName))
                {
                    foreach (var run in runsWithText)
                    {
                        var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? DefaultListFont;
                        if (font != null && font != gost.BulletFontName)
                        {
                            errors.Add($"Неверный шрифт списка: '{font}' (требуется '{gost.BulletFontName}')");
                            paragraphHasError = true;
                            break;
                        }
                    }
                }

                // Проверка размера шрифта
                if (gost.BulletFontSize.HasValue)
                {
                    foreach (var run in runsWithText)
                    {
                        var fontSize = run.RunProperties?.FontSize?.Val?.Value ?? DefaultListSize.ToString();

                        if (fontSize != null)
                        {
                            double actualSize = -1;
                            var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value;

                            if (fontSizeVal != null)
                            {
                                actualSize = double.Parse(fontSizeVal) / 2;
                            }
                            else
                            {
                                actualSize = DefaultTocSize;
                            }

                            if (Math.Abs(actualSize - gost.BulletFontSize.Value) > 0.1)
                            {
                                errors.Add($"Неверный размер шрифта: {actualSize}pt (требуется {gost.BulletFontSize.Value}pt) в параграфе: '{GetShortText(paragraph)}'");
                                paragraphHasError = true;
                                break;
                            }
                        }
                        else if (gost.BulletFontSize.Value != 0) // 0 - значение по умолчанию
                        {
                            errors.Add("Отсутствует размер шрифта");
                            paragraphHasError = true;
                            break;
                        }
                    }
                }

                if (paragraphHasError)
                {
                    listsValid = false;
                    foreach (var run in runsWithText)
                    {
                        run.RunProperties ??= new RunProperties();
                        run.RunProperties.RemoveAllChildren<Highlight>();
                        run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                    }
                }
            }

            UpdateBulletedListsUI(errors.Distinct().ToList(), listsValid, true);
            return listsValid;
        }

        /// <summary>
        /// Метод проверки оглавления с поиском всех ошибок
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckTableOfContents(WordprocessingDocument doc, Gost gost)
        {
            // 1. Проверка, требуется ли оглавление по ГОСТу
            if (gost.RequireTOC.HasValue && !gost.RequireTOC.Value)
            {
                Dispatcher.UIThread.Post(() => {
                    Error_ControlToc_Spacing.Text = "Оглавление не требуется по ГОСТу";
                    Error_ControlToc_Spacing.Foreground = Brushes.Gray;
                });
                return true;
            }

            // 2. Поиск оглавления всеми доступными способами
            var body = doc.MainDocumentPart.Document.Body;

            // Основной поиск по полю TOC
            var tocField = body.Descendants<FieldCode>().FirstOrDefault(f => f.Text.Contains(" TOC ") || f.Text.Contains("TOC \\"));

            // Поиск по стилям и характерным признакам
            var tocParagraphs = body.Descendants<Paragraph>().Where(IsTocParagraph).ToList();

            // 3. Определение контейнера оглавления
            OpenXmlElement tocContainer = null;

            if (tocField != null)
            {
                tocContainer = tocField.Ancestors<Table>().FirstOrDefault() ?? tocField.Ancestors<Paragraph>().FirstOrDefault()?.Parent;
            }
            else if (tocParagraphs.Any())
            {
                tocContainer = tocParagraphs.First().Parent;
            }

            // 4. Проверка наличия оглавления
            if (tocContainer == null)
            {
                Dispatcher.UIThread.Post(() => {
                    Error_ControlToc_Spacing.Text = "Автоматическое оглавление не найдено! Создайте через 'Ссылки → Оглавление'";
                    Error_ControlToc_Spacing.Foreground = Brushes.Red;
                });
                return false;
            }

            // 5. Проверка форматирования оглавления
            bool hasErrors = false;
            var tocErrors = new List<string>();
            var tocStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles.Elements<Style>().Where(s => s.StyleId?.Value?.StartsWith("TOC") == true).ToDictionary(s => s.StyleId.Value);

            foreach (var paragraph in tocContainer.Descendants<Paragraph>())
            {
                if (IsEmptyParagraph(paragraph)) continue;

                bool paragraphHasError = false;
                var errorDetails = new List<string>();
                var indent = paragraph.ParagraphProperties?.Indentation;
                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                var justification = paragraph.ParagraphProperties?.Justification;

                Style paragraphStyle = null;
                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && tocStyles.TryGetValue(styleId, out var style))
                {
                    paragraphStyle = style;
                }

                // Проверка шрифта
                if (!string.IsNullOrEmpty(gost.TocFontName))
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                        var font = run.RunProperties?.RunFonts?.Ascii?.Value
                                 ?? paragraphStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value;

                        if (font != null && !string.Equals(font, gost.TocFontName, StringComparison.OrdinalIgnoreCase))
                        {
                            errorDetails.Add($"шрифт: '{font}' (требуется '{gost.TocFontName}')");
                            paragraphHasError = true;
                            break;
                        }
                    }
                }

                // Проверка размера шрифта
                if (gost.TocFontSize.HasValue)
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                        var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value
                                        ?? paragraphStyle?.StyleRunProperties?.FontSize?.Val?.Value;

                        double actualSize = fontSizeVal != null
                            ? double.Parse(fontSizeVal) / 2
                            : DefaultTocSize;

                        if (Math.Abs(actualSize - gost.TocFontSize.Value) > 0.1)
                        {
                            errorDetails.Add($"размер: {actualSize:F1} pt (требуется {gost.TocFontSize.Value:F1} pt)");
                            paragraphHasError = true;
                            break;
                        }
                    }
                }

                // Проверка выравнивания
                if (!string.IsNullOrEmpty(gost.TocAlignment))
                {
                    var actualAlignment = GetAlignmentString(justification);
                    if (actualAlignment != gost.TocAlignment)
                    {
                        errorDetails.Add($"выравнивание: {actualAlignment} (требуется {gost.TocAlignment})");
                        paragraphHasError = true;
                    }
                }

                // Проверка отступов
                if (gost.TocIndentLeft.HasValue || gost.TocIndentRight.HasValue ||
                    gost.TocFirstLineIndent.HasValue)
                {
                    // Левый отступ
                    if (gost.TocIndentLeft.HasValue)
                    {
                        double actualLeft = indent?.Left?.Value != null
                            ? TwipsToCm(double.Parse(indent.Left.Value))
                            : DefaultTocLeftIndent;

                        if (Math.Abs(actualLeft - gost.TocIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"левый отступ: {actualLeft:F2} см (требуется {gost.TocIndentLeft.Value:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    // Правый отступ
                    if (gost.TocIndentRight.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null
                            ? TwipsToCm(double.Parse(indent.Right.Value))
                            : DefaultTocRightIndent;

                        if (Math.Abs(actualRight - gost.TocIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {gost.TocIndentRight.Value:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    // Отступ первой строки
                    if (gost.TocFirstLineIndent.HasValue && gost.TocIndentOrOutdent != "нет")
                    {
                        string currentType = "не определен";
                        double? currentValue = null;

                        if (indent?.Hanging != null)
                        {
                            currentType = "выступ";
                            currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                        }
                        else if (indent?.FirstLine != null)
                        {
                            currentType = "отступ";
                            currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                        }

                        // Проверка типа
                        if (!string.IsNullOrEmpty(gost.TocIndentOrOutdent))
                        {
                            string requiredType = gost.TocIndentOrOutdent == "Выступ" ? "выступ" : "отступ";

                            if (currentType != requiredType)
                            {
                                errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                                paragraphHasError = true;
                            }
                        }

                        // Проверка значения
                        if (currentValue.HasValue)
                        {
                            if (Math.Abs(currentValue.Value - gost.TocFirstLineIndent.Value) > 0.05)
                            {
                                errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gost.TocFirstLineIndent.Value:F2} см)");
                                paragraphHasError = true;
                            }
                        }
                        else
                        {
                            errorDetails.Add($"отсутствует {gost.TocIndentOrOutdent} первой строки");
                            paragraphHasError = true;
                        }
                    }
                }

                // Проверка межстрочного интервала
                if (gost.TocLineSpacing.HasValue)
                {
                    double actualSpacing = DefaultTocLineSpacingValue;
                    string actualSpacingType = DefaultTocLineSpacingType;
                    LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                    if (spacing?.Line != null)
                    {
                        if (spacing.LineRule?.Value == LineSpacingRuleValues.Exact)
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 20.0;
                            actualSpacingType = "Точно";
                            actualRule = LineSpacingRuleValues.Exact;
                        }
                        else if (spacing.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 20.0;
                            actualSpacingType = "Минимум";
                            actualRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 240.0;
                            actualSpacingType = "Множитель";
                            actualRule = LineSpacingRuleValues.Auto;
                        }
                    }

                    // Проверка типа интервала
                    LineSpacingRuleValues requiredRule = (gost.TocLineSpacingType ?? DefaultTocLineSpacingType) switch
                    {
                        "Минимум" => LineSpacingRuleValues.AtLeast,
                        "Точно" => LineSpacingRuleValues.Exact,
                        _ => LineSpacingRuleValues.Auto
                    };

                    if (actualRule != requiredRule)
                    {
                        errorDetails.Add($"тип интервала: '{actualSpacingType}' (требуется '{gost.TocLineSpacingType ?? DefaultTocLineSpacingType}')");
                        paragraphHasError = true;
                    }

                    // Проверка значения интервала
                    if (Math.Abs(actualSpacing - gost.TocLineSpacing.Value) > 0.01)
                    {
                        errorDetails.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {gost.TocLineSpacing.Value:F2})");
                        paragraphHasError = true;
                    }
                }

                // Проверка интервалов перед/после
                if (gost.TocLineSpacingBefore.HasValue || gost.TocLineSpacingAfter.HasValue)
                {
                    if (gost.TocLineSpacingBefore.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null
                            ? ConvertTwipsToPoints(spacing.Before.Value)
                            : DefaultTocSpacingBefore;

                        if (Math.Abs(actualBefore - gost.TocLineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.TocLineSpacingBefore.Value:F1} pt)");
                            paragraphHasError = true;
                        }
                    }

                    if (gost.TocLineSpacingAfter.HasValue)
                    {
                        double actualAfter = spacing?.After?.Value != null
                            ? ConvertTwipsToPoints(spacing.After.Value)
                            : DefaultTocSpacingAfter;

                        if (Math.Abs(actualAfter - gost.TocLineSpacingAfter.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал после: {actualAfter:F1} pt (требуется {gost.TocLineSpacingAfter.Value:F1} pt)");
                            paragraphHasError = true;
                        }
                    }
                }

                if (paragraphHasError)
                {
                    string shortText = GetShortTocText(paragraph);
                    tocErrors.Add($"Оглавление '{shortText}': {string.Join(", ", errorDetails)}");
                    HighlightTocItem(paragraph);
                    hasErrors = true;
                }
            }

            // 6. Вывод результатов
            Dispatcher.UIThread.Post(() => {
                if (hasErrors)
                {
                    string errorMessage = $"Ошибки в оглавлении:\n{string.Join("\n", tocErrors.Take(3))}";
                    if (tocErrors.Count > 3) errorMessage += $"\n...и ещё {tocErrors.Count - 3} ошибок";
                    Error_ControlToc_Spacing.Text = errorMessage;
                    Error_ControlToc_Spacing.Foreground = Brushes.Red;
                }
                else
                {
                    Error_ControlToc_Spacing.Text = "Оглавление полностью соответствует ГОСТу";
                    Error_ControlToc_Spacing.Foreground = Brushes.Green;
                }
            });

            return !hasErrors;
        }

        /// <summary>
        /// Выделяет ошибки в Оглавлении
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        private void HighlightTocErrors(WordprocessingDocument doc, Gost gost)
        {
            try
            {
                var body = doc.MainDocumentPart?.Document?.Body;
                if (body == null) return;

                var tocStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles.Elements<Style>().Where(s => s.StyleId?.Value?.StartsWith("TOC") == true).ToDictionary(s => s.StyleId.Value);
                var tocField = body.Descendants<FieldCode>().FirstOrDefault(f => f.Text?.Contains(" TOC ") == true || f.Text?.Contains("TOC \\") == true);

                if (tocField == null)
                {
                    Debug.WriteLine("Оглавление не найдено");
                    return;
                }

                var tocContainer = tocField.Ancestors<Table>().FirstOrDefault() ?? tocField.Ancestors<Paragraph>().FirstOrDefault()?.Parent;

                if (tocContainer == null)
                {
                    Debug.WriteLine("Контейнер оглавления не найден");
                    return;
                }

                bool hasAnyErrors = false;

                foreach (var paragraph in tocContainer.Descendants<Paragraph>())
                {
                    if (IsEmptyParagraph(paragraph)) continue;

                    bool hasError = false;
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    var justification = paragraph.ParagraphProperties?.Justification;
                    
                    Style paragraphStyle = null;
                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value; // Получаем стиль параграфа
                    if (styleId != null && tocStyles.TryGetValue(styleId, out var style))
                    {
                        paragraphStyle = style;
                    }

                    // 1. Проверка шрифта и размера
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (string.IsNullOrWhiteSpace(run.InnerText) || run.InnerText.Contains("\t") || run.InnerText.Contains("..."))
                            continue;

                        // Проверка шрифта
                        var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? paragraphStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value;
                        if (font == null || !string.Equals(font, gost.TocFontName, StringComparison.OrdinalIgnoreCase))
                        {
                            hasError = true;
                            break;
                        }

                        // Проверка размера шрифта
                        double actualSize = -1;
                        var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value ?? paragraphStyle?.StyleRunProperties?.FontSize?.Val?.Value;

                        if (fontSizeVal != null)
                        {
                            actualSize = double.Parse(fontSizeVal) / 2;
                        }
                        else
                        {
                            actualSize = DefaultTocSize;
                        }

                        if (Math.Abs(actualSize - gost.TocFontSize.Value) > 0.1)
                        {
                            hasError = true;
                            break;
                        }
                    }

                    // 2. Проверка выравнивания
                    var actualAlignment = GetAlignmentString(justification);
                    if (actualAlignment != (gost.TocAlignment ?? DefaultTocAlignment))
                    {
                        hasError = true;
                    }

                    // 3. Проверка отступов
                    // Левый отступ
                    if (gost.TocIndentLeft.HasValue)
                    {
                        double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultTocLeftIndent;
                        if (Math.Abs(actualLeft - gost.TocIndentLeft.Value) > 0.05)
                        {
                            hasError = true;
                        }
                    }

                    // Правый отступ
                    if (gost.TocIndentRight.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTocRightIndent;
                        if (Math.Abs(actualRight - gost.TocIndentRight.Value) > 0.05)
                        {
                            hasError = true;
                        }
                    }

                    // 4. Проверка отступа/выступа первой строки
                    if (gost.TocFirstLineIndent.HasValue && !string.IsNullOrEmpty(gost.TocIndentOrOutdent))
                    {
                        if (gost.TocIndentOrOutdent != "нет")
                        {
                            string currentType = "не определен";
                            double? currentValue = null;

                            if (indent?.Hanging != null)
                            {
                                currentType = "выступ";
                                currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                            }
                            else if (indent?.FirstLine != null)
                            {
                                currentType = "отступ";
                                currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                            }

                            string requiredType = gost.TocIndentOrOutdent.ToLower().Trim();
                            currentType = currentType.ToLower().Trim();

                            // Проверка типа
                            if (currentType != requiredType)
                            {
                                hasError = true;
                                Debug.WriteLine($"Ошибка типа первой строки: текущий '{currentType}', требуется '{requiredType}'");
                            }

                            if (currentValue.HasValue)
                            {
                                if (Math.Abs(currentValue.Value - gost.TocFirstLineIndent.Value) > 0.05)
                                {
                                    hasError = true;
                                    Debug.WriteLine($"Ошибка значения первой строки: текущее '{currentValue:F2} см', требуется '{gost.TocFirstLineIndent.Value:F2} см'");
                                }
                            }
                            else
                            {
                                hasError = true;
                                Debug.WriteLine("Отсутствует отступ/выступ первой строки");
                            }
                        }
                    }

                    // 5. Проверка межстрочного интервала
                    if (gost.TocLineSpacing.HasValue)
                    {
                        double actualSpacing = DefaultTocLineSpacingValue;
                        LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                        if (spacing?.Line != null)
                        {
                            if (spacing.LineRule?.Value == LineSpacingRuleValues.Exact)
                            {
                                actualSpacing = double.Parse(spacing.Line.Value) / 20.0;
                                actualRule = LineSpacingRuleValues.Exact;
                            }
                            else if (spacing.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                            {
                                actualSpacing = double.Parse(spacing.Line.Value) / 20.0;
                                actualRule = LineSpacingRuleValues.AtLeast;
                            }
                            else
                            {
                                actualSpacing = double.Parse(spacing.Line.Value) / 240.0;
                                actualRule = LineSpacingRuleValues.Auto;
                            }
                        }

                        // Проверка типа интервала
                        LineSpacingRuleValues requiredRule = (gost.TocLineSpacingType ?? DefaultTocLineSpacingType) switch
                        {
                            "Минимум" => LineSpacingRuleValues.AtLeast,
                            "Точно" => LineSpacingRuleValues.Exact,
                            _ => LineSpacingRuleValues.Auto
                        };

                        if (actualRule != requiredRule || Math.Abs(actualSpacing - gost.TocLineSpacing.Value) > 0.01)
                        {
                            hasError = true;
                        }
                    }

                    // 6. Проверка интервалов перед/после
                    if (gost.TocLineSpacingBefore.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultTocSpacingBefore;

                        if (Math.Abs(actualBefore - gost.TocLineSpacingBefore.Value) > 0.01)
                        {
                            hasError = true;
                        }
                    }

                    if (gost.TocLineSpacingAfter.HasValue)
                    {
                        double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultTocSpacingAfter;

                        if (Math.Abs(actualAfter - gost.TocLineSpacingAfter.Value) > 0.01)
                        {
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        hasAnyErrors = true;
                        HighlightParagraph(paragraph);
                    }
                }

                if (!hasAnyErrors)
                {
                    Debug.WriteLine("Ошибок форматирования в оглавлении не найдено.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка при подсветке ошибок в оглавлении: {ex.Message}");
            }
        }

        /// <summary>
        /// Выделение в Оглавлении
        /// </summary>
        /// <param name="paragraph"></param>
        private void HighlightTocItem(Paragraph paragraph)
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                run.RunProperties ??= new RunProperties();

                // Удаление старого выделение
                var existingHighlight = run.RunProperties.Elements<Highlight>().FirstOrDefault();
                if (existingHighlight != null)
                {
                    run.RunProperties.RemoveChild(existingHighlight);
                }
                
                run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red }); // Красное выделение фона

            }
        }

        /// <summary>
        /// Метод выделяет ошибки в простом тексте
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <param name="errors"></param>
        private void HighlightTextFormattingErrors(List<Paragraph> paragraphs, WordprocessingDocument doc, Gost gost, List<string> errors)
        {
            const double defaultWordSpacingAfter = 0.35; // см
            const double defaultWordSpacingBefore = 0.0; // см

            string requiredAlignment = gost.TextAlignment;
            string requiredFont = gost.FontName;
            double requiredSize = FontSize;

            var defaultStyle = GetDefaultStyle(doc);
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                // Пропускаем заголовки, списки, таблицы и пустые абзацы
                if (ShouldSkipParagraph(paragraph, headerTexts, gost) || ShouldSkipSpacingCheck(paragraph, headerTexts))
                    continue;

                bool hasError = false;
                var errorDetails = new List<string>();
                var indent = paragraph.ParagraphProperties?.Indentation;
                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;

                // 1. Межстрочный интервал 
                if (gost.LineSpacingValue.HasValue)
                {
                    double? lineVal = spacing?.Line != null ? double.Parse(spacing.Line.Value) / 240.0 : null;

                    if (!lineVal.HasValue || Math.Abs(lineVal.Value - gost.LineSpacingValue.Value) > 0.01)
                    {
                        errorDetails.Add("Неверный межстрочный интервал");
                        hasError = true;
                    }
                }

                // 2. Абзацные отступы 
                if (gost.IndentLeftText.HasValue)
                {
                    double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                    if (Math.Abs(actualLeft - gost.IndentLeftText.Value) > 0.05)
                    {
                        errorDetails.Add($"Левый отступ: {actualLeft:F2} см (требуется {gost.IndentLeftText.Value:F2} см)");
                        hasError = true;
                    }
                }

                if (gost.IndentRightText.HasValue)
                {
                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : 0;
                    if (Math.Abs(actualRight - gost.IndentRightText.Value) > 0.05)
                    {
                        errorDetails.Add($"Правый отступ: {actualRight:F2} см (требуется {gost.IndentRightText.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 2.1 Отступ/выступ первой строки
                if (!string.IsNullOrEmpty(gost.TextIndentOrOutdent) && gost.TextIndentOrOutdent != "нет")
                {
                    string currentType = "не определен";
                    double? currentValue = null;

                    if (indent?.Hanging != null)
                    {
                        currentType = "выступ";
                        currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                    }
                    else if (indent?.FirstLine != null)
                    {
                        currentType = "отступ";
                        currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                    }

                    string expectedType = gost.TextIndentOrOutdent.ToLower();

                    if (currentType != expectedType)
                    {
                        errorDetails.Add($"Тип первой строки: {currentType} (требуется {expectedType})");
                        hasError = true;
                    }

                    if (currentValue.HasValue)
                    {
                        if (Math.Abs(currentValue.Value - (double)gost.FirstLineIndent) > 0.05)
                        {
                            errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gost.FirstLineIndent:F2} см)");
                            hasError = true;
                        }
                    }
                    else
                    {
                        errorDetails.Add($"Отсутствует {expectedType} первой строки");
                        hasError = true;
                    }
                }

                // 3. Выравнивание
                var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification) ?? DefaultTextAlignment;
                if (currentAlignment != requiredAlignment)
                {
                    errorDetails.Add($"Выравнивание: {currentAlignment} (требуется {requiredAlignment})");
                    hasError = true;
                }

                // 4. Интервалы до и после
                if (gost.LineSpacingBefore.HasValue)
                {
                    double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : defaultWordSpacingBefore;
                    if (Math.Abs(actualBefore - gost.LineSpacingBefore.Value) > 0.1)
                    {
                        errorDetails.Add($"Интервал перед абзацем: {actualBefore:F2} см (требуется {gost.LineSpacingBefore.Value:F2} см)");
                        hasError = true;
                    }
                }

                if (gost.LineSpacingAfter.HasValue)
                {
                    double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : defaultWordSpacingAfter;
                    if (Math.Abs(actualAfter - gost.LineSpacingAfter.Value) > 0.1)
                    {
                        errorDetails.Add($"Интервал после абзаца: {actualAfter:F2} см (требуется {gost.LineSpacingAfter.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 5. Шрифт и размер
                foreach (var run in paragraph.Descendants<Run>())
                {
                    if (string.IsNullOrWhiteSpace(run.InnerText))
                        continue;

                    var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? defaultStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value ?? DefaultTextFont;
                    var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value ?? defaultStyle?.StyleRunProperties?.FontSize?.Val?.Value;
                    double fontSize = fontSizeVal != null ? double.Parse(fontSizeVal) / 2 : DefaultTextSize;
                    bool isWrongFont = font != requiredFont;
                    bool isWrongSize = Math.Abs(fontSize - requiredSize) > 0.1;

                    if (isWrongFont || isWrongSize)
                    {
                        run.RunProperties ??= new RunProperties();
                        run.RunProperties.RemoveAllChildren<Highlight>();
                        run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });

                        if (isWrongFont) errorDetails.Add($"Шрифт: {font} (требуется {requiredFont})");
                        if (isWrongSize) errorDetails.Add($"Размер шрифта: {fontSize:F1} pt (требуется {requiredSize:F1} pt)");

                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string shortText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                    errors.Add($"Абзац '{shortText}': {string.Join(", ", errorDetails)}");

                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (!string.IsNullOrWhiteSpace(run.InnerText))
                        {
                            run.RunProperties ??= new RunProperties();
                            run.RunProperties.RemoveAllChildren<Highlight>();
                            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                        }
                    }
                }
            }

            doc.MainDocumentPart.Document.Save();
        }

        /// <summary>
        /// Проверяет является ли параграф заголовком
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool IsHeaderParagraph(Paragraph paragraph, Gost gost)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections))
                return false;

            // Проверка по стилю
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) &&
                (styleId.StartsWith("Heading") || styleId.StartsWith("Заголовок")))
                return true;

            // Проверка по тексту
            var requiredSections = GetRequiredSectionsList(gost);
            var paragraphText = paragraph.InnerText.Trim();

            // Удаляем нумерацию (например "1. Введение" -> "Введение")
            string cleanText = Regex.Replace(paragraphText, @"^\d+[\s\.]*", "").Trim();

            return requiredSections.Any(section =>
                cleanText.Equals(section, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Обрезает текст элемента оглавления до 30 символов
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private string GetShortTocText(Paragraph paragraph)
        {
            string text = paragraph.InnerText.Trim();
            if (text.Length > 30)
            {
                return text.Substring(0, 27) + "...";
            }
            return text;
        }

        /// <summary>
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twips"></param>
        /// <returns></returns>
        private double TwipsToCm(double twips) => twips / 567.0;

        /// <summary>
        /// Вспомогательный метод для получения сокращенного текста
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string GetShortText2(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "[пустой элемент]";

            return text.Length > 30 ? text.Substring(0, 27) + "..." : text;
        }

        /// <summary>
        /// Обрезает текст параграфа до 50 символов с добавлением многоточия
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private string GetShortText(Paragraph paragraph)
        {
            string text = paragraph.InnerText.Trim();
            return text.Length > 50 ? text.Substring(0, 47) + "..." : text;
        }

        /// <summary>
        /// Конвертирует строковое значение в twips в пункты 
        /// </summary>
        /// <param name="twipsValue"></param>
        /// <returns></returns>
        private double ConvertTwipsToPoints(string twipsValue)
        {
            if (string.IsNullOrEmpty(twipsValue))
                return 0;

            double value = double.Parse(twipsValue);
            return value / 567.0; // Стандартное преобразование twips -> pt
        }

        /// <summary>
        /// Проверяет, нужно ли пропускать проверку интервалов для данного параграфа
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headerTexts"></param>
        /// <returns></returns>
        private bool ShouldSkipSpacingCheck(Paragraph paragraph, HashSet<string> headerTexts)
        {
            // Пропуск заголовков
            if (headerTexts.Contains(paragraph.InnerText.Trim()))
                return true;

            // Пропуск пустых параграфов
            if (IsEmptyParagraph(paragraph))
                return true;

            // Пропуск элементов списков
            if (IsListItem(paragraph))
                return true;

            // Пропуск таблиц
            if (paragraph.Ancestors<Table>().Any())
                return true;

            // Пропуск специальных стилей
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(style) && style.Contains("TOC"))
                return true;

            return false;
        }

        /// <summary>
        /// Выделяет красным только указанный параграф 
        /// </summary>
        /// <param name="paragraph"></param>
        private void HighlightOnlyThisParagraph(Paragraph paragraph)
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                run.RunProperties ??= new RunProperties();

                var existingHighlight = run.RunProperties.Elements<Highlight>().FirstOrDefault();
                if (existingHighlight != null) run.RunProperties.RemoveChild(existingHighlight);

                run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
            }
        }

        /// <summary>
        /// Выделяет красным цветом ошибки интервалов в параграфе
        /// </summary>
        /// <param name="paragraph"></param>
        private void HighlightParagraphSpacingErrors(Paragraph paragraph)
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                if (!string.IsNullOrWhiteSpace(run.InnerText))
                {
                    run.RunProperties ??= new RunProperties();
                    run.RunProperties.RemoveAllChildren<Highlight>();
                    run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                }
            }
        }

        /// <summary>
        /// Определяет уровень вложенности списка (1-9)
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private int GetListLevel(Paragraph paragraph, Gost gost)
        {
            var numberingProps = paragraph.ParagraphProperties?.NumberingProperties;

            if (numberingProps?.NumberingLevelReference?.Val?.Value != null)
            {
                return numberingProps.NumberingLevelReference.Val.Value + 1;
            }

            var indent = paragraph.ParagraphProperties?.Indentation;
            if (indent?.Left != null)
            {
                double leftIndent = double.Parse(indent.Left.Value) / 567.0; // в см

                if (gost.ListLevel3Indent.HasValue && leftIndent >= gost.ListLevel3Indent.Value - 0.5)
                    return 3;
                if (gost.ListLevel2Indent.HasValue && leftIndent >= gost.ListLevel2Indent.Value - 0.5)
                    return 2;
            }

            return 1; // По умолчанию 
        }

        /// <summary>
        /// Проверяет является ли параграф нумерованным списком по формату первого символа
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsNumberedList(Paragraph paragraph)
        {
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun == null) return false;

            var text = firstRun.InnerText.Trim();

            return Regex.IsMatch(text, @"^(\d+[\.\)]|[a-z]\)|[A-Z]\.|I+\.|V+\.|X+\.)");// Форматы нумерации: 1., 1), a., a), I., и т.д.
        }

        /// <summary>
        /// Проверяет соответствие формата нумерации требуемому
        /// </summary>
        /// <param name="text"></param>
        /// <param name="requiredFormat"></param>
        /// <returns></returns>
        private bool CheckNumberFormat(string text, string requiredFormat)
        {
            if (requiredFormat.EndsWith(".") && text.EndsWith("."))
                return true;
            if (requiredFormat.EndsWith(")") && text.EndsWith(")"))
                return true;
            return false;
        }

        /// <summary>
        /// Выделяет красным указанный участок текста (Run)
        /// </summary>
        /// <param name="run"></param>
        private void HighlightRun(Run run)
        {
            run.RunProperties ??= new RunProperties();
            run.RunProperties.RemoveAllChildren<Highlight>();
            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
        }

        /// <summary>
        /// Выделяет красным весь параграф и все вложенные Run элементы
        /// </summary>
        /// <param name="paragraph"></param>
        private void HighlightParagraph(Paragraph paragraph)
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                if (!string.IsNullOrWhiteSpace(run.InnerText))
                {
                    HighlightRun(run);
                }
            }
        }

        /// <summary>
        /// Проверяет, является ли параграф пустым
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private bool IsEmptyParagraph(Paragraph p)
        {
            return !p.Descendants<Run>().Any(r => !string.IsNullOrWhiteSpace(r.InnerText));
        }

        /// <summary>
        /// Определяет нужно ли пропускать проверку форматирования для данного параграфа (списки, пустые, заголовки, оглавление)
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headerTexts"></param>
        /// <returns></returns>
        private bool ShouldSkipHighlighting(Paragraph paragraph, HashSet<string> headerTexts)
        {
            //// Пропуск титульного листа
            //if (IsTitleParagraph(paragraph))
            //    return true;

            return IsListItem(paragraph) || IsEmptyParagraph(paragraph) || headerTexts.Contains(paragraph.InnerText.Trim()) ||IsTocParagraph(paragraph);

        }

        /// <summary>
        /// Выводит сообщение об ошибке в интерфейс для заголовков
        /// </summary>
        /// <param name="message"></param>
        private void ShowHeaderError(string message)
        {
            Dispatcher.UIThread.Post(() => {

                ErrorControlHeaderSpacing.Text = message;
                ErrorControlHeaderSpacing.Foreground = Brushes.Red;

            });
        }

        /// <summary>
        /// Выводит сообщение об успехе в интерфейс для заголовков
        /// </summary>
        /// <param name="message"></param>
        private void ShowHeaderSuccess(string message)
        {
            Dispatcher.UIThread.Post(() => {

                ErrorControlHeaderSpacing.Text = message;
                ErrorControlHeaderSpacing.Foreground = Brushes.Green;

            });
        }

        /// <summary>
        /// Обновляет интерфейс с результатами проверки списков 
        /// </summary>
        /// <param name="errors"></param>
        /// <param name="listsValid"></param>
        /// <param name="hasLists"></param>
        private void UpdateBulletedListsUI(List<string> errors, bool listsValid, bool hasLists)
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (errors.Any())
                {
                    ErrorControlBulletedLists.Text = "Проблемы в списках:\n" + string.Join("\n", errors.Distinct());
                    ErrorControlBulletedLists.Foreground = Brushes.Red;
                }
                else if (hasLists)
                {
                    ErrorControlBulletedLists.Text = "Списки соответствуют ГОСТу";
                    ErrorControlBulletedLists.Foreground = Brushes.Green;
                }
                else
                {
                    ErrorControlBulletedLists.Text = "Списки не обнаружены - проверка не требуется";
                    ErrorControlBulletedLists.Foreground = Brushes.Gray;
                }
            });
        }

        /// <summary>
        /// Проверка обязательных разделов у ЗАГОЛОВКОВ
        /// </summary>
        /// <param name="gost"></param>
        /// <param name="body"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool CheckRequiredSections(Gost gost, List<Paragraph> paragraphs, WordprocessingDocument doc)
        {
            if (string.IsNullOrEmpty(gost.RequiredSections))
                return true;

            string requiredFont = gost.HeaderFontName;
            double? requiredSize = gost.HeaderFontSize;

            bool checkFont = !string.IsNullOrEmpty(requiredFont);
            bool checkSize = requiredSize.HasValue;

            var requiredSections = GetRequiredSectionsList(gost);
            bool allSectionsFound = true;
            bool allSectionsValid = true;
            var missingSections = new List<string>();
            var invalidSections = new List<string>();

            // Получение стилей заголовков
            var headerStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles.Elements<Style>().Where(s => s.StyleName?.Val?.Value?.StartsWith("Heading") == true ||
                                                                           s.StyleName?.Val?.Value?.StartsWith("Заголовок") == true).ToDictionary(s => s.StyleId.Value);

            foreach (var section in requiredSections)
            {
                bool sectionFound = false;
                bool sectionValid = true;

                foreach (var paragraph in paragraphs)
                {
                    var text = paragraph.InnerText.Trim();
                    if (text.IndexOf(section, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        sectionFound = true;

                        // Стиль параграфа
                        Style paragraphStyle = null;
                        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

                        if (styleId != null && headerStyles.TryGetValue(styleId, out var style))
                        {
                            paragraphStyle = style;
                        }

                        foreach (var run in paragraph.Elements<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            // Проверка типа шрифта 
                            if (checkFont)
                            {
                                var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? paragraphStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value;

                                if (font != null && !string.Equals(font, requiredFont, StringComparison.OrdinalIgnoreCase))
                                {
                                    invalidSections.Add($"{section} (неверный шрифт: '{font}')");
                                    HighlightRun(run);
                                    sectionValid = false;
                                }
                            }

                            // Проверка размера шрифта
                            if (checkSize)
                            {
                                double? size = null;
                                var fontSize = run.RunProperties?.FontSize?.Val?.Value ?? paragraphStyle?.StyleRunProperties?.FontSize?.Val?.Value;

                                if (fontSize != null)
                                {
                                    size = double.Parse(fontSize) / 2;
                                }

                                if (size.HasValue && Math.Abs(size.Value - requiredSize.Value) > 0.1)
                                {
                                    invalidSections.Add($"{section} (неверный размер: {size:F1} pt)");
                                    HighlightRun(run);
                                    sectionValid = false;
                                }
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
                    ErrorControlSections.Text = $"Ошибки в разделах:\n{string.Join("\n", invalidSections.Take(3))}";
                    if (invalidSections.Count > 3)
                        ErrorControlSections.Text += $"\n...и ещё {invalidSections.Count - 3} ошибок";
                    ErrorControlSections.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlSections.Text = "Все обязательные разделы соответствуют требованиям ГОСТ";
                    ErrorControlSections.Foreground = Brushes.Green;
                }
            });

            return allSectionsFound && allSectionsValid;
        }

        /// <summary>
        /// Проверяет тип шрифта у ПРОСТОГО ТЕКСТА 
        /// </summary>
        /// <param name="requiredFontName"></param>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool CheckFontName(string requiredFontName, List<Paragraph> paragraphs, Gost gost, WordprocessingDocument doc)
        {
            var defaultStyle = GetDefaultStyle(doc);
            string defaultFont = defaultStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value;

            bool isValid = true;
            var errors = new List<string>();
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                if (IsHeaderParagraph(paragraph, gost) || ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {
                    if (ShouldSkipRun(run)) continue;

                    var fontName = run.RunProperties?.RunFonts?.Ascii?.Value;

                    if (string.IsNullOrEmpty(fontName))
                        fontName = defaultFont;

                    if (string.IsNullOrEmpty(fontName))
                        fontName = DefaultTextFont;

                    if (fontName != requiredFontName)
                    {
                        errors.Add($"{run.InnerText.Trim()} (шрифт: '{fontName}') - требуется ('{requiredFontName}')");
                        isValid = false;
                    }
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (!isValid)
                {
                    ErrorControlFont.Text = "Ошибки в шрифте:\n" + string.Join("\n", errors.Take(3));
                    if (errors.Count > 3) ErrorControlFont.Text += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlFont.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlFont.Text = "Шрифт соответствует ГОСТу";
                    ErrorControlFont.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Проверяет размер шрифта у ПРОСТОГО ТЕКСТА 
        /// </summary>
        /// <param name="requiredFontSize"></param>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool CheckFontSize(double requiredFontSize, List<Paragraph> paragraphs, Gost gost, WordprocessingDocument doc)
        {
            var defaultStyle = GetDefaultStyle(doc);
            var defaultSize = defaultStyle?.StyleRunProperties?.FontSize?.Val?.Value;

            bool isValid = true;
            var errors = new List<string>();
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                if (IsHeaderParagraph(paragraph, gost) || ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                foreach (var run in paragraph.Elements<Run>())
                {

                    if (ShouldSkipRun(run)) continue;

                    var fontSize = run.RunProperties?.FontSize?.Val?.Value;

                    double actualSize;
                    if (fontSize == null)
                        fontSize = defaultSize;

                    if (fontSize == null)
                    {
                        actualSize = DefaultTextSize;
                    }
                    else
                    {
                        actualSize = double.Parse(fontSize) / 2;
                    }

                    if (Math.Abs(actualSize - requiredFontSize) > 0.1)
                    {
                        errors.Add($"{run.InnerText.Trim()} (размер: {actualSize:F1} pt) - требуется ({requiredFontSize:F1} pt)");
                        isValid = false;
                    }
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (!isValid)
                {
                    ErrorControlFontSize.Text = "Ошибки в размере шрифта:\n" + string.Join("\n", errors.Take(3));
                    if (errors.Count > 3) ErrorControlFontSize.Text += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlFontSize.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlFontSize.Text = "Размер шрифта соответствует ГОСТу";
                    ErrorControlFontSize.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Проверяет выравнивания ПРОСТОГО ТЕКСТА 
        /// </summary>
        /// <param name="requiredAlignment"></param>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool CheckTextAlignment(string requiredAlignment, List<Paragraph> paragraphs, Gost gost)
        {
            bool isValid = true;
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                if (ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);
                if (currentAlignment != requiredAlignment)
                {
                    Dispatcher.UIThread.Post(() => {

                        ErrorControlViravnivanie.Text = $"{paragraph.InnerText.Trim()} (выравнивание: {currentAlignment}) - требуется ({requiredAlignment})";
                        ErrorControlViravnivanie.Foreground = Brushes.Red;

                    });
                    isValid = false;
                    break;
                }
            }

            if (isValid)
            {
                Dispatcher.UIThread.Post(() => {
                    ErrorControlViravnivanie.Text = "Выравнивание соответствует ГОСТу";
                    ErrorControlViravnivanie.Foreground = Brushes.Green;
                });
            }

            return isValid;
        }

        /// <summary>
        /// Проверяет междустрочный интервал ПРОСТОГО ТЕКСТА
        /// </summary>
        /// <param name="requiredLineSpacing"></param>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool CheckLineSpacing(double requiredLineSpacing, List<Paragraph> paragraphs, Gost gost)
        {
            bool isValid = true;
            var errors = new List<string>();
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                if (ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                // Текст параграфа вывода в ошибке
                string paragraphText = GetShortText(paragraph);

                // Текущие значения 
                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                double actualSpacing;
                string actualSpacingType;

                if (spacing != null && spacing.Line != null)
                {
                    actualSpacing = CalculateActualSpacing(spacing);
                    actualSpacingType = ConvertSpacingRuleToName(spacing.LineRule?.Value);
                }
                else
                {
                    actualSpacing = DefaultTextLineSpacingValue;
                    actualSpacingType = DefaultTextLineSpacingType;
                }

                // Проверка типа интервала
                string requiredSpacingType = gost.LineSpacingType ?? DefaultTextLineSpacingType;
                if (actualSpacingType != requiredSpacingType)
                {
                    errors.Add($"{paragraphText} (тип интервала: '{actualSpacingType}') - требуется ('{requiredSpacingType}')");
                    isValid = false;
                }

                // Проверка значений интервала
                if (Math.Abs(actualSpacing - requiredLineSpacing) > 0.01)
                {
                    errors.Add($"{paragraphText} (интервал: {actualSpacing:F2}) - требуется ({requiredLineSpacing:F2})");
                    isValid = false;
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (!isValid)
                {
                    ErrorControlMnochitel.Text = "Ошибки в межстрочном интервале:\n" + string.Join("\n", errors.Take(3));
                    if (errors.Count > 3)
                        ErrorControlMnochitel.Text += $"\n...и ещё {errors.Count - 3} ошибок";
                    ErrorControlMnochitel.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlMnochitel.Text = "Межстрочный интервал соответствует ГОСТу";
                    ErrorControlMnochitel.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Определяет типп межстрочного интервала
        /// </summary>
        /// <param name="rule"></param>
        /// <returns></returns>
        private string ConvertSpacingRuleToName(LineSpacingRuleValues? rule)
        {
            if (rule == null) return "Не задан";

            if (rule.Value == LineSpacingRuleValues.AtLeast)
            {
                return "Минимум";
            }
            else if (rule.Value == LineSpacingRuleValues.Exact)
            {
                return "Точно";
            }
            else if (rule.Value == LineSpacingRuleValues.Auto)
            {
                return "Множитель";
            }
            else
            {
                return "Неизвестный";
            }
        }

        /// <summary>
        /// Определяет тип межстрочного интервала
        /// </summary>
        /// <param name="spacing"></param>
        /// <returns></returns>
        private double CalculateActualSpacing(SpacingBetweenLines spacing)
        {
            if (spacing.Line == null) return 0;

            if (spacing.LineRule == LineSpacingRuleValues.Exact)
            {
                // Точно
                return double.Parse(spacing.Line.Value) / 20.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                // Минимум
                return double.Parse(spacing.Line.Value) / 20.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                // Множитель
                return double.Parse(spacing.Line.Value) / 240.0;
            }
            else
            {
                // По умолчанию множитель
                return double.Parse(spacing.Line.Value) / 240.0;
            }
        }

        /// <summary>
        /// Проверка интервалов между абзацами (перед и после) для простого текста
        /// </summary>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool CheckParagraphSpacing(List<Paragraph> paragraphs, Gost gost, WordprocessingDocument doc)
        {
            const double defaultWordSpacingAfter = 0.35; // 0.35 см после абзаца в Word
            const double defaultWordSpacingBefore = 0.0; // 0 см перед абзацем в Word

            if (!gost.LineSpacingBefore.HasValue && !gost.LineSpacingAfter.HasValue)
            {
                Dispatcher.UIThread.Post(() => {

                    ErrorControlParagraphSpacing_Unique.Text = "Интервалы между абзацами не требуются";
                    ErrorControlParagraphSpacing_Unique.Foreground = Brushes.Gray;

                });
                return true;
            }

            bool isValid = true;
            var headerTexts = GetHeaderTexts(paragraphs, gost);
            var paragraphErrors = new List<string>();

            foreach (var paragraph in paragraphs)
            {
                if (ShouldSkipParagraph(paragraph, headerTexts, gost)) continue;

                var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                bool hasError = false;
                string errorDetails = "";

                // Проверка интервала "Перед"
                if (gost.LineSpacingBefore.HasValue)
                {
                    double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : defaultWordSpacingBefore;

                    if (Math.Abs(actualBefore - gost.LineSpacingBefore.Value) > 0.01)
                    {
                        errorDetails += $"Интервал перед: {actualBefore:F1} pt (требуется {gost.LineSpacingBefore.Value:F1} pt)";
                        hasError = true;
                    }
                }

                // Проверка интервала "После"
                if (gost.LineSpacingAfter.HasValue)
                {
                    double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : defaultWordSpacingAfter;

                    if (Math.Abs(actualAfter - gost.LineSpacingAfter.Value) > 0.01)
                    {
                        if (!string.IsNullOrEmpty(errorDetails)) errorDetails += ", ";
                        errorDetails += $"Интервал после: {actualAfter:F1} pt (требуется {gost.LineSpacingAfter.Value:F1} pt)";
                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string shortText = GetShortText(paragraph);
                    paragraphErrors.Add($"Текст: '{shortText}' - {errorDetails}");
                    HighlightOnlyThisParagraph(paragraph);
                    isValid = false;
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (paragraphErrors.Any())
                {
                    string error_Message = "Ошибки в интервалах между абзацами:\n" + string.Join("\n", paragraphErrors.Take(3));
                    if (paragraphErrors.Count > 3) error_Message += $"\n...и ещё {paragraphErrors.Count - 3} ошибок";

                    ErrorControlParagraphSpacing_Unique.Text = error_Message;
                    ErrorControlParagraphSpacing_Unique.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlParagraphSpacing_Unique.Text = "Интервалы между абзацами соответствуют ГОСТу";
                    ErrorControlParagraphSpacing_Unique.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Проверка отступов (левого/правого) и первой строки ДЛЯ ПРОСТОГО ТЕКСТА
        /// </summary>
        /// <param name="requiredFirstLineIndent"></param>
        /// <param name="body"></param>
        /// <param name="gost"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        private bool CheckFirstLineIndent(double requiredFirstLineIndent, List<Paragraph> paragraphs, Gost gost)
        {
            if (gost.TextIndentOrOutdent == "нет")
            {
                Dispatcher.UIThread.Post(() => {

                    ErrorControlFirstLineIndent.Text = "Отступ первой строки не требуется";
                    ErrorControlFirstLineIndent.Foreground = Brushes.Gray;

                });
                return true;
            }

            bool isValid = true;
            var errors = new List<string>();
            var headerTexts = GetHeaderTexts(paragraphs, gost);

            foreach (var paragraph in paragraphs)
            {
                if (ShouldSkipParagraph(paragraph, headerTexts, gost))
                    continue;

                var indent = paragraph.ParagraphProperties?.Indentation;
                bool hasError = false;
                var errorDetails = new List<string>();

                // 1. Проверка левого отступа
                if (gost.IndentLeftText.HasValue)
                {
                    double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultTextLeftIndent;

                    if (Math.Abs(actualLeft - gost.IndentLeftText.Value) > 0.05)
                    {
                        errorDetails.Add($"левый отступ: {actualLeft:F2} см (требуется {gost.IndentLeftText.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 2. Проверка правого отступа
                if (gost.IndentRightText.HasValue)
                {
                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTextRightIndent;

                    if (Math.Abs(actualRight - gost.IndentRightText.Value) > 0.05)
                    {
                        errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {gost.IndentRightText.Value:F2} см)");
                        hasError = true;
                    }
                }

                // 3. Проверка отступа/выступа первой строки
                if (gost.TextIndentOrOutdent != "нет")
                {
                    string currentType = "не определен";
                    double? currentValue = null;

                    if (indent?.Hanging != null)
                    {
                        currentType = "выступ";
                        currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                    }
                    else if (indent?.FirstLine != null)
                    {
                        currentType = "отступ";
                        currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                    }

                    if (!string.IsNullOrEmpty(gost.TextIndentOrOutdent))
                    {
                        string requiredType = gost.TextIndentOrOutdent == "Выступ" ? "выступ" : "отступ";

                        if (currentType != requiredType)
                        {
                            errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                            hasError = true;
                        }
                    }

                    if (currentValue.HasValue)
                    {
                        if (Math.Abs(currentValue.Value - requiredFirstLineIndent) > 0.05)
                        {
                            errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {requiredFirstLineIndent:F2} см)");
                            hasError = true;
                        }
                    }
                    else if (!string.IsNullOrEmpty(gost.TextIndentOrOutdent))
                    {
                        errorDetails.Add($"отсутствует {gost.TextIndentOrOutdent} первой строки");
                        hasError = true;
                    }
                }

                if (hasError)
                {
                    string shortText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                    errors.Add($"Абзац '{shortText}': {string.Join(", ", errorDetails)}");
                    isValid = false;

                    // Выделение ошибки
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        if (!string.IsNullOrWhiteSpace(run.InnerText))
                        {
                            run.RunProperties ??= new RunProperties();
                            run.RunProperties.RemoveAllChildren<Highlight>();
                            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                        }
                    }
                }
            }

            Dispatcher.UIThread.Post(() => {
                if (errors.Any())
                {
                    string errorMessage = $"Ошибки в отступах:\n{string.Join("\n", errors.Take(3))}";
                    if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";

                    ErrorControlFirstLineIndent.Text = errorMessage;
                    ErrorControlFirstLineIndent.Foreground = Brushes.Red;
                }
                else
                {
                    ErrorControlFirstLineIndent.Text = "Отступы соответствуют ГОСТу";
                    ErrorControlFirstLineIndent.Foreground = Brushes.Green;
                }
            });

            return isValid;
        }

        /// <summary>
        /// Определяет нужно ли пропускать параграф при проверке основного текста
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headerTexts"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool ShouldSkipParagraph(Paragraph paragraph, HashSet<string> headerTexts, Gost gost)
        {
            // 1. Пропуск заголовков
            if (IsHeaderParagraph(paragraph, gost))
                return true;

            // 2. Пропуск пустых абзацев
            if (IsEmptyParagraph(paragraph))
                return true;

            // 3. Пропуск элементов списка
            if (IsListItem(paragraph))
                return true;

            // 4. Пропуск таблиц
            if (paragraph.Ancestors<Table>().Any())
                return true;

            // 5. Пропуск только реальных абзацев TOC (не соседних)
            if (IsTocParagraph(paragraph))
                return true;

            // 6. Пропуск специальных стилей
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(style) && (style.Contains("Caption") || style.Contains("Footer") || style.Contains("Header") || style.Contains("Title")))
                return true;

            // 7. Пропуск номеров страниц
            if (paragraph.Descendants<SimpleField>().Any(f => f.Instruction?.Value?.Contains("PAGE") == true))
                return true;

            // 8. Пропуск подписей к рисункам/таблицам
            string text = paragraph.InnerText.Trim();
            if (text.StartsWith("Рисунок") || text.StartsWith("Таблица") || text.StartsWith("Рис.") || text.StartsWith("Табл."))
                return true;

            return false;
        }

        /// <summary>
        /// Определяет нужно ли пропускать Run при проверке
        /// </summary>
        private bool ShouldSkipRun(Run run)
        {
            // Пропуск пустых Run элементов
            if (string.IsNullOrWhiteSpace(run.InnerText))
                return true;

            // Пропуск специальных символов
            if (run.Elements<Break>().Any() || run.Elements<TabChar>().Any())
                return true;

            return false;
        }

        /// <summary>
        /// Проверяет формат листа "к примеру А4"
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

            // Конвертация значений в миллиметры
            double widthMm = (pgSz.Width.Value / 1440.0) * 25.4;
            double heightMm = (pgSz.Height.Value / 1440.0) * 25.4;

            // Допустимое отклонение в 1 мм
            bool isCorrectSize = Math.Abs(widthMm - gost.PaperWidthMm.Value) <= 1 && Math.Abs(heightMm - gost.PaperHeightMm.Value) <= 1;

            Dispatcher.UIThread.Post(() => {

                ErrorControlPaperSize.Text = isCorrectSize ? $"Формат бумаги: {gost.PaperSize} ({widthMm:F1}×{heightMm:F1} мм)" :
                $"Требуется {gost.PaperSize} ({gost.PaperWidthMm}×{gost.PaperHeightMm} мм), " + $"текущий: {widthMm:F1}×{heightMm:F1} мм";

                ErrorControlPaperSize.Foreground = isCorrectSize ? Brushes.Green : Brushes.Red;
            });

            return isCorrectSize;
        }

        /// <summary>
        /// Проверяет ориентацию листа "Альбомная или Книжная"
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
                                    $"Ориентация: {(isPortrait ? "Книжная" : "Альбомная")} " + $"( Должна быть {(shouldBePortrait ? "Книжная" : "Альбомная")} )";

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
        private HashSet<string> GetHeaderTexts(List<Paragraph> paragraphs, Gost gost)
        {
            var requiredSections = GetRequiredSectionsList(gost);
            var headerTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var paragraph in paragraphs)
            {
                var text = paragraph.InnerText.Trim();

                // Удаление номеров (например из "1 Введение" получаем "Введение")
                string cleanText = Regex.Replace(text, @"^\d+[\s\.]*", "").Trim();

                foreach (var section in requiredSections)
                {
                    if (cleanText.Equals(section, StringComparison.OrdinalIgnoreCase))
                    {
                        headerTexts.Add(text);
                        break;
                    }
                }
            }
            return headerTexts;
        }

        /// <summary>
        /// Проверяет соответствие шрифтов в стилях документа
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
        /// Вспомогательный метод который преобразует объект выравнивания в строковое представление
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
        /// Вспомогательный метод который получает список обязательных разделов из строки
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
              
                if (element is SimpleField field)
                {
                    var parentRun = field.Parent as Run;
                    if (parentRun != null) runs.Add(parentRun);
                }

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

                foreach (var run in runs)
                {
                    run.RunProperties ??= new RunProperties();

                    // Удаляем старое выделение
                    var existingHighlight = run.RunProperties.Elements<Highlight>().FirstOrDefault();
                    if (existingHighlight != null)
                    {
                        run.RunProperties.RemoveChild(existingHighlight);
                    }

                    // Добавляем красное выделение
                    run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
                }
            }

            // Проверка верхних кононтитулов
            if (wordDoc.MainDocumentPart.HeaderParts != null)
            {
                foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                {
                    foreach (var paragraph in headerPart.Header.Elements<Paragraph>())
                    {
                        var pageField = paragraph.Descendants<SimpleField>().FirstOrDefault(f => f.Instruction?.Value?.Contains("PAGE") == true);

                        if (pageField != null)
                        {
                            var justification = paragraph.ParagraphProperties?.Justification;
                            string alignment = GetAlignmentString(justification);
                            string position = "Top";

                            // Проверка на  соответствие требованиям
                            bool positionMatch = string.IsNullOrEmpty(requiredPosition) || position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase);
                            bool alignmentMatch = string.IsNullOrEmpty(requiredAlignment) || alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase);

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

            // Проверка нижних колонтитулов
            if (wordDoc.MainDocumentPart.FooterParts != null)
            {
                foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                {
                    foreach (var paragraph in footerPart.Footer.Elements<Paragraph>())
                    {
                        var pageField = paragraph.Descendants<SimpleField>().FirstOrDefault(f => f.Instruction?.Value?.Contains("PAGE") == true);

                        if (pageField != null)
                        {
                            var justification = paragraph.ParagraphProperties?.Justification;
                            string alignment = GetAlignmentString(justification);
                            string position = "Bottom";

                            // Проверка на соответствие требованиям
                            bool positionMatch = string.IsNullOrEmpty(requiredPosition) || position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase);
                            bool alignmentMatch = string.IsNullOrEmpty(requiredAlignment) || alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase);

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
                    string message = string.IsNullOrEmpty(requiredAlignment) && string.IsNullOrEmpty(requiredPosition) ? "Нумерация страниц присутствует" : 
                                                                            $"Нумерация соответствует ({actualCorrectPosition}, {actualCorrectAlignment})";

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
        /// Проверяет, является ли параграф элементом списка
        /// </summary>
        private bool IsListItem(Paragraph paragraph)
        {
            // 1. Проверка нумерации
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // 2. Проверка стиля списка
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) && (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering")))
                return true;

            // 3. Проверка по форматированию
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                var text = firstRun.InnerText.Trim();

                // Маркированные списки
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—"))
                    return true;

                // Нумерованные списки
                if (Regex.IsMatch(text, @"^\d+[\.\)]") || Regex.IsMatch(text, @"^[a-z]\)"))
                    return true;
            }

            return false;
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
        /// Получает стиль из документа
        /// </summary>
        private Style GetDefaultStyle(WordprocessingDocument doc)
        {
            var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
            if (stylesPart == null) return null;

            return stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.Type == StyleValues.Paragraph && (s.Default?.Value ?? false));
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

            // Основное окно
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

            // Кнопки
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

            // Содержимое окна
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

            // Обработчики кнопок
            yesButton.Click += (s, e) => { result = true; dialog.Close(); };
            noButton.Click += (s, e) => { result = false; dialog.Close(); };

            await dialog.ShowDialog(this);
            return result;
        }

    }
}