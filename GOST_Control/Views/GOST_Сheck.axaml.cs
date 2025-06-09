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
using Xceed.Words.NET;
using Avalonia.Layout;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using Avalonia.Styling;
using Styles = DocumentFormat.OpenXml.Wordprocessing.Styles;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace GOST_Control
{
    /// <summary>
    /// Класс провероки документа на соответствие ГОСТу
    /// </summary>
    public partial class GOST_Сheck : Window
    {
        private readonly string _filePath; // Путь к файлу документа, который будет проверяться на соответствие ГОСТу
        private JsonGostService _gostService; // Сервис для работы с данными ГОСТов из JSON-файла
        private readonly Task _initializationTask; // Задача инициализации сервиса ГОСТов, запускаемая при создании экземпляра класса

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
                    return new JsonGostService("GOST_Control.Resources.gosts.json", "gosts_modified.json");
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
        /// Проверяет, является ли параграф частью титульного листа документа.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsTitleParagraph(Paragraph paragraph)
        {
            try
            {
                if (paragraph?.InnerText == null)
                    return false;

                // Безопасное удаление переносов
                string text = Regex.Replace(paragraph.InnerText, @"[\r\n]+", " ").Trim();

                return !string.IsNullOrEmpty(text) && Regex.IsMatch(text, @"(^|\s)20\d{2}(\s*г\.)?($|\s)");
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Отделяет титульный лист
        /// </summary>
        /// <param name="body"></param>
        /// <param name="titlePageParagraphs"></param>
        /// <param name="bodyParagraphsAfterTitle"></param>
        /// <param name="allParagraphs"></param>
        private void ExtractTitleAndBody(Body body, out List<Paragraph> titlePageParagraphs, out List<Paragraph> bodyParagraphsAfterTitle, out List<Paragraph> allParagraphs)
        {
            titlePageParagraphs = new List<Paragraph>();
            bodyParagraphsAfterTitle = new List<Paragraph>();
            allParagraphs = new List<Paragraph>();

            try
            {
                // Получение всех параграфов
                allParagraphs = body?.Elements<Paragraph>()?.ToList() ?? new List<Paragraph>();
            }
            catch
            {
                allParagraphs = new List<Paragraph>();
            }

            bool isTitlePage = true;

            foreach (var paragraph in allParagraphs)
            {
                try
                {
                    if (isTitlePage)
                    {
                        titlePageParagraphs.Add(paragraph);

                        // 1. Проверка разрыва страницы
                        bool hasPageBreak = false;
                        try
                        {
                            hasPageBreak = paragraph.Descendants<Break>()?.Any(b => b.Type == BreakValues.Page) ?? false;
                        }
                        catch { }

                        // 2. Проверка года
                        bool hasYear = IsTitleParagraph(paragraph);

                        // Критерии завершения титульника:
                        if (hasPageBreak || hasYear)
                        {
                            isTitlePage = false;
                        }
                    }
                    else
                    {
                        bodyParagraphsAfterTitle.Add(paragraph);
                    }
                }
                catch
                {
                    // Если ошибка в параграфе - продолжаем обработку
                    bodyParagraphsAfterTitle.Add(paragraph);
                    isTitlePage = false;
                }
            }
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
                    bool imagesValid = true;
                    bool tablesValid = true;
                    bool additionalHeadersValid = true;

                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Список хранения ошибок
                    var errors = new List<string>();
                    var allErrors = new List<TextErrorInfo>();


                    var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
                    // Принудительная перезагрузка
                    stylesPart.Styles.Reload();

                    // ======================= ТИТУЛЬНИК =======================
                    List<Paragraph> allParagraphs;
                    ExtractTitleAndBody(body, out var titlePageParagraphs, out var bodyParagraphsAfterTitle, out allParagraphs);

                    var checkTasks = new List<Task>();

                    // ======================= ПРОСТОЙ ТЕКСТ =======================
                    if (!string.IsNullOrEmpty(gost.FontName) || gost.FontSize.HasValue)
                    {
                        var headerTexts = GetHeaderTexts(bodyParagraphsAfterTitle, gost);
                        var checkingTextPlain = new CheckingPlainText(wordDoc, (p) => ShouldSkipParagraph(p, headerTexts, gost));

                        // === Проверка шрифта ===
                        if (!string.IsNullOrEmpty(gost.FontName))
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, fontErrors) = await checkingTextPlain.CheckFontNameAsync(gost.FontName, bodyParagraphsAfterTitle, wordDoc, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlFont.Text = text;
                                        ErrorControlFont.Foreground = brush;
                                    });
                                });

                                fontNameValid = isValid;
                                if (!isValid)
                                    errors.AddRange(fontErrors.Select(e => $"Шрифт: {e}"));
                                    allErrors.AddRange(fontErrors);
                            }));
                        }

                        // === Проверка размера шрифта ===
                        if (gost.FontSize.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, fontSizeErrors) = await checkingTextPlain.CheckFontSizeAsync(gost.FontSize.Value, bodyParagraphsAfterTitle, wordDoc, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlFontSize.Text = text;
                                        ErrorControlFontSize.Foreground = brush;
                                    });
                                });

                                fontSizeValid = isValid;
                                if (!isValid)
                                    errors.AddRange(fontSizeErrors.Select(e => $"Размер шрифта: {e}"));
                                    allErrors.AddRange(fontSizeErrors);
                            }));
                        }

                        // === Проверка выравнивания шрифта ===
                        if (!string.IsNullOrEmpty(gost.TextAlignment))
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, textAlignmentErrors) = await checkingTextPlain.CheckTextAlignmentAsync(gost.TextAlignment, bodyParagraphsAfterTitle, wordDoc, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlViravnivanie.Text = text;
                                        ErrorControlViravnivanie.Foreground = brush;
                                    });
                                });

                                textAlignmentValid = isValid;
                                if (!isValid)
                                    errors.AddRange(textAlignmentErrors.Select(e => $"Выравнивание текста: {e}"));
                                    allErrors.AddRange(textAlignmentErrors);
                            }));
                        }

                        // === Проверка межстрочного интервала шрифта ===
                        if (gost.LineSpacingValue.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, lineSpacingErrors) = await checkingTextPlain.CheckLineSpacingAsync(gost.LineSpacingValue.Value, gost.LineSpacingType, bodyParagraphsAfterTitle,
                                    wordDoc, (text, brush) =>
                                    {
                                        Dispatcher.UIThread.Post(() =>
                                        {
                                            ErrorControlMnochitel.Text = text;
                                            ErrorControlMnochitel.Foreground = brush;
                                        });
                                    });

                                lineSpacingValid = isValid;
                                if (!isValid)
                                    errors.AddRange(lineSpacingErrors.Select(e => $"Межстрочный интервал: {e}"));
                                    allErrors.AddRange(lineSpacingErrors);
                            }));
                        }

                        // === Проверка интервалов между абзацами ===
                        if (gost.LineSpacingBefore.HasValue || gost.LineSpacingAfter.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, paragraphSpacingErrors) = await checkingTextPlain.CheckParagraphSpacingAsync(gost.LineSpacingBefore.HasValue, gost.LineSpacingAfter.HasValue,
                                    bodyParagraphsAfterTitle, gost, wordDoc, (text, brush) =>
                                    {
                                        Dispatcher.UIThread.Post(() =>
                                        {
                                            ErrorControlParagraphSpacing_Unique.Text = text;
                                            ErrorControlParagraphSpacing_Unique.Foreground = brush;
                                        });
                                    });

                                paragraphSpacingValid = isValid;
                                if (!isValid)
                                    errors.AddRange(paragraphSpacingErrors.Select(e => $"Интервалы между абзацами: {e}"));
                                    allErrors.AddRange(paragraphSpacingErrors);
                            }));
                        }

                        // === Проверка отступов ===
                        if (gost.FirstLineIndent.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, firstLineIndentErrors) = await checkingTextPlain.CheckFirstLineIndentAsync(gost.FirstLineIndent.Value, bodyParagraphsAfterTitle, wordDoc, gost, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlFirstLineIndent.Text = text;
                                        ErrorControlFirstLineIndent.Foreground = brush;
                                    });
                                });

                                firstLineIndentValid = isValid;
                                if (!isValid)
                                    errors.AddRange(firstLineIndentErrors.Select(e => $"Отступы: {e}"));
                                    allErrors.AddRange(firstLineIndentErrors);
                            }));
                        }
                    }

                    // ======================= НАСТРОЙКА ДОКУМЕНТА =======================
                    if (!string.IsNullOrEmpty(gost.PageOrientation) || gost.PaperWidthMm.HasValue || gost.PaperHeightMm.HasValue || gost.PageNumbering.HasValue ||
                                                  gost.MarginTop.HasValue || gost.MarginBottom.HasValue || gost.MarginLeft.HasValue || gost.MarginRight.HasValue)
                    {
                        var docChecker = new CheckingSettingDoc(wordDoc, gost);

                        // === Проверка нумерации страниц ===
                        if (gost.PageNumbering.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var result = await docChecker.CheckPageNumberingAsync(wordDoc, gost.PageNumbering.Value, gost.PageNumberingAlignment, gost.PageNumberingPosition,
                                    (text, brush) =>
                                    {
                                        Dispatcher.UIThread.Post(() =>
                                        {
                                            ErrorControlNumberPage.Text = text;
                                            ErrorControlNumberPage.Foreground = brush;
                                        });
                                    });

                                pageNumberingValid = result.IsValid;
                                if (!result.IsValid)
                                    errors.AddRange(result.Errors.Select(e => $"Нумерация страниц: {e}"));
                            }));
                        }
                        else
                        {
                            Dispatcher.UIThread.Post(() =>
                            {
                                ErrorControlNumberPage.Text = "Нумерация страниц не требуется.";
                                ErrorControlNumberPage.Foreground = Brushes.Gray;
                            });
                        }

                        // === Проверка формата бумаги ===
                        if (gost.PaperWidthMm.HasValue && gost.PaperHeightMm.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, paperSizeErrors) = await docChecker.CheckPaperSizeAsync(wordDoc, gost, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlPaperSize.Text = text;
                                        ErrorControlPaperSize.Foreground = brush;
                                    });
                                });

                                paperSizeValid = isValid;
                                if (!isValid)
                                    errors.AddRange(paperSizeErrors.Select(e => $"Размер бумаги: {e}"));
                            }));
                        }

                        // === Проверка ориентации страницы ===
                        if (!string.IsNullOrEmpty(gost.PageOrientation))
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, orientationErrors) = await docChecker.CheckPageOrientationAsync(wordDoc, gost, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlOrientation.Text = text;
                                        ErrorControlOrientation.Foreground = brush;
                                    });
                                });

                                orientationValid = isValid;
                                if (!isValid)
                                    errors.AddRange(orientationErrors.Select(e => $"Ориентация страницы: {e}"));
                            }));
                        }

                        // === Проверка полей документа ===
                        if (gost.MarginTop.HasValue || gost.MarginBottom.HasValue || gost.MarginLeft.HasValue || gost.MarginRight.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (marginsValid, marginErrors) = await docChecker.CheckMarginsAsync(gost.MarginTop, gost.MarginBottom, gost.MarginLeft, gost.MarginRight, body, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlMargins.Text = text;
                                        ErrorControlMargins.Foreground = brush;
                                    });
                                });

                                if (!marginsValid)
                                    errors.AddRange(marginErrors.Select(e => $"Поля документа: {e}"));
                            }));
                        }
                    }

                    // ======================= ЗАГОЛОВКИ =======================
                    if (!string.IsNullOrEmpty(gost.RequiredSections))
                    {
                        var headerTexts = GetHeaderTexts(bodyParagraphsAfterTitle, gost);

                        // Передаем делегаты в конструктор CheckingeContents
                        var checkingeContents = new CheckingeContents(wordDoc, gost, (run) => ShouldSkipRun(run), (paragraph, gostParam) => IsAdditionalHeader(paragraph, gost));

                        // Проверка обязательных разделов (Введение, Заключение) ЗАГОЛОВКИ
                        if (!string.IsNullOrEmpty(gost.RequiredSections))
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, sectionErrors) = await checkingeContents.CheckRequiredSectionsAsync(gost, bodyParagraphsAfterTitle, wordDoc, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlSections.Text = text;
                                        ErrorControlSections.Foreground = brush;
                                    });
                                });

                                sectionsValid = isValid;
                                if (!isValid)
                                    errors.AddRange(sectionErrors.Select(e => $"Разделы: {e.ErrorMessage}")); 
                                allErrors.AddRange(sectionErrors);
                            }));
                        }

                        // Проверка интервалов для заголовков
                        if (gost.HeaderLineSpacingBefore.HasValue || gost.HeaderLineSpacingAfter.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, spacingErrors) = await checkingeContents.CheckHeaderParagraphSpacingAsync(bodyParagraphsAfterTitle, wordDoc, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlHeaderSpacing.Text = text;
                                        ErrorControlHeaderSpacing.Foreground = brush;
                                    });
                                });

                                headerSpacingValid = isValid;
                                if (!isValid)
                                    errors.AddRange(spacingErrors.Select(e => $"Интервалы заголовков: {e.ErrorMessage}")); // Изменено
                                allErrors.AddRange(spacingErrors);
                            }));
                        }

                        // Проверка отступов для заголовков
                        if (gost.HeaderIndentLeft.HasValue || gost.HeaderIndentRight.HasValue || gost.HeaderFirstLineIndent.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, indentErrors) = await checkingeContents.CheckHeaderIndentsAsync(bodyParagraphsAfterTitle, wordDoc, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlHeaderIndents.Text = text;
                                        ErrorControlHeaderIndents.Foreground = brush;
                                    });
                                });

                                headerIndentsValid = isValid;
                                if (!isValid)
                                    errors.AddRange(indentErrors.Select(e => $"Отступы заголовков: {e.ErrorMessage}")); // Изменено
                                allErrors.AddRange(indentErrors);
                            }));
                        }

                        // = ДОП.ЗАГОЛОВКИ =
                        if (!string.IsNullOrEmpty(gost.AdditionalHeaderFontName) || gost.AdditionalHeaderFontSize.HasValue)
                        {
                            // Проверка дополнительных заголовков
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, additionalHeaderErrors) = await checkingeContents.CheckAdditionalHeadersAsync(wordDoc, bodyParagraphsAfterTitle, gost, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlAdditionalHeaders.Text = text;
                                        ErrorControlAdditionalHeaders.Foreground = brush;
                                    });
                                });

                                additionalHeadersValid = isValid;
                                if (!isValid)
                                {
                                    errors.AddRange(additionalHeaderErrors.Select(e => $"Доп. заголовки: {e.ErrorMessage}")); // Изменено
                                    allErrors.AddRange(additionalHeaderErrors);
                                }
                            }));
                        }
                    }

                    // ======================= ТАБЛИЦЫ =======================
                    if (!string.IsNullOrEmpty(gost.TableCaptionFontName) || gost.TableCaptionFontSize.HasValue || !string.IsNullOrEmpty(gost.TableAlignment) || gost.TableFontSize.HasValue)
                    {
                        var checkingTables = new CheckingTableDoc(wordDoc, gost, ShouldSkipRun);

                        checkTasks.Add(Task.Run(async () =>
                        {
                            var (isValid, tableErrors) = await checkingTables.CheckTablesAsync((text, brush) =>
                            {
                                Dispatcher.UIThread.Post(() =>
                                {
                                    ErrorControlTables.Text = text;
                                    ErrorControlTables.Foreground = brush;
                                });
                            });

                            tablesValid = isValid;
                            if (!isValid)
                            {
                                errors.AddRange(tableErrors.Select(e => $"Таблицы: {e.ErrorMessage}"));
                                allErrors.AddRange(tableErrors);
                            }
                        }));
                    }

                    // ======================= Картинки =======================
                    if (!string.IsNullOrEmpty(gost.ImageCaptionFontName) || gost.ImageCaptionFontSize.HasValue || !string.IsNullOrEmpty(gost.ImageCaptionAlignment) || gost.ImageCaptionFirstLineIndent.HasValue ||
                        !string.IsNullOrEmpty(gost.ImageCaptionLineSpacingType))
                    {
                        var checkingImages = new CheckingImageDoc(wordDoc, gost, ShouldSkipRun);

                        checkTasks.Add(Task.Run(async () =>
                        {
                            var (isValid, imageErrors) = await checkingImages.CheckImagesAsync(bodyParagraphsAfterTitle, (text, brush) =>
                            {
                                Dispatcher.UIThread.Post(() =>
                                {
                                    ErrorControlImages.Text = text;
                                    ErrorControlImages.Foreground = brush;
                                });
                            });

                            imagesValid = isValid;
                            if (!isValid)
                                errors.AddRange(imageErrors.Select(e => $"Изображения: {e}"));
                                allErrors.AddRange(imageErrors);
                        }));
                    }

                    // ======================= ОГЛАВЛЕНИЯ =======================
                    if (!string.IsNullOrEmpty(gost.TocFontName))
                    {
                        var checkOglavleniya = new CheckOglavleniya(wordDoc, gost, (paragraph) => IsTocParagraph(paragraph), (paragraph) => IsEmptyParagraph(paragraph));
                        checkTasks.Add(Task.Run(async () =>
                        {
                            var (isValid, tocErrors) = await checkOglavleniya.CheckTocFormattingAsync((text, brush) =>
                            {
                                Dispatcher.UIThread.Post(() =>
                                {
                                    Error_ControlToc_Spacing.Text = text;
                                    Error_ControlToc_Spacing.Foreground = brush;
                                });
                            });

                            tocValid = isValid;
                            if (!isValid)
                            {
                                errors.AddRange(tocErrors.Select(e => $"Оглавление: {e.ErrorMessage}"));
                                allErrors.AddRange(tocErrors);
                            }
                        }));
                    }

                    // ======================= СПИСКИ =======================
                    bool hasLists = body.Descendants<Paragraph>().Any(IsListItem);
                    if (hasLists)
                    {
                        var checkingLists = new CheckListDocText(gost, IsAdditionalHeader);

                        // Проверка базовых параметров списков
                        checkTasks.Add(Task.Run(async () =>
                        {
                            var (isValid, listErrors) = await checkingLists.CheckBulletedListsAsync(wordDoc, bodyParagraphsAfterTitle, gost, (text, brush) =>
                            {
                                Dispatcher.UIThread.Post(() =>
                                {
                                    ErrorControlBulletedLists.Text = text;
                                    ErrorControlBulletedLists.Foreground = brush;
                                });
                            });

                            bulletedListsValid = isValid;
                            if (!isValid)
                            {
                                errors.AddRange(listErrors.Select(e => $"Списки: {e.ErrorMessage}"));
                                allErrors.AddRange(listErrors);
                            }
                        }));

                        // Проверка интервалов списков
                        if (gost.BulletLineSpacingBefore.HasValue || gost.BulletLineSpacingAfter.HasValue || gost.BulletLineSpacingValue.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, spacingErrors) = await checkingLists.CheckListParagraphSpacingAsync(wordDoc, bodyParagraphsAfterTitle, gost, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlListSpacing.Text = text;
                                        ErrorControlListSpacing.Foreground = brush;
                                    });
                                });

                                listSpacingValid = isValid;
                                if (!isValid)
                                {
                                    errors.AddRange(spacingErrors.Select(e => $"Интервалы списков: {e.ErrorMessage}"));
                                    allErrors.AddRange(spacingErrors);
                                }
                            }));
                        }

                        // Проверка отступов списков
                        if (gost.ListLevel1BulletIndentLeft.HasValue || gost.ListLevel1BulletIndentRight.HasValue || gost.ListLevel1Indent.HasValue || gost.ListLevel2Indent.HasValue || 
                            gost.ListLevel3Indent.HasValue)
                        {
                            checkTasks.Add(Task.Run(async () =>
                            {
                                var (isValid, indentErrors) = await checkingLists.CheckListIndentsAsync(bodyParagraphsAfterTitle, gost, (text, brush) =>
                                {
                                    Dispatcher.UIThread.Post(() =>
                                    {
                                        ErrorControlListIndents.Text = text;
                                        ErrorControlListIndents.Foreground = brush;
                                    });
                                });

                                listHangingValid = isValid;
                                if (!isValid)
                                {
                                    errors.AddRange(indentErrors.Select(e => $"Отступы списков: {e.ErrorMessage}"));
                                    allErrors.AddRange(indentErrors);
                                }
                            }));
                        }
                    }
                    else
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            ErrorControlBulletedLists.Text = "Списки не обнаружены - проверка не требуется";
                            ErrorControlBulletedLists.Foreground = Brushes.Gray;
                        });
                    }

                    // ======================= ПРОВЕРКА НЕОФОРМЛЕННЫХ ГИПЕРССЫЛОК =======================
                    var (isValid, linkErrors) = await CheckPlainTextLinksAsync(wordDoc);
                    plainTextLinksValid = isValid;
                    if (!isValid)
                    {
                        errors.AddRange(linkErrors.Select(e => e.ErrorMessage)); 
                        allErrors.AddRange(linkErrors);
                    }

                    // Ожидаем завершения всех проверок
                    await Task.WhenAll(checkTasks);

                    // Общий результат проверки
                    if (fontNameValid && fontSizeValid && marginsValid && lineSpacingValid && firstLineIndentValid && textAlignmentValid && pageNumberingValid && 
                        sectionsValid && paperSizeValid && orientationValid && tocValid && bulletedListsValid && textIndentsValid && paragraphSpacingValid && 
                        headerSpacingValid && tocSpacingValid && listSpacingValid && listHangingValid && headerIndentsValid && tocIndentsValid && 
                        plainTextLinksValid && imagesValid && tablesValid && additionalHeadersValid)
                    {
                        GostControl.Text = "Документ соответствует ГОСТу.";
                        GostControl.Foreground = Brushes.Green;
                    }
                    else
                    {
                        GostControl.Text = "Документ не соответствует ГОСТу:";
                        GostControl.Foreground = Brushes.Red;

                        // Создаем документ с ошибками
                        await CreateErrorReportDocument(wordDoc, gost, allErrors, filePath, titlePageParagraphs, bodyParagraphsAfterTitle);
                    }
                }
            }
            catch (Exception ex)
            {
                GostControl.Text = $"Ошибка при открытии документа! Закройте документ!";
                GostControl.Foreground = Brushes.Red;
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
        private async Task CreateErrorReportDocument(WordprocessingDocument originalDoc, Gost gost, List<TextErrorInfo> errors, string originalFilePath, List<Paragraph> oldTitlePageParagraphs, List<Paragraph> oldBodyParagraphsAfterTitle)
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
                var mainPart = errorDoc.MainDocumentPart;
                var body = mainPart.Document.Body;

                // Собираем все параграфы и runs из документа
                var allParagraphs = body.Descendants<Paragraph>().ToList();
                var allRuns = body.Descendants<Run>().ToList();

                // Выделения ошибок
                foreach (var error in errors)
                {
                    if (error != null && error.ProblemRun != null)
                    {
                        var originalRuns = originalDoc.MainDocumentPart.Document.Body.Descendants<Run>().ToList();
                        int runIndex = originalRuns.IndexOf(error.ProblemRun);

                        if (runIndex >= 0 && runIndex < allRuns.Count)
                        {
                            var runToMark = allRuns[runIndex];
                            MarkRunWithBackgroundHighlight(runToMark);
                        }
                    }
                    else if (error != null && error.ProblemParagraph != null)
                    {
                        var originalParagraphs = originalDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();
                        int paraIndex = originalParagraphs.IndexOf(error.ProblemParagraph);

                        if (paraIndex >= 0 && paraIndex < allParagraphs.Count)
                        {
                            var paragraphToMark = allParagraphs[paraIndex];
                            MarkParagraphWithBackgroundHighlight(paragraphToMark);
                        }
                    }
                }

                mainPart.Document.Save();
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
        /// Применяет красное фоновое выделение к указанному текстовому фрагменту (Run)
        /// </summary>
        /// <param name="run"></param>
        private void MarkRunWithBackgroundHighlight(Run run)
        {
            run.RunProperties ??= new RunProperties();

            // Удаляем все возможные предыдущие форматирования
            run.RunProperties.RemoveAllChildren<Highlight>();
            run.RunProperties.RemoveAllChildren<Color>();
            run.RunProperties.RemoveAllChildren<Bold>();

            // Добавляем красный фон
            run.RunProperties.Append(new Highlight { Val = HighlightColorValues.Red });
        }

        /// <summary>
        /// Применяет красное фоновое выделение ко всем текстовым фрагментам в абзаце
        /// </summary>
        /// <param name="paragraph"></param>
        private void MarkParagraphWithBackgroundHighlight(Paragraph paragraph)
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                MarkRunWithBackgroundHighlight(run);
            }

            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.RemoveAllChildren<ParagraphMarkRunProperties>();
            paragraph.ParagraphProperties.Append(new ParagraphMarkRunProperties( new Highlight { Val = HighlightColorValues.Red }
            ));
        }

        /// <summary>
        /// Метод для определения, является ли параграф дополнительным заголовком
        /// </summary>
        /// <param name="paragraph">Проверяемый параграф</param>
        /// <param name="gost">Настройки ГОСТа</param>
        /// <returns>True, если это дополнительный заголовок</returns>
        private bool IsAdditionalHeader(Paragraph paragraph, Gost gost)
        {
            // 1. Сначала проверяем по стилю (если стиль соответствует шаблону заголовков)
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) && (styleId.StartsWith("Heading") || styleId.StartsWith("Заголовок") ||  styleId.StartsWith("TOC") || styleId.Contains("Subtitle")))
            {
                return true;
            }

            var text = paragraph.InnerText?.Trim();
            if (string.IsNullOrWhiteSpace(text))
                return false;

            // 2. Проверяем по содержанию текста:

            // Шаблон для обычных нумерованных заголовков (например: "1.1 Общие положения")
            bool isNumberedHeader = Regex.IsMatch(text, @"^\d+(\.\d+)*[\s\t]+[А-Яа-яA-Za-z]", RegexOptions.IgnoreCase);

            // Шаблон для глав (например: "Глава 2")
            bool isChapterHeader = Regex.IsMatch(text, @"^Глава\s+\d+", RegexOptions.IgnoreCase);

            return isNumberedHeader || isChapterHeader;
        }

        /// <summary>
        /// Ищет Гиперссылки
        /// </summary>
        /// <param name="doc"></param>
        private async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckPlainTextLinksAsync(WordprocessingDocument doc)
        {
            return await Task.Run(() =>
            {
                var linkErrors = new List<TextErrorInfo>();
                var regex = new Regex(@"https?://[^\s]+", RegexOptions.Compiled);
                var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();
                bool isValid = true;
                bool instructionShown = false;

                foreach (var paragraph in paragraphs)
                {
                    var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
                    var matches = regex.Matches(text);

                    foreach (Match match in matches)
                    {
                        bool isLinked = paragraph.Descendants<Hyperlink>().Any(h =>
                        {
                            var hyperlinkText = string.Concat(h.Descendants<Text>().Select(t => t.Text));
                            return hyperlinkText.Contains(match.Value);
                        });

                        if (!isLinked)
                        {
                            string url = match.Value.Length > 50 ? match.Value.Substring(0, 47) + "..." : match.Value;
                            linkErrors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"URL без гиперссылки: '{url}'",
                                ProblemRun = null, // Можно определить конкретный Run при необходимости
                                ProblemParagraph = paragraph
                            });
                            isValid = false;
                        }
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (!isValid)
                    {
                        string errorMessage = "Ошибки в гиперссылках:\n" +
                            string.Join("\n", linkErrors.Select(e => e.ErrorMessage).Take(5));

                        if (linkErrors.Count > 5)
                            errorMessage += $"\n...и ещё {linkErrors.Count - 5} ошибок";

                        if (!instructionShown)
                        {
                            errorMessage += "\n\nКак исправить:\n1. Выделите URL\n" +
                                          "2. Нажмите Ctrl+K\n" +
                                          "3. Вставьте URL в поле адреса";
                        }

                        ErrorControlLinks.Text = errorMessage;
                        ErrorControlLinks.Foreground = Brushes.Red;
                    }
                    else
                    {
                        ErrorControlLinks.Text = "✓ Все гиперссылки оформлены корректно";
                        ErrorControlLinks.Foreground = Brushes.Green;
                    }
                });

                return (isValid, linkErrors);
            });
        }

        /// <summary>
        /// Метод проверки, является ли параграф частью оглавления
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsTocParagraph(Paragraph paragraph)
        {
            if (paragraph == null) return false;

            // Проверка стилей оглавления
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) && (styleId.StartsWith("TOC") || styleId.Contains("Contents") || styleId.Contains("оглавление")))
            {
                return true;
            }

            // 1. Проверка поля TOC (автоматическое оглавление)
            if (paragraph.Descendants<FieldCode>().Any(f =>
                Regex.IsMatch(f.Text, @"\bTOC\b", RegexOptions.IgnoreCase)))
            {
                return true;
            }

            // 2. Проверка стиля, чтобы учесть не только точные совпадения
            if (!string.IsNullOrEmpty(styleId) && (styleId.Contains("toc") || styleId.Contains("contents") || styleId.Contains("table")))
            {
                return true;
            }

            // 3. Проверка на характерные признаки для оглавлений, но исключаем нежелательные строки
            string text = paragraph.InnerText;
            if (text.Contains(".........") || text.Contains("\t") || Regex.IsMatch(text, @"\.{3,}\s*\d+$"))
            {
                return true;
            }

            // 4. Проверка на родительскую таблицу для оглавлений
            var parentTable = paragraph.Ancestors<Table>().FirstOrDefault();
            if (parentTable != null)
            {
                return parentTable.Descendants<FieldCode>().Any(f =>
                    Regex.IsMatch(f.Text, @"\bTOC\b", RegexOptions.IgnoreCase));
            }

            return false;
        }

        /// <summary>
        /// Проверяет является ли параграф заголовком
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool IsHeaderParagraph(Paragraph paragraph, Gost gost)
        {
            var text = paragraph.InnerText?.Trim();
            if (string.IsNullOrWhiteSpace(text))
                return false;

            // 1. Проверка обязательных разделов (если они указаны в ГОСТе)
            if (!string.IsNullOrEmpty(gost.RequiredSections))
            {
                // Проверка по стилю
                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (!string.IsNullOrEmpty(styleId) &&
                    (styleId.StartsWith("Heading") || styleId.StartsWith("Заголовок")))
                    return true;

                // Проверка по тексту
                var requiredSections = GetRequiredSectionsList(gost);

                // Удаляем нумерацию (например "1. Введение" -> "Введение")
                string cleanText = Regex.Replace(text, @"^\d+[\s\.]*", "").Trim();

                if (requiredSections.Any(section => cleanText.Equals(section, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            // 2. Проверка на приложение 
            bool isAppendix = Regex.IsMatch(text, @"^ПРИЛОЖЕНИЕ\s+([А-Я]|\d+)(\.\d+)*(\s|$)", RegexOptions.IgnoreCase);

            return isAppendix;
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
        /// Определение того что убрать из проверок текста
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headerTexts"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private bool ShouldSkipParagraph(Paragraph paragraph, HashSet<string> headerTexts, Gost gost)
        {
            // 1. Пропуск заголовков (основных и дополнительных)
            if (IsHeaderParagraph(paragraph, gost) || IsAdditionalHeader(paragraph, gost))
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
            
            //9. Пропуск картинок в случае если у них попадается текст в проверку
            if (paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() || paragraph.Descendants<Picture>().Any())
                return true;

            // 10. Пропускаем ВЕСЬ текст после любого "ПРИЛОЖЕНИЕ [А-Я]"
            if (IsInsideAppendixSection(paragraph, gost))
                return true;

            return false;
        }

        /// <summary>
        /// Метод для работы с ПРИЛОЖЕНИЕМ А-Я
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsInsideAppendixSection(Paragraph paragraph, Gost gost)
        {
            // Получаем все абзацы документа
            var allParagraphs = paragraph.Ancestors<Document>()?.FirstOrDefault()?.Descendants<Paragraph>().ToList();
            if (allParagraphs == null)
                return false;

            int currentIndex = allParagraphs.IndexOf(paragraph);
            if (currentIndex < 0)
                return false;

            // Ищем ближайший заголовок "ПРИЛОЖЕНИЕ [А-Я]" ВЫШЕ текущего абзаца
            for (int i = currentIndex - 1; i >= 0; i--)
            {
                var prevParagraph = allParagraphs[i];
                var text = prevParagraph.InnerText?.Trim();

                // Если нашли заголовок "ПРИЛОЖЕНИЕ [А-Я]" → значит текущий абзац внутри раздела приложений
                if (!string.IsNullOrWhiteSpace(text) && Regex.IsMatch(text, @"^ПРИЛОЖЕНИЕ?\s+[А-Я](\s|$)", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                // Если встретили другой заголовок (не приложение) то выходим, т.к. это уже не приложение
                if (IsHeaderParagraph(prevParagraph, gost) || IsAdditionalHeader(prevParagraph, gost))
                    break;
            }

            return false;
        }

        /// <summary>
        /// Определяет нужно ли пропускать Run при проверке
        /// </summary>
        /// <param name="run"></param>
        /// <returns></returns>
        private bool ShouldSkipRun(Run run)
        {
            // Пропуск пустых Run элементов
            if (string.IsNullOrWhiteSpace(run.InnerText))
                return true;

            // Пропуск специальных символов
            if (run.Elements<Break>().Any() || run.Elements<TabChar>().Any())
                return true;

            // Пропуск Run внутри ссылок оглавлений (например, если это гиперссылка)
            if (run.Descendants<Hyperlink>().Any())
                return true;

            return false;
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
        /// Вспомогательный метод который получает список обязательных разделов из строки
        /// </summary>
        /// <param name="gost"></param>
        /// <returns></returns>
        private List<string> GetRequiredSectionsList(Gost gost)
        {

            if (string.IsNullOrEmpty(gost.RequiredSections))
                return new List<string>();

            return gost.RequiredSections.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
        }

        /// <summary>
        /// Проверяет, является ли параграф элементом списка
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
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
        /// Кнопка проверки на соответствие ГОСТу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_LogOut(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            Close();
        }

        /// <summary>
        /// Показывает диалоговое окно подтверждения
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="title"></param>
        /// <param name="message"></param>
        /// <returns></returns>
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