using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок заголовков
    /// </summary>
    public class CheckingeContents
    {
        private readonly WordprocessingDocument _wordDoc;
        private readonly Gost _gost;

        private readonly Func<Run, bool> _shouldSkipRun;   // Делаем метод для Run
        private readonly Func<Paragraph, Gost, bool> _isAdditionalHeader; // Для проверки дополнительных заголовков

        public CheckingeContents(WordprocessingDocument wordDoc, Gost gost, Func<Run, bool> shouldSkipRun, Func<Paragraph, Gost, bool> isAdditionalHeader)
        {
            _wordDoc = wordDoc;
            _gost = gost;
            _shouldSkipRun = shouldSkipRun;
            _isAdditionalHeader = isAdditionalHeader;
        }

        /// <summary>
        /// Проверка отступов заголовков
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckHeaderIndentsAsync(List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var errors = new List<TextErrorInfo>();
                bool isValid = true;
                var headerTexts = GetHeaderTexts(paragraphs, _gost);

                // Получаем все стили документа
                var allStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                foreach (var paragraph in paragraphs)
                {
                    if (!headerTexts.Contains(paragraph.InnerText.Trim()))
                        continue;

                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    var paragraphStyle = styleId != null && allStyles.TryGetValue(styleId, out var style) ? style : null;

                    // Получаем отступы
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    var styleIndent = GetStyleIndentation(paragraphStyle, allStyles);
                    indent ??= styleIndent;

                    double leftIndent = indent?.Left?.Value != null ? ConvertTwipsToCm(indent.Left.Value) : 0;
                    double rightIndent = indent?.Right?.Value != null ? ConvertTwipsToCm(indent.Right.Value) : 0;
                    double firstLineIndent = indent?.FirstLine?.Value != null ? ConvertTwipsToCm(indent.FirstLine.Value) : 0;
                    double hangingIndent = indent?.Hanging?.Value != null ? ConvertTwipsToCm(indent.Hanging.Value) : 0;

                    bool hasError = false;
                    var errorDetails = new List<string>();

                    // 1. Проверка левого отступа
                    if (_gost.HeaderIndentLeft.HasValue)
                    {
                        double actualTextIndent = leftIndent;
                        if (hangingIndent > 0)
                            actualTextIndent = leftIndent - hangingIndent;

                        if (Math.Abs(actualTextIndent - _gost.HeaderIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"\n       • Левый отступ заголовка: {actualTextIndent:F2} см (требуется {_gost.HeaderIndentLeft.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // 2. Проверка типа и значения первой строки
                    if (_gost.HeaderFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.HeaderIndentOrOutdent))
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

                        // Тип первой строки
                        if (!string.IsNullOrEmpty(_gost.HeaderIndentOrOutdent))
                        {
                            string actualType = isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет";
                            if (actualType != _gost.HeaderIndentOrOutdent)
                            {
                                errorDetails.Add($"\n       • Тип первой строки: '{actualType}' (требуется '{_gost.HeaderIndentOrOutdent}')");
                                hasError = true;
                            }
                        }

                        // Значение отступа/выступа
                        if (_gost.HeaderFirstLineIndent.HasValue)
                        {
                            double actualValue = isHanging ? hangingIndent : firstLineIndent;
                            if ((isHanging || isFirstLine) && Math.Abs(actualValue - _gost.HeaderFirstLineIndent.Value) > 0.05)
                            {
                                errorDetails.Add($"\n       • {(isHanging ? "Выступ" : "Отступ")} первой строки: {actualValue:F2} см (требуется {_gost.HeaderFirstLineIndent.Value:F2} см)");
                                hasError = true;
                            }
                            else if (!isHanging && !isFirstLine && _gost.HeaderIndentOrOutdent != "Нет")
                            {
                                errorDetails.Add($"\n       • Отсутствует {_gost.HeaderIndentOrOutdent} первой строки");
                                hasError = true;
                            }
                        }
                    }

                    // 3. Проверка правого отступа
                    if (_gost.HeaderIndentRight.HasValue)
                    {
                        if (Math.Abs(rightIndent - _gost.HeaderIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"\n       • Правый отступ: {rightIndent:F2} см (требуется {_gost.HeaderIndentRight.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string headerText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"\n       • Заголовок '{headerText}': {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (errors.Any())
                    {
                        string errorMessage = $"Ошибки в отступах заголовков:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(15))}";

                        if (errors.Count > 3) errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";

                        updateUI?.Invoke(errorMessage, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Отступы заголовков соответствуют ГОСТу", Brushes.Green);
                    }
                });

                return (isValid, errors);
            });
        }

        /// <summary>
        /// Проверка интервалов заголовков
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckHeaderParagraphSpacingAsync(List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
                bool isValid = true;
                var headerTexts = GetHeaderTexts(paragraphs, _gost);

                // Получаем все стили документа для проверки наследования
                var allStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                foreach (var paragraph in paragraphs)
                {
                    // Проверяем по тексту параграфа, а не Run
                    if (!headerTexts.Contains(paragraph.InnerText.Trim()))
                        continue;

                    var firstRun = paragraph.Elements<Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.InnerText));
                    if (firstRun == null) continue;

                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    var paragraphStyle = styleId != null && allStyles.TryGetValue(styleId, out var style) ? style : null;

                    var errorMessages = new List<string>();

                    // Проверка выравнивания
                    if (!string.IsNullOrEmpty(_gost.HeaderAlignment))
                    {
                        var (actualAlignment, isAlignmentDefined) = GetActualAlignmentForHeader(paragraph, paragraphStyle, allStyles);

                        if (!isAlignmentDefined)
                        {
                            errorMessages.Add($"\n       • не удалось определить выравнивание");
                            isValid = false;
                        }
                        else if (actualAlignment != _gost.HeaderAlignment)
                        {
                            errorMessages.Add($"\n       • выравнивание: '{actualAlignment}' (требуется '{_gost.HeaderAlignment}')");
                            isValid = false;
                        }
                    }

                    // Проверка междустрочного интервала
                    if (_gost.HeaderLineSpacingValue.HasValue || !string.IsNullOrEmpty(_gost.HeaderLineSpacingType))
                    {
                        var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacingForHeader(paragraph, paragraphStyle, allStyles);

                        if (!isSpacingDefined)
                        {
                            errorMessages.Add($"\n       • не удалось определить междустрочный интервал");
                            isValid = false;
                        }
                        else
                        {
                            // Проверка типа интервала
                            if (!string.IsNullOrEmpty(_gost.HeaderLineSpacingType))
                            {
                                string requiredType = _gost.HeaderLineSpacingType;
                                if (actualSpacingType != requiredType)
                                {
                                    errorMessages.Add($"\n       • тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                                    isValid = false;
                                }
                            }

                            // Проверка значения интервала
                            if (_gost.HeaderLineSpacingValue.HasValue && Math.Abs(actualSpacingValue - _gost.HeaderLineSpacingValue.Value) > 0.01)
                            {
                                errorMessages.Add($"\n       • межстрочный интервал: {actualSpacingValue:F2} (требуется {_gost.HeaderLineSpacingValue.Value:F2})");
                                isValid = false;
                            }
                        }
                    }

                    // Проверка интервалов перед/после
                    if (_gost.HeaderLineSpacingBefore.HasValue || _gost.HeaderLineSpacingAfter.HasValue)
                    {
                        var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacingForHeader(paragraph, paragraphStyle, allStyles);

                        if (!isSpacingDefined)
                        {
                            errorMessages.Add($"\n       • не удалось определить интервалы перед/после");
                            isValid = false;
                        }
                        else
                        {
                            if (_gost.HeaderLineSpacingBefore.HasValue && Math.Abs(actualBefore - _gost.HeaderLineSpacingBefore.Value) > 0.01)
                            {
                                errorMessages.Add($"\n       • интервал перед: {actualBefore:F1} pt (требуется {_gost.HeaderLineSpacingBefore.Value:F1} pt)");
                                isValid = false;
                            }

                            if (_gost.HeaderLineSpacingAfter.HasValue && Math.Abs(actualAfter - _gost.HeaderLineSpacingAfter.Value) > 0.01)
                            {
                                errorMessages.Add($"\n       • интервал после: {actualAfter:F1} pt (требуется {_gost.HeaderLineSpacingAfter.Value:F1} pt)");
                                isValid = false;
                            }
                        }
                    }

                    if (errorMessages.Any())
                    {
                        string headerText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"\n       • Заголовок '{headerText}': {string.Join(", ", errorMessages)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = firstRun
                        });
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (!isValid)
                    {
                        var msg = $"Ошибки в заголовках:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(15))}";

                        if (tempErrors.Count > 3)
                            msg += $"\n...и ещё {tempErrors.Count - 3} ошибок";

                        updateUI?.Invoke(msg, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Интервалы и выравнивание заголовков соответствуют ГОСТу", Brushes.Green);
                    }
                });

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Проверка основных заголовков (наличие и тип с размером шрифта)
        /// </summary>
        /// <param name="gost"></param>
        /// <param name="paragraphs"></param>
        /// <param name="doc"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckRequiredSectionsAsync(Gost gost, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var errors = new List<TextErrorInfo>();

                if (string.IsNullOrEmpty(gost.RequiredSections))
                    return (true, errors);

                string requiredFont = gost.HeaderFontName;
                double? requiredSize = gost.HeaderFontSize;

                bool checkFont = !string.IsNullOrEmpty(requiredFont);
                bool checkSize = requiredSize.HasValue;

                var requiredSections = GetRequiredSectionsList(gost);
                bool allSectionsFound = true;
                bool allSectionsValid = true;
                var missingSections = new List<string>();
                var invalidSections = new List<string>();

                // Получаем все стили документа для проверки наследования
                var allStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                // 1. Проверяем дублирование заголовков
                var (hasDuplicates, duplicateErrors) = CheckDuplicateMainHeaders(paragraphs, requiredSections);
                if (hasDuplicates)
                {
                    errors.AddRange(duplicateErrors);
                }

                // 2. Проверяем обязательные разделы (Введение, Заключение и т.д.)
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
                            var firstRun = paragraph.Elements<Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.InnerText));

                            if (firstRun != null)
                            {
                                var errorMessages = new List<string>();

                                // Проверка шрифта для обязательных разделов
                                if (checkFont)
                                {
                                    var (actualFont, isFontDefined) = GetActualFontForHeader(firstRun, paragraph, allStyles);

                                    if (!isFontDefined)
                                    {
                                        errorMessages.Add("\n       • не удалось определить шрифт");
                                        sectionValid = false;
                                        invalidSections.Add(section);
                                    }
                                    else if (!string.Equals(actualFont, requiredFont, StringComparison.OrdinalIgnoreCase))
                                    {
                                        errorMessages.Add($"\n       • неверный шрифт: '{actualFont}' (требуется: '{requiredFont}')");
                                        sectionValid = false;
                                        invalidSections.Add(section);
                                    }
                                }

                                // Проверка размера шрифта для обязательных разделов
                                if (checkSize)
                                {
                                    var (actualSize, isSizeDefined) = GetActualFontSizeForHeader(firstRun, paragraph, allStyles);

                                    if (!isSizeDefined)
                                    {
                                        errorMessages.Add("\n       • не удалось определить размер шрифта");
                                        sectionValid = false;
                                        invalidSections.Add(section);
                                    }
                                    else if (Math.Abs(actualSize - requiredSize.Value) > 0.1)
                                    {
                                        errorMessages.Add($"\n       • неверный размер: {actualSize:F1} pt (требуется: {requiredSize.Value:F1})");
                                        sectionValid = false;
                                        invalidSections.Add(section);
                                    }
                                }

                                if (errorMessages.Any())
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"\n       {section} {string.Join(": ", errorMessages)}",
                                        ProblemRun = firstRun,
                                        ProblemParagraph = paragraph
                                    });
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

                // 3. ПРОВЕРКА ПРИЛОЖЕНИЙ 
                foreach (var paragraph in paragraphs)
                {
                    var text = paragraph.InnerText.Trim();
                    var match = Regex.Match(text, @"^(ПРИЛОЖЕНИЕ\s+[А-Я0-9]+(\.\d+)*)", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        string appendixFullName = match.Groups[1].Value.ToUpper();
                        var firstRun = paragraph.Elements<Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.InnerText));

                        if (firstRun != null)
                        {
                            var errorMessages = new List<string>();

                            if (checkFont)
                            {
                                var (actualFont, isFontDefined) = GetActualFontForHeader(firstRun, paragraph, allStyles);
                                if (!isFontDefined)
                                {
                                    errorMessages.Add("\n       • не удалось определить шрифт");
                                    allSectionsValid = false;
                                }
                                else if (!string.Equals(actualFont, requiredFont, StringComparison.OrdinalIgnoreCase))
                                {
                                    errorMessages.Add($"\n       • неверный шрифт: '{actualFont}' (требуется: '{requiredFont}')");
                                    allSectionsValid = false;
                                }
                            }

                            if (checkSize)
                            {
                                var (actualSize, isSizeDefined) = GetActualFontSizeForHeader(firstRun, paragraph, allStyles);
                                if (!isSizeDefined)
                                {
                                    errorMessages.Add("\n       • не удалось определить размер шрифта");
                                    allSectionsValid = false;
                                }
                                else if (Math.Abs(actualSize - requiredSize.Value) > 0.1)
                                {
                                    errorMessages.Add($"\n       • неверный размер: {actualSize:F1} pt (требуется: {requiredSize.Value:F1} pt)");
                                    allSectionsValid = false;
                                }
                            }

                            if (errorMessages.Any())
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"\n       {appendixFullName} {string.Join(": ", errorMessages)}",
                                    ProblemRun = firstRun,
                                    ProblemParagraph = paragraph
                                });
                            }
                        }
                    }
                }

                // Добавляем ошибки для отсутствующих разделов
                if (!allSectionsFound)
                {
                    errors.AddRange(missingSections.Select(s => new TextErrorInfo
                    {
                        ErrorMessage = $"\n       • Отсутствует раздел: {s}",
                        ProblemRun = null,
                        ProblemParagraph = null
                    }));
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (errors.Any())
                    {
                        string errorMessage = "Ошибки шрифта и размера в основных заголовках:\n" + string.Join("\n", 
                                              errors.Where(e => !string.IsNullOrEmpty(e.ErrorMessage)).Select(e => e.ErrorMessage).Take(15));

                        if (errors.Count > 3)
                            errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";

                        updateUI?.Invoke(errorMessage, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Основные заголовки соответствуют ГОСТу", Brushes.Green);
                    }
                });

                return (allSectionsFound && allSectionsValid, errors);
            });
        }

        /// <summary>
        /// метод для проверки дополнительных заголовков
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="paragraphs"></param>
        /// <param name="gost"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckAdditionalHeadersAsync(WordprocessingDocument doc, List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var errors = new List<TextErrorInfo>();
                bool isValid = true;

                // 1. Сначала проверяем дублирование номеров
                var (hasDuplicates, duplicateErrors) = CheckDuplicateHeaderNumbers(paragraphs, gost);
                if (hasDuplicates)
                {
                    errors.AddRange(duplicateErrors);
                    isValid = false;
                }

                // Получаем все стили документа для проверки наследования
                var allStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                foreach (var paragraph in paragraphs)
                {
                    if (!_isAdditionalHeader(paragraph, gost))
                        continue;

                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    var paragraphStyle = styleId != null && allStyles.TryGetValue(styleId, out var style) ? style : null;

                    bool hasError = false;
                    var errorDetails = new List<string>();

                    // 1. Проверка выравнивания
                    if (!string.IsNullOrEmpty(gost.AdditionalHeaderAlignment))
                    {
                        var (actualAlignment, isAlignmentDefined) = GetActualAlignmentForHeader(paragraph, paragraphStyle, allStyles);

                        if (!isAlignmentDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить выравнивание");
                            hasError = true;
                        }
                        else if (actualAlignment != gost.AdditionalHeaderAlignment)
                        {
                            errorDetails.Add($"\n          • выравнивание: '{actualAlignment}' (требуется '{gost.AdditionalHeaderAlignment}')");
                            hasError = true;
                        }
                    }

                    // 2. Проверка междустрочного интервала
                    if (gost.AdditionalHeaderLineSpacingValue.HasValue || !string.IsNullOrEmpty(gost.AdditionalHeaderLineSpacingType))
                    {
                        var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacingForHeader(paragraph, paragraphStyle, allStyles);

                        if (!isSpacingDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить междустрочный интервал");
                            hasError = true;
                        }
                        else
                        {
                            // Проверка типа интервала
                            if (!string.IsNullOrEmpty(gost.AdditionalHeaderLineSpacingType))
                            {
                                string requiredType = gost.AdditionalHeaderLineSpacingType;
                                if (actualSpacingType != requiredType)
                                {
                                    errorDetails.Add($"\n          • тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                                    hasError = true;
                                }
                            }

                            // Проверка значения интервала
                            if (gost.AdditionalHeaderLineSpacingValue.HasValue && Math.Abs(actualSpacingValue - gost.AdditionalHeaderLineSpacingValue.Value) > 0.01)
                            {
                                errorDetails.Add($"\n          • межстрочный интервал: {actualSpacingValue:F2} (требуется {gost.AdditionalHeaderLineSpacingValue.Value:F2})");
                                hasError = true;
                            }
                        }
                    }

                    // 3. Проверка интервалов перед/после
                    if (gost.AdditionalHeaderLineSpacingBefore.HasValue || gost.AdditionalHeaderLineSpacingAfter.HasValue)
                    {
                        var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacingForHeader(paragraph, paragraphStyle, allStyles);

                        if (!isSpacingDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить интервалы перед/после");
                            hasError = true;
                        }
                        else
                        {
                            if (gost.AdditionalHeaderLineSpacingBefore.HasValue && Math.Abs(actualBefore - gost.AdditionalHeaderLineSpacingBefore.Value) > 0.1)
                            {
                                errorDetails.Add($"\n          • интервал перед: {actualBefore:F1} pt (требуется {gost.AdditionalHeaderLineSpacingBefore.Value:F1} pt)");
                                hasError = true;
                            }

                            if (gost.AdditionalHeaderLineSpacingAfter.HasValue && Math.Abs(actualAfter - gost.AdditionalHeaderLineSpacingAfter.Value) > 0.1)
                            {
                                errorDetails.Add($"\n          • интервал после: {actualAfter:F1} pt (требуется {gost.AdditionalHeaderLineSpacingAfter.Value:F1} pt)");
                                hasError = true;
                            }
                        }
                    }

                    // 4. Проверка шрифта
                    if (!string.IsNullOrEmpty(gost.AdditionalHeaderFontName))
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            if (_shouldSkipRun(run)) continue;

                            var (actualFont, isFontDefined) = GetActualFontForHeader(run, paragraph, allStyles);

                            if (!isFontDefined)
                            {
                                errorDetails.Add($"\n       • не удалось определить шрифт");
                                hasError = true;
                                break;
                            }
                            else if (!string.Equals(actualFont, gost.AdditionalHeaderFontName, StringComparison.OrdinalIgnoreCase))
                            {
                                errorDetails.Add($"\n          • шрифт: '{actualFont}' (требуется '{gost.AdditionalHeaderFontName}')");
                                hasError = true;
                                break;
                            }
                        }
                    }

                    // 5. Проверка размера шрифта
                    if (gost.AdditionalHeaderFontSize.HasValue)
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            if (_shouldSkipRun(run)) continue;

                            var (actualSize, isSizeDefined) = GetActualFontSizeForHeader(run, paragraph, allStyles);

                            if (!isSizeDefined)
                            {
                                errorDetails.Add($"\n       • не удалось определить размер шрифта");
                                hasError = true;
                                break;
                            }
                            else if (Math.Abs(actualSize - gost.AdditionalHeaderFontSize.Value) > 0.1)
                            {
                                errorDetails.Add($"\n          • размер шрифта: {actualSize:F1} pt (требуется {gost.AdditionalHeaderFontSize.Value:F1} pt)");
                                hasError = true;
                                break;
                            }
                        }
                    }

                    // Получаем отступы для дополнительных проверок
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    var styleIndent = GetStyleIndentation(paragraphStyle, allStyles);
                    indent ??= styleIndent;

                    double leftIndent = indent?.Left?.Value != null ? ConvertTwipsToCm(indent.Left.Value) : 0;
                    double rightIndent = indent?.Right?.Value != null ? ConvertTwipsToCm(indent.Right.Value) : 0;
                    double firstLineIndent = indent?.FirstLine?.Value != null ? ConvertTwipsToCm(indent.FirstLine.Value) : 0;
                    double hangingIndent = indent?.Hanging?.Value != null ? ConvertTwipsToCm(indent.Hanging.Value) : 0;

                    // 6. Проверка левого отступа
                    if (gost.AdditionalHeaderIndentLeft.HasValue)
                    {
                        double actualTextIndent = leftIndent;

                        if (hangingIndent > 0)
                            actualTextIndent = leftIndent - hangingIndent;

                        if (Math.Abs(actualTextIndent - gost.AdditionalHeaderIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"\n          • Левый отступ: {actualTextIndent:F2} см (требуется {gost.AdditionalHeaderIndentLeft.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // 7. Проверка типа и значения первой строки
                    if (gost.AdditionalHeaderFirstLineIndent.HasValue || !string.IsNullOrEmpty(gost.AdditionalHeaderIndentOrOutdent))
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

                        // Тип первой строки
                        if (!string.IsNullOrEmpty(gost.AdditionalHeaderIndentOrOutdent))
                        {
                            string actualType = isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет";
                            if (actualType != gost.AdditionalHeaderIndentOrOutdent)
                            {
                                errorDetails.Add($"\n          • Тип первой строки: '{actualType}' (требуется '{gost.AdditionalHeaderIndentOrOutdent}')");
                                hasError = true;
                            }
                        }

                        // Значение отступа/выступа
                        if (gost.AdditionalHeaderFirstLineIndent.HasValue)
                        {
                            double actualValue = isHanging ? hangingIndent : firstLineIndent;
                            if ((isHanging || isFirstLine) && Math.Abs(actualValue - gost.AdditionalHeaderFirstLineIndent.Value) > 0.05)
                            {
                                errorDetails.Add($"\n          • {(isHanging ? "Выступ" : "Отступ")} первой строки: {actualValue:F2} см (требуется {gost.AdditionalHeaderFirstLineIndent.Value:F2} см)");
                                hasError = true;
                            }
                            else if (!isHanging && !isFirstLine && gost.AdditionalHeaderIndentOrOutdent != "Нет")
                            {
                                errorDetails.Add($"\n          • Отсутствует {gost.AdditionalHeaderIndentOrOutdent} первой строки");
                                hasError = true;
                            }
                        }
                    }

                    // 8. Проверка правого отступа
                    if (gost.AdditionalHeaderIndentRight.HasValue)
                    {
                        if (Math.Abs(rightIndent - gost.AdditionalHeaderIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"\n          • Правый отступ: {rightIndent:F2} см (требуется {gost.AdditionalHeaderIndentRight.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string shortText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"\n       • Доп. заголовок '{shortText}': {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (errors.Any())
                    {
                        string errorMessage = "Ошибки в доп. заголовках:\n" +
                            string.Join("\n", errors.Select(e => e.ErrorMessage).Take(15));

                        if (errors.Count > 3)
                            errorMessage += $"\n...и ещё {errors.Count - 3} ошибок";

                        updateUI?.Invoke(errorMessage, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Дополнительные заголовки соответствуют ГОСТу", Brushes.Green);
                    }
                });

                return (isValid, errors);
            });
        }

        /// <summary>
        /// Проверяет дублирование номеров у заголовков
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="gost"></param>
        /// <returns></returns>
        private (bool HasDuplicates, List<TextErrorInfo> Errors) CheckDuplicateHeaderNumbers(List<Paragraph> paragraphs, Gost gost)
        {
            var errors = new List<TextErrorInfo>();
            var headerGroups = new Dictionary<string, List<(Paragraph paragraph, string fullText)>>();
            bool hasDuplicates = false;

            // Группируем заголовки по номерам
            foreach (var paragraph in paragraphs)
            {
                if (!_isAdditionalHeader(paragraph, gost))
                    continue;

                var text = paragraph.InnerText?.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                var match = Regex.Match(text, @"^(\d+(?:\.\d+)*)[\s\t]+(.+)");
                if (!match.Success) continue;

                var number = match.Groups[1].Value;
                var headerText = $"{number} {match.Groups[2].Value.Trim()}";

                if (!headerGroups.ContainsKey(number))
                {
                    headerGroups[number] = new List<(Paragraph, string)>();
                }
                headerGroups[number].Add((paragraph, headerText));
            }

            // Обрабатываем группы с дубликатами
            foreach (var group in headerGroups.Where(g => g.Value.Count > 1))
            {
                hasDuplicates = true;

                var errorMessage = new StringBuilder("\n       • Обнаружено дублирование номеров в заголовках!");
                foreach (var header in group.Value)
                {
                    errorMessage.Append($"\n            - {header.fullText}");
                }
                errorMessage.Append($"\n            Данные дополнительные заголовки имеют одинаковый номер: '{group.Key}'");

                // Добавляем одно сообщение для отображения ошибки
                errors.Add(new TextErrorInfo
                {
                    ErrorMessage = errorMessage.ToString(),
                    ProblemParagraph = group.Value.First().paragraph,
                    ProblemRun = null
                });

                foreach (var header in group.Value.Skip(1))
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = null,
                        ProblemParagraph = header.paragraph,
                        ProblemRun = null
                    });
                }
            }

            return (hasDuplicates, errors);
        }

        /// <summary>
        /// Проверяет дублирование номеров у главных заголовков
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="gost"></param>
        /// <param name="requiredSections"></param>
        /// <returns></returns>
        private (bool HasDuplicates, List<TextErrorInfo> Errors) CheckDuplicateMainHeaders(List<Paragraph> paragraphs, List<string> requiredSections)
        {
            var errors = new List<TextErrorInfo>();
            var headerGroups = new Dictionary<string, List<Paragraph>>(StringComparer.OrdinalIgnoreCase);
            bool hasDuplicates = false;

            // Собираем все заголовки, которые соответствуют обязательным разделам
            foreach (var paragraph in paragraphs)
            {
                var text = paragraph.InnerText?.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                var cleanText = Regex.Replace(text, @"^\d+[\s\.]*", "").Trim();

                // Проверяем, является ли этот заголовок обязательным разделом
                if (requiredSections.Contains(cleanText, StringComparer.OrdinalIgnoreCase))
                {
                    if (!headerGroups.ContainsKey(cleanText))
                    {
                        headerGroups[cleanText] = new List<Paragraph>();
                    }
                    headerGroups[cleanText].Add(paragraph);
                }
            }

            // Проверяем дубликаты
            foreach (var group in headerGroups)
            {
                if (group.Value.Count > 1) 
                {
                    hasDuplicates = true;

                    var errorMessage = new StringBuilder($"\n       • Обнаружено дублирование заголовка '{group.Key}':");
                    foreach (var para in group.Value)
                    {
                        errorMessage.Append($"\n            - {GetShortText2(para.InnerText.Trim())}");
                    }

                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = errorMessage.ToString(),
                        ProblemParagraph = group.Value.First(),
                        ProblemRun = null
                    });

                    foreach (var para in group.Value.Skip(1))
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = null,
                            ProblemParagraph = para,
                            ProblemRun = null
                        });
                    }
                }
            }

            return (hasDuplicates, errors);
        }

        private Indentation GetStyleIndentation(Style style, Dictionary<string, Style> allStyles)
        {
            while (style != null)
            {
                var indent = style.StyleParagraphProperties?.Indentation;
                if (indent != null)
                    return indent;

                if (style.BasedOn?.Val?.Value != null && allStyles.TryGetValue(style.BasedOn.Val.Value, out var basedOnStyle))
                    style = basedOnStyle;
                else
                    break;
            }
            return null;
        }

        private (string Alignment, bool IsDefined) GetActualAlignmentForHeader(Paragraph paragraph, Style paragraphStyle, Dictionary<string, Style> allStyles)
        {
            // 1. Проверяем явное выравнивание в параграфе
            if (paragraph.ParagraphProperties?.Justification?.Val?.Value != null)
            {
                return (GetAlignmentString(paragraph.ParagraphProperties.Justification), true);
            }

            // 2. Проверяем стиль параграфа и его родителей
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var currentStyle))
            {
                while (currentStyle != null)
                {
                    if (currentStyle.StyleParagraphProperties?.Justification?.Val?.Value != null)
                    {
                        return (GetAlignmentString(currentStyle.StyleParagraphProperties.Justification), true);
                    }
                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            // 3. Проверяем переданный стиль (если есть)
            if (paragraphStyle != null)
            {
                if (paragraphStyle.StyleParagraphProperties?.Justification?.Val?.Value != null)
                {
                    return (GetAlignmentString(paragraphStyle.StyleParagraphProperties.Justification), true);
                }
            }

            // 4. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle) && normalStyle.StyleParagraphProperties?.Justification?.Val?.Value != null)
            {
                return (GetAlignmentString(normalStyle.StyleParagraphProperties.Justification), true);
            }

            return ("Left", true);
        }

        private (string Type, double Value, bool IsDefined) GetActualLineSpacingForHeader(Paragraph paragraph, Style paragraphStyle, Dictionary<string, Style> allStyles)
        {
            // 1. Явные свойства
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            var parsed = ParseLineSpacing(spacing);

            if (parsed.IsDefined)
                return parsed;

            // 2. Стиль абзаца и его родители
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var currentStyle))
            {
                while (currentStyle != null)
                {
                    var styleSpacing = currentStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    parsed = ParseLineSpacing(styleSpacing);
                    if (parsed.IsDefined)
                        return parsed;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            // 3. Проверяем переданный стиль (если есть)
            if (paragraphStyle != null)
            {
                var styleSpacing = paragraphStyle.StyleParagraphProperties?.SpacingBetweenLines;
                parsed = ParseLineSpacing(styleSpacing);
                if (parsed.IsDefined)
                    return parsed;
            }

            // 4. Normal
            if (allStyles.TryGetValue("Normal", out var normalStyle))
            {
                parsed = ParseLineSpacing(normalStyle.StyleParagraphProperties?.SpacingBetweenLines);
                if (parsed.IsDefined)
                    return parsed;
            }

            return ("Множитель", 1.0, true);
        }

        private (double Before, double After, bool IsDefined) GetActualParagraphSpacingForHeader(Paragraph paragraph, Style paragraphStyle, Dictionary<string, Style> allStyles)
        {
            double? before = null;
            double? after = null;

            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing?.Before?.Value != null) before = ConvertTwipsToPoints(spacing.Before.Value);
            if (spacing?.After?.Value != null) after = ConvertTwipsToPoints(spacing.After.Value);

            if (before == null || after == null)
            {
                var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                var currentStyle = paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var style) ? style : paragraphStyle;

                while (currentStyle != null)
                {
                    var styleSpacing = currentStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    if (styleSpacing != null)
                    {
                        if (before == null && styleSpacing.Before?.Value != null)
                            before = ConvertTwipsToPoints(styleSpacing.Before.Value);

                        if (after == null && styleSpacing.After?.Value != null)
                            after = ConvertTwipsToPoints(styleSpacing.After.Value);
                    }

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            if (before == null || after == null)
            {
                if (allStyles.TryGetValue("Normal", out var normalStyle))
                {
                    var spacingNorm = normalStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    if (spacingNorm != null)
                    {
                        if (before == null && spacingNorm.Before?.Value != null)
                            before = ConvertTwipsToPoints(spacingNorm.Before.Value);

                        if (after == null && spacingNorm.After?.Value != null)
                            after = ConvertTwipsToPoints(spacingNorm.After.Value);
                    }
                }
            }

            var isDefined = before.HasValue || after.HasValue;
            return (before ?? 0, after ?? 0, isDefined);
        }

        private (string Type, double Value, bool IsDefined) ParseLineSpacing(SpacingBetweenLines spacing)
        {
            // Если spacing вообще не задан - возвращаем неопределенное значение
            if (spacing == null)
            {
                return (null, 0, false);
            }

            // Если есть LineRule, но нет Line - считаем неопределенным
            if (spacing.LineRule != null && spacing.Line == null)
            {
                return (null, 0, false);
            }

            // Если есть Line, но нет LineRule - интерпретируем как множитель
            if (spacing.Line != null && spacing.LineRule == null)
            {
                double lineValue = double.Parse(spacing.Line.Value);
                return ("Множитель", lineValue / 240.0, true);
            }

            // Если оба значения заданы
            if (spacing.Line != null && spacing.LineRule != null)
            {
                double lineValue = double.Parse(spacing.Line.Value);

                if (spacing.LineRule.Value == LineSpacingRuleValues.Exact)
                    return ("Точно", lineValue / 567.0, true);

                if (spacing.LineRule.Value == LineSpacingRuleValues.AtLeast)
                    return ("Минимум", lineValue / 567.0, true);

                // По умолчанию - множитель
                return ("Множитель", lineValue / 240.0, true);
            }

            // Если ничего не задано - не определено
            return (null, 0, false);
        }

        private double ConvertTwipsToCm(string twipsValue)
        {
            if (string.IsNullOrEmpty(twipsValue))
                return 0;

            if (double.TryParse(twipsValue, out double twips))
                return twips / 567.0;

            return 0;
        }

        private (string FontName, bool IsDefined) GetActualFontForHeader(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles)
        {
            // 1. Проверяем явные свойства Run
            var explicitFont = GetExplicitRunFont(run);
            if (!string.IsNullOrEmpty(explicitFont))
                return (explicitFont, true);

            // 2. Проверяем стиль Run
            var runStyleId = run.RunProperties?.RunStyle?.Val?.Value;
            if (runStyleId != null && allStyles.TryGetValue(runStyleId, out var runStyle))
            {
                var runStyleFont = GetStyleFont(runStyle);
                if (!string.IsNullOrEmpty(runStyleFont))
                    return (runStyleFont, true);
            }

            // 3. Проверяем стиль Paragraph с учетом наследования
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null)
            {
                var currentStyle = allStyles.TryGetValue(paraStyleId, out var style) ? style : null;
                while (currentStyle != null)
                {
                    var font = GetStyleFont(currentStyle);
                    if (!string.IsNullOrEmpty(font))
                        return (font, true);

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            // 4. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle))
            {
                var font = GetStyleFont(normalStyle);
                if (!string.IsNullOrEmpty(font))
                    return (font, true);
            }

            return (null, false);
        }

        private (double Size, bool IsDefined) GetActualFontSizeForHeader(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles)
        {
            // 1. Проверяем явные свойства Run
            var runSize = run.RunProperties?.FontSize?.Val?.Value;
            if (runSize != null)
                return (double.Parse(runSize) / 2, true);

            // 2. Проверяем стиль Run
            var runStyleId = run.RunProperties?.RunStyle?.Val?.Value;
            if (runStyleId != null && allStyles.TryGetValue(runStyleId, out var runStyle))
            {
                if (runStyle?.StyleRunProperties?.FontSize?.Val?.Value != null)
                    return (double.Parse(runStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);
            }

            // 3. Проверяем стиль Paragraph с учетом наследования
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var paraStyle))
            {
                var currentStyle = paraStyle;
                while (currentStyle != null)
                {
                    if (currentStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
                        return (double.Parse(currentStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            // 4. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle) && normalStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
            {
                return (double.Parse(normalStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);
            }

            return (0, false);
        }

        private string GetExplicitRunFont(Run run)
        {
            var runProps = run.RunProperties;
            if (runProps == null) return null;

            return runProps.RunFonts?.Ascii?.Value ?? runProps.RunFonts?.HighAnsi?.Value ?? runProps.RunFonts?.ComplexScript?.Value ?? runProps.RunFonts?.EastAsia?.Value;
        }

        private string GetStyleFont(Style style)
        {
            if (style?.StyleRunProperties == null) return null;
            return style.StyleRunProperties.RunFonts?.Ascii?.Value ?? style.StyleRunProperties.RunFonts?.HighAnsi?.Value ?? style.StyleRunProperties.RunFonts?.ComplexScript?.Value ?? style.StyleRunProperties.RunFonts?.EastAsia?.Value;
        }

        /// <summary>
        /// Вспомогательный метод. Получает тексты заголовков из тела документа на основе обязательных разделов ГОСТа
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

                if (string.IsNullOrEmpty(text)) continue;

                // Проверка на приложение
                bool isAppendix = Regex.IsMatch(text, @"^ПРИЛОЖЕНИЕ\s+([А-Я]|\d+)(\.\d+)*(\s|$)", RegexOptions.IgnoreCase);
                if (isAppendix)
                {
                    headerTexts.Add(text);
                    continue;
                }

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
    }
}
