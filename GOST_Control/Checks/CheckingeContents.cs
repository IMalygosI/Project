using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia.Media;
using Avalonia.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок заголовков
    /// </summary>
    public class CheckingeContents
    {

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

        // -- Доп Значения для Доп Заголовков
        private const string DefaultAdditionalHeaderFontName = "Arial";
        private const double DefaultAdditionalHeaderFontSize = 20.0;
        private const string DefaultAdditionalHeaderAlignment = "Left";
        private const string DefaultAdditionalHeaderLineSpacingType = "Множитель";
        private const double DefaultAdditionalHeaderLineSpacingValue = 1.15;
        private const double DefaultAdditionalHeaderLineSpacingBefore = 0.85;
        private const double DefaultAdditionalHeaderLineSpacingAfter = 0.35;
        private const string DefaultAdditionalHeaderIndentOrOutdent = "Нет";
        private const double DefaultAdditionalHeaderFirstLineIndent = 0.0;
        private const double DefaultAdditionalHeaderIndentLeftt = 0.0;
        private const double DefaultAdditionalHeaderIndentRight = 0.0;

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
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckHeaderIndentsAsync(List<Paragraph> paragraphs, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                bool isValid = true;
                var errors = new List<TextErrorInfo>();
                var headerTexts = GetHeaderTexts(paragraphs, _gost);

                foreach (var paragraph in paragraphs)
                {
                    if (!headerTexts.Contains(paragraph.InnerText.Trim()))
                        continue;

                    var indent = paragraph.ParagraphProperties?.Indentation;
                    bool hasError = false;
                    var errorDetails = new List<string>();

                    // Преобразуем все значения в сантиметры
                    double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                    double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                    double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

                    // 1. ПРОВЕРКА ВИЗУАЛЬНОГО ЛЕВОГО ОТСТУПА ЗАГОЛОВКА
                    if (_gost.HeaderIndentLeft.HasValue)
                    {
                        double actualTextIndent = leftIndent; // Базовый отступ

                        // Корректировка если есть выступ (hanging)
                        if (hangingIndent > 0)
                        {
                            actualTextIndent = leftIndent - hangingIndent;
                        }

                        if (Math.Abs(actualTextIndent - _gost.HeaderIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"Левый отступ заголовка: {actualTextIndent:F2} см (требуется {_gost.HeaderIndentLeft.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // 2. ПРОВЕРКА ПЕРВОЙ СТРОКИ ЗАГОЛОВКА
                    if (_gost.HeaderFirstLineIndent.HasValue)
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

                        // Проверка типа (выступ/отступ)
                        if (!string.IsNullOrEmpty(_gost.HeaderIndentOrOutdent))
                        {
                            bool typeError = false;

                            if (_gost.HeaderIndentOrOutdent == "Выступ" && !isHanging)
                                typeError = true;
                            else if (_gost.HeaderIndentOrOutdent == "Отступ" && !isFirstLine)
                                typeError = true;
                            else if (_gost.HeaderIndentOrOutdent == "Нет" && (isHanging || isFirstLine))
                                typeError = true;

                            if (typeError)
                            {
                                errorDetails.Add($"Тип первой строки: {(isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет")} (требуется {_gost.HeaderIndentOrOutdent})");
                                hasError = true;
                            }
                        }

                        // Проверка значения
                        double currentValue = isHanging ? hangingIndent : firstLineIndent;
                        if ((isHanging || isFirstLine) && Math.Abs(currentValue - _gost.HeaderFirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{(isHanging ? "Выступ" : "Отступ")} первой строки: {currentValue:F2} см (требуется {_gost.HeaderFirstLineIndent.Value:F2} см)");
                            hasError = true;
                        }
                        else if (_gost.HeaderIndentOrOutdent != "Нет" && !isHanging && !isFirstLine)
                        {
                            errorDetails.Add($"Отсутствует {_gost.HeaderIndentOrOutdent} первой строки");
                            hasError = true;
                        }
                    }

                    // 3. ПРОВЕРКА ПРАВОГО ОТСТУПА ЗАГОЛОВКА
                    if (_gost.HeaderIndentRight.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultHeaderRightIndent;

                        if (Math.Abs(actualRight - _gost.HeaderIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"Правый отступ: {actualRight:F2} см (требуется {_gost.HeaderIndentRight.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string headerText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"Заголовок '{headerText}': {string.Join(", ", errorDetails)}",
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
                        string errorMessage = $"Ошибки в отступах заголовков:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
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
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckHeaderParagraphSpacingAsync(List<Paragraph> paragraphs, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                bool hasErrors = false;
                var errors = new List<TextErrorInfo>();
                var headerTexts = GetHeaderTexts(paragraphs, _gost);

                var stylesPart = _wordDoc.MainDocumentPart?.StyleDefinitionsPart;
                var styles = stylesPart?.Styles;

                foreach (var paragraph in paragraphs)
                {
                    if (!headerTexts.Contains(paragraph.InnerText.Trim()))
                        continue;

                    var paraErrors = new List<string>();
                    bool paragraphHasErrors = false;

                    var explicitSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;

                    Style style = null;
                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    if (styleId != null && styles != null)
                    {
                        style = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                    }

                    var styleSpacing = style?.StyleParagraphProperties?.SpacingBetweenLines;

                    // Межстрочный интервал
                    if (_gost.HeaderLineSpacingValue.HasValue)
                    {
                        double actualSpacing = DefaultHeaderLineSpacingValue;
                        string actualSpacingType = DefaultHeaderLineSpacingType;
                        LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                        var spacingSource = explicitSpacing ?? styleSpacing;
                        if (spacingSource?.Line != null)
                        {
                            var rule = spacingSource.LineRule?.Value;
                            var val = spacingSource.Line.Value;
                            if (rule == LineSpacingRuleValues.Exact)
                            {
                                actualSpacing = double.Parse(val) / 567.0;
                                actualSpacingType = "Точно";
                                actualRule = LineSpacingRuleValues.Exact;
                            }
                            else if (rule == LineSpacingRuleValues.AtLeast)
                            {
                                actualSpacing = double.Parse(val) / 567.0;
                                actualSpacingType = "Минимум";
                                actualRule = LineSpacingRuleValues.AtLeast;
                            }
                            else
                            {
                                actualSpacing = double.Parse(val) / 240.0;
                                actualSpacingType = "Множитель";
                                actualRule = LineSpacingRuleValues.Auto;
                            }
                        }

                        var requiredRule = (_gost.HeaderLineSpacingType ?? DefaultHeaderLineSpacingType) switch
                        {
                            "Минимум" => LineSpacingRuleValues.AtLeast,
                            "Точно" => LineSpacingRuleValues.Exact,
                            _ => LineSpacingRuleValues.Auto
                        };

                        if (actualRule != requiredRule)
                        {
                            paraErrors.Add($"тип интервала: '{actualSpacingType}' (требуется '{_gost.HeaderLineSpacingType}')");
                            paragraphHasErrors = true;
                        }

                        if (Math.Abs(actualSpacing - _gost.HeaderLineSpacingValue.Value) > 0.1)
                        {
                            paraErrors.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {_gost.HeaderLineSpacingValue.Value:F2})");
                            paragraphHasErrors = true;
                        }
                    }

                    // Интервалы "Перед" и "После"
                    double actualBefore = styleSpacing?.Before?.Value != null ? ConvertTwipsToPoints(styleSpacing.Before.Value) : DefaultHeaderSpacingBefore;
                    if (explicitSpacing?.Before?.Value != null)
                        actualBefore = ConvertTwipsToPoints(explicitSpacing.Before.Value);

                    if (_gost.HeaderLineSpacingBefore.HasValue && Math.Abs(actualBefore - _gost.HeaderLineSpacingBefore.Value) > 0.1)
                    {
                        paraErrors.Add($"интервал перед: {actualBefore:F1} pt (требуется {_gost.HeaderLineSpacingBefore.Value:F1} pt)");
                        paragraphHasErrors = true;
                    }

                    double actualAfter = styleSpacing?.After?.Value != null ? ConvertTwipsToPoints(styleSpacing.After.Value) : DefaultHeaderSpacingAfter;
                    if (explicitSpacing?.After?.Value != null)
                        actualAfter = ConvertTwipsToPoints(explicitSpacing.After.Value);

                    if (_gost.HeaderLineSpacingAfter.HasValue && Math.Abs(actualAfter - _gost.HeaderLineSpacingAfter.Value) > 0.1)
                    {
                        paraErrors.Add($"интервал после: {actualAfter:F1} pt (требуется {_gost.HeaderLineSpacingAfter.Value:F1} pt)");
                        paragraphHasErrors = true;
                    }

                    // Выравнивание
                    if (!string.IsNullOrEmpty(_gost.HeaderAlignment))
                    {
                        var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification);
                        if (currentAlignment != _gost.HeaderAlignment)
                        {
                            paraErrors.Add($"выравнивание: {currentAlignment} (требуется {_gost.HeaderAlignment})");
                            paragraphHasErrors = true;
                        }
                    }

                    if (paragraphHasErrors)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"Заголовок '{paragraph.InnerText.Trim()}': {string.Join(", ", paraErrors)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        hasErrors = true;
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (hasErrors)
                    {
                        var msg = $"Ошибки в заголовках:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3)
                            msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Интервалы и выравнивание заголовков соответствуют ГОСТу", Brushes.Green);
                    }
                });

                return (!hasErrors, errors);
            });
        }

        /// <summary>
        /// Проверка на наличие верных заголовков
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

                            Style paragraphStyle = null;
                            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

                            if (styleId != null && headerStyles.TryGetValue(styleId, out var style))
                            {
                                paragraphStyle = style;
                            }

                            foreach (var run in paragraph.Elements<Run>())
                            {
                                if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                                if (checkFont)
                                {
                                    var font = run.RunProperties?.RunFonts?.Ascii?.Value ??
                                             paragraphStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value ??
                                             DefaultHeaderFont;

                                    if (font != null && !string.Equals(font, requiredFont, StringComparison.OrdinalIgnoreCase))
                                    {
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"{section} (неверный шрифт: '{font}')",
                                            ProblemRun = run,
                                            ProblemParagraph = paragraph
                                        });
                                        sectionValid = false;
                                        invalidSections.Add(section);
                                    }
                                }

                                if (checkSize)
                                {
                                    double? size = null;
                                    var fontSize = run.RunProperties?.FontSize?.Val?.Value ??
                                                 paragraphStyle?.StyleRunProperties?.FontSize?.Val?.Value;

                                    if (fontSize != null)
                                    {
                                        size = double.Parse(fontSize) / 2;
                                    }
                                    else
                                    {
                                        size = DefaultHeaderSize;
                                    }

                                    if (size.HasValue && Math.Abs(size.Value - requiredSize.Value) > 0.1)
                                    {
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"{section} (неверный размер: {size:F1} pt)",
                                            ProblemRun = run,
                                            ProblemParagraph = paragraph
                                        });
                                        sectionValid = false;
                                        invalidSections.Add(section);
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

                // Добавляем ошибки для отсутствующих разделов
                if (!allSectionsFound)
                {
                    errors.AddRange(missingSections.Select(s => new TextErrorInfo
                    {
                        ErrorMessage = $"Отсутствует раздел: {s}",
                        ProblemRun = null,
                        ProblemParagraph = null
                    }));
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (!allSectionsFound)
                    {
                        updateUI?.Invoke($"Не найдены разделы: {string.Join(", ", missingSections)}", Brushes.Red);
                    }
                    else if (!allSectionsValid)
                    {
                        string msg = $"Ошибки в разделах:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3)
                            msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Все обязательные разделы соответствуют требованиям ГОСТ", Brushes.Green);
                    }
                });

                return (allSectionsFound && allSectionsValid, errors);
            });
        }

        /// <summary>
        /// метод для проверки стилей дополнительных заголовков
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
                bool isValid = true;
                var errors = new List<TextErrorInfo>();
                var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
                var styles = stylesPart?.Styles;

                foreach (var paragraph in paragraphs)
                {
                    if (!_isAdditionalHeader(paragraph, gost))
                        continue;

                    bool hasError = false;
                    var errorDetails = new List<string>();
                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    Style style = styleId != null ? styles?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId) : null;

                    // Проверка шрифта
                    if (!string.IsNullOrEmpty(gost.AdditionalHeaderFontName))
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            if (_shouldSkipRun(run)) continue;

                            var font = run.RunProperties?.RunFonts?.Ascii?.Value;

                            if (string.IsNullOrEmpty(font))
                            {
                                font = style?.StyleRunProperties?.RunFonts?.Ascii?.Value ?? DefaultAdditionalHeaderFontName;
                            }

                            if (font != gost.AdditionalHeaderFontName)
                            {
                                errorDetails.Add($"шрифт: '{font}' (требуется '{gost.AdditionalHeaderFontName}')\n");
                                hasError = true;
                            }
                        }
                    }

                    // Проверка размера шрифта
                    if (gost.AdditionalHeaderFontSize.HasValue)
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            if (_shouldSkipRun(run)) continue;

                            var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value ?? style?.StyleRunProperties?.FontSize?.Val?.Value;

                            double actualSize = fontSizeVal != null ? double.Parse(fontSizeVal) / 2 : DefaultAdditionalHeaderFontSize;

                            if (Math.Abs(actualSize - gost.AdditionalHeaderFontSize.Value) > 0.1)
                            {
                                errorDetails.Add($"размер: {actualSize:F1} pt (требуется {gost.AdditionalHeaderFontSize.Value:F1} pt)\n");
                                hasError = true;
                            }
                        }
                    }

                    // Проверка выравнивания
                    if (!string.IsNullOrEmpty(gost.AdditionalHeaderAlignment))
                    {
                        string currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification ?? style?.StyleParagraphProperties?.Justification) ?? DefaultAdditionalHeaderAlignment;

                        if (currentAlignment != gost.AdditionalHeaderAlignment)
                        {
                            errorDetails.Add($"выравнивание: {currentAlignment} (требуется {gost.AdditionalHeaderAlignment})\n");
                            hasError = true;
                        }
                    }

                    // Проверка межстрочного интервала
                    if (gost.AdditionalHeaderLineSpacingValue.HasValue)
                    {
                        double actualSpacing = DefaultAdditionalHeaderLineSpacingValue;
                        string actualSpacingType = DefaultAdditionalHeaderLineSpacingType;
                        LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                        var explicitSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                        var styleSpacing = style?.StyleParagraphProperties?.SpacingBetweenLines;
                        var spacingSource = explicitSpacing ?? styleSpacing;

                        if (spacingSource?.Line != null)
                        {
                            if (spacingSource.LineRule?.Value == LineSpacingRuleValues.Exact)
                            {
                                actualSpacing = double.Parse(spacingSource.Line.Value) / 567.0;
                                actualSpacingType = "Точно";
                                actualRule = LineSpacingRuleValues.Exact;
                            }
                            else if (spacingSource.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                            {
                                actualSpacing = double.Parse(spacingSource.Line.Value) / 567.0;
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

                        // Тип интервала, который должен быть по ГОСТу
                        LineSpacingRuleValues requiredRule = gost.AdditionalHeaderLineSpacingType switch
                        {
                            "Минимум" => LineSpacingRuleValues.AtLeast,
                            "Точно" => LineSpacingRuleValues.Exact,
                            _ => LineSpacingRuleValues.Auto
                        };

                        // Проверка типа интервала
                        if (actualRule != requiredRule)
                        {
                            errorDetails.Add($"тип интервала: '{actualSpacingType}' (требуется '{gost.AdditionalHeaderLineSpacingType}')\n");
                            hasError = true;
                        }

                        // Проверка значения интервала
                        if (Math.Abs(actualSpacing - gost.AdditionalHeaderLineSpacingValue.Value) > 0.1)
                        {
                            errorDetails.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {gost.AdditionalHeaderLineSpacingValue:F2})\n");
                            hasError = true;
                        }
                    }

                    // Проверка интервалов "Перед" и "После"
                    if (gost.AdditionalHeaderLineSpacingBefore.HasValue || gost.AdditionalHeaderLineSpacingAfter.HasValue)
                    {
                        double actualBefore = DefaultAdditionalHeaderLineSpacingBefore;
                        double actualAfter = DefaultAdditionalHeaderLineSpacingAfter;

                        var explicitSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                        var styleSpacing = style?.StyleParagraphProperties?.SpacingBetweenLines;

                        if (explicitSpacing?.Before?.Value != null)
                        {
                            actualBefore = ConvertTwipsToPoints(explicitSpacing.Before.Value);
                        }
                        else if (styleSpacing?.Before?.Value != null)
                        {
                            actualBefore = ConvertTwipsToPoints(styleSpacing.Before.Value);
                        }

                        if (gost.AdditionalHeaderLineSpacingBefore.HasValue && Math.Abs(actualBefore - gost.AdditionalHeaderLineSpacingBefore.Value) > 0.1)
                        {
                            errorDetails.Add($"интервал перед: {actualBefore:F1} pt! (требуется {gost.AdditionalHeaderLineSpacingBefore.Value:F1} pt)\n");
                            hasError = true;
                        }

                        if (explicitSpacing?.After?.Value != null)
                        {
                            actualAfter = ConvertTwipsToPoints(explicitSpacing.After.Value);
                        }
                        else if (styleSpacing?.After?.Value != null)
                        {
                            actualAfter = ConvertTwipsToPoints(styleSpacing.After.Value);
                        }

                        if (gost.AdditionalHeaderLineSpacingAfter.HasValue && Math.Abs(actualAfter - gost.AdditionalHeaderLineSpacingAfter.Value) > 0.1)
                        {
                            errorDetails.Add($"интервал после: {actualAfter:F1} pt! (требуется {gost.AdditionalHeaderLineSpacingAfter.Value:F1} pt)\n");
                            hasError = true;
                        }
                    }

                    // Проверка отступов
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    var styleIndent = style?.StyleParagraphProperties?.Indentation;

                    // Преобразуем все значения в сантиметры
                    double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) :
                                     styleIndent?.Left?.Value != null ? TwipsToCm(double.Parse(styleIndent.Left.Value)) : 0;

                    double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) :
                                          styleIndent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(styleIndent.FirstLine.Value)) : 0;

                    double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) :
                                        styleIndent?.Hanging?.Value != null ? TwipsToCm(double.Parse(styleIndent.Hanging.Value)) : 0;

                    // 1. ПРОВЕРКА ВИЗУАЛЬНОГО ЛЕВОГО ОТСТУПА ЗАГОЛОВКА
                    if (gost.AdditionalHeaderIndentLeft.HasValue)
                    {
                        double actualTextIndent = leftIndent; // Базовый отступ

                        // Корректировка если есть выступ (hanging)
                        if (hangingIndent > 0)
                        {
                            actualTextIndent = leftIndent - hangingIndent;
                        }

                        if (Math.Abs(actualTextIndent - gost.AdditionalHeaderIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"Левый отступ заголовка: {actualTextIndent:F2} см (требуется {gost.AdditionalHeaderIndentLeft.Value:F2} см)\n");
                            hasError = true;
                        }
                    }

                    // 2. ПРОВЕРКА ПРАВОГО ОТСТУПА ЗАГОЛОВКА
                    if (gost.AdditionalHeaderIndentRight.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) :
                                          styleIndent?.Right?.Value != null ? TwipsToCm(double.Parse(styleIndent.Right.Value)) :
                                          DefaultAdditionalHeaderIndentRight;

                        if (Math.Abs(actualRight - gost.AdditionalHeaderIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"Правый отступ: {actualRight:F2} см (требуется {gost.AdditionalHeaderIndentRight.Value:F2} см)\n");
                            hasError = true;
                        }
                    }

                    // 3. ПРОВЕРКА ПЕРВОЙ СТРОКИ ЗАГОЛОВКА
                    if (gost.AdditionalHeaderFirstLineIndent.HasValue)
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

                        // Проверка типа (выступ/отступ)
                        if (!string.IsNullOrEmpty(gost.AdditionalHeaderIndentOrOutdent))
                        {
                            bool typeError = false;

                            if (gost.AdditionalHeaderIndentOrOutdent == "Выступ" && !isHanging)
                                typeError = true;
                            else if (gost.AdditionalHeaderIndentOrOutdent == "Отступ" && !isFirstLine)
                                typeError = true;
                            else if (gost.AdditionalHeaderIndentOrOutdent == "Нет" && (isHanging || isFirstLine))
                                typeError = true;

                            if (typeError)
                            {
                                errorDetails.Add($"Тип первой строки: {(isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет")} (требуется {gost.AdditionalHeaderIndentOrOutdent})\n");
                                hasError = true;
                            }
                        }

                        // Проверка значения
                        double currentValue = isHanging ? hangingIndent : firstLineIndent;
                        if ((isHanging || isFirstLine) && Math.Abs(currentValue - gost.AdditionalHeaderFirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{(isHanging ? "Выступ" : "Отступ")} первой строки: {currentValue:F2} см (требуется {gost.AdditionalHeaderFirstLineIndent.Value:F2} см)\n");
                            hasError = true;
                        }
                        else if (gost.AdditionalHeaderIndentOrOutdent != "Нет" && !isHanging && !isFirstLine)
                        {
                            errorDetails.Add($"Отсутствует {gost.AdditionalHeaderIndentOrOutdent} первой строки\n");
                            hasError = true;
                        }
                    }


                    if (hasError)
                    {
                        string shortText = GetShortText2(paragraph.InnerText?.Trim() ?? "");
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"Заголовок '{shortText}': {string.Join("        - ", errorDetails)}",
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
                        string errorMessage = "Ошибки в дополнительных заголовках:\n";
                        errorMessage += string.Join("\n", errors.Select(error => $"  • {error.ErrorMessage}"));
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
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twips"></param>
        /// <returns></returns>
        private double TwipsToCm(double twips) => twips / 567.0;

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
