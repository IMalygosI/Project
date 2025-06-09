using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avalonia.Media;
using Avalonia.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок оглавления
    /// </summary>
    public class CheckOglavleniya
    {
        private readonly WordprocessingDocument _wordDoc;
        private readonly Gost _gost;
        private readonly Func<Paragraph, bool> _isTocParagraph;
        private readonly Func<Paragraph, bool> _isEmptyParagraph;

        public CheckOglavleniya(WordprocessingDocument wordDoc, Gost gost, Func<Paragraph, bool> isTocParagraph, Func<Paragraph, bool> isEmptyParagraph)
        {
            _wordDoc = wordDoc;
            _gost = gost;
            _isTocParagraph = isTocParagraph;
            _isEmptyParagraph = isEmptyParagraph;
        }

        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckTocFormattingAsync(Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
                bool isValid = true;

                // 1. Проверка, требуется ли оглавление по ГОСТу
                if (_gost.RequireTOC.HasValue && !_gost.RequireTOC.Value)
                {
                    updateUI?.Invoke("Оглавление не требуется по ГОСТу", Brushes.Gray);
                    return (true, tempErrors);
                }

                // 2. Поиск оглавления
                var body = _wordDoc.MainDocumentPart.Document.Body;
                var tocField = body.Descendants<FieldCode>().FirstOrDefault(f => f.Text.Contains(" TOC ") || f.Text.Contains("TOC \\"));
                var tocParagraphs = body.Descendants<Paragraph>().Where(_isTocParagraph).ToList();

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
                    updateUI?.Invoke("Автоматическое оглавление не найдено! Создайте через 'Ссылки → Оглавление'", Brushes.Red);
                    tempErrors.Add(new TextErrorInfo
                    {
                        ErrorMessage = "Автоматическое оглавление не найдено",
                        ProblemRun = null,
                        ProblemParagraph = null
                    });
                    return (false, tempErrors);
                }

                // 5. Получаем все стили документа
                var stylesPart = _wordDoc.MainDocumentPart.StyleDefinitionsPart;
                var allStyles = stylesPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                // 6. Проверяем каждый параграф оглавления
                foreach (var paragraph in tocContainer.Descendants<Paragraph>())
                {
                    if (_isEmptyParagraph(paragraph)) continue;

                    bool paragraphHasError = false;
                    var errorDetails = new List<string>();

                    var pPr = paragraph.ParagraphProperties;
                    var spacing = pPr?.SpacingBetweenLines;
                    var justification = pPr?.Justification;
                    var indent = pPr?.Indentation;

                    Style paragraphStyle = null;
                    var styleId = pPr?.ParagraphStyleId?.Val?.Value;
                    if (styleId != null && allStyles.TryGetValue(styleId, out var style))
                    {
                        paragraphStyle = GetEffectiveStyle(style, allStyles);
                    }

                    // Проверка шрифта для каждого Run в параграфе
                    if (!string.IsNullOrEmpty(_gost.TocFontName))
                    {
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            if (run.Descendants<FieldChar>().Any() || run.InnerText.Trim().StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase)
                                                                   || run.InnerText.Trim().StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase) 
                                                                   || run.InnerText.Trim().StartsWith("MERGEFORMAT", StringComparison.OrdinalIgnoreCase))
                                continue;

                            var (actualFont, isFontDefined) = GetActualFontForToc(run, paragraph);

                            if (!isFontDefined)
                            {
                                errorDetails.Add($"\n       • не удалось определить шрифт");
                                paragraphHasError = true;
                                break;
                            }
                            else if (!string.Equals(actualFont, _gost.TocFontName, StringComparison.OrdinalIgnoreCase))
                            {
                                errorDetails.Add($"\n       • шрифт: '{actualFont}' (требуется '{_gost.TocFontName}')");
                                paragraphHasError = true;
                                break;
                            }
                        }

                    }

                    // Проверка размера шрифта
                    if (_gost.TocFontSize.HasValue)
                    {
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            if (run.Descendants<FieldChar>().Any()
                                || run.InnerText.Trim().StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase)
                                || run.InnerText.Trim().StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase)
                                || run.InnerText.Trim().StartsWith("MERGEFORMAT", StringComparison.OrdinalIgnoreCase))
                                continue;

                            var (actualSize, isSizeDefined) = GetActualFontSize(run, paragraph);

                            if (!isSizeDefined)
                            {
                                errorDetails.Add($"\n       • не удалось определить размер шрифта");
                                paragraphHasError = true;
                                break;
                            }
                            else if (Math.Abs(actualSize - _gost.TocFontSize.Value) > 0.1)
                            {
                                errorDetails.Add($"\n       • размер шрифта: {actualSize:F1} pt (требуется {_gost.TocFontSize.Value:F1} pt)");
                                paragraphHasError = true;
                                break;
                            }
                        }
                    }

                    // Проверка выравнивания
                    if (!string.IsNullOrEmpty(_gost.TocAlignment))
                    {
                        var (actualAlignment, isAlignmentDefined) = GetActualAlignment(paragraph, paragraphStyle);

                        if (!isAlignmentDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить выравнивание");
                            paragraphHasError = true;
                        }
                        else if (actualAlignment != _gost.TocAlignment)
                        {
                            errorDetails.Add($"\n       • выравнивание: '{actualAlignment}' (требуется '{_gost.TocAlignment}')");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка междустрочного интервала
                    if (_gost.TocLineSpacing.HasValue || !string.IsNullOrEmpty(_gost.TocLineSpacingType))
                    {
                        var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacing(paragraph, paragraphStyle);

                        if (!isSpacingDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить междустрочный интервал");
                            paragraphHasError = true;
                        }
                        else
                        {
                            // Проверка типа интервала
                            if (!string.IsNullOrEmpty(_gost.TocLineSpacingType))
                            {
                                string requiredType = _gost.TocLineSpacingType;
                                if (actualSpacingType != requiredType)
                                {
                                    errorDetails.Add($"\n       • тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                                    paragraphHasError = true;
                                }
                            }

                            // Проверка значения интервала
                            if (_gost.TocLineSpacing.HasValue && Math.Abs(actualSpacingValue - _gost.TocLineSpacing.Value) > 0.01)
                            {
                                errorDetails.Add($"\n       • межстрочный интервал: {actualSpacingValue:F2} (требуется {_gost.TocLineSpacing.Value:F2})");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // Проверка интервалов перед/после
                    if (_gost.TocLineSpacingBefore.HasValue || _gost.TocLineSpacingAfter.HasValue)
                    {
                        var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacing(paragraph, paragraphStyle);

                        if (!isSpacingDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить интервалы перед/после");
                            paragraphHasError = true;
                        }
                        else
                        {
                            if (_gost.TocLineSpacingBefore.HasValue &&
                                Math.Abs(actualBefore - _gost.TocLineSpacingBefore.Value) > 0.01)
                            {
                                errorDetails.Add($"\n       • интервал перед: {actualBefore:F1} pt (требуется {_gost.TocLineSpacingBefore.Value:F1} pt)");
                                paragraphHasError = true;
                            }

                            if (_gost.TocLineSpacingAfter.HasValue &&
                                Math.Abs(actualAfter - _gost.TocLineSpacingAfter.Value) > 0.01)
                            {
                                errorDetails.Add($"\n       • интервал после: {actualAfter:F1} pt (требуется {_gost.TocLineSpacingAfter.Value:F1} pt)");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // Проверка отступов
                    if (_gost.TocIndentLeft.HasValue || _gost.TocIndentRight.HasValue || _gost.TocFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TocIndentOrOutdent))
                    {
                        var (leftIndent, rightIndent, firstLineIndent, firstLineType, isIndentDefined) = GetActualIndents(paragraph, paragraphStyle);

                        // Если отступы не требуются по ГОСТу, пропускаем проверку
                        bool needCheckIndents = _gost.TocIndentLeft.HasValue || _gost.TocIndentRight.HasValue || _gost.TocFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TocIndentOrOutdent);

                        if (needCheckIndents)
                        {
                            // Проверка левого отступа
                            if (_gost.TocIndentLeft.HasValue && Math.Abs(leftIndent - _gost.TocIndentLeft.Value) > 0.05)
                            {
                                errorDetails.Add($"\n       • левый отступ: {leftIndent:F2} см (требуется {_gost.TocIndentLeft.Value:F2} см)");
                                paragraphHasError = true;
                            }

                            // Проверка правого отступа
                            if (_gost.TocIndentRight.HasValue && Math.Abs(rightIndent - _gost.TocIndentRight.Value) > 0.05)
                            {
                                errorDetails.Add($"\n       • правый отступ: {rightIndent:F2} см (требуется {_gost.TocIndentRight.Value:F2} см)");
                                paragraphHasError = true;
                            }

                            // Проверка первой строки
                            if (_gost.TocFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TocIndentOrOutdent))
                            {
                                if (!string.IsNullOrEmpty(_gost.TocIndentOrOutdent) && firstLineType != _gost.TocIndentOrOutdent)
                                {
                                    errorDetails.Add($"\n       • тип первой строки: '{firstLineType}' (требуется '{_gost.TocIndentOrOutdent}')");
                                    paragraphHasError = true;
                                }

                                if (_gost.TocFirstLineIndent.HasValue && firstLineType != "Нет" && Math.Abs(firstLineIndent - _gost.TocFirstLineIndent.Value) > 0.05)
                                {
                                    errorDetails.Add($"\n       • {firstLineType.ToLower()} первой строки: {firstLineIndent:F2} см (требуется {_gost.TocFirstLineIndent.Value:F2} см)");
                                    paragraphHasError = true;
                                }
                            }
                        }
                    }

                    if (paragraphHasError)
                    {
                        string shortText = GetShortTocText(paragraph);
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            tempErrors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"\nОшибка поля: '{shortText}' - {string.Join(", ", errorDetails)}",
                                ProblemParagraph = paragraph,
                                ProblemRun = run
                            });
                        }
                        isValid = false;
                    }
                }

                var distinctErrors = tempErrors.GroupBy(e => new { e.ErrorMessage, Paragraph = e.ProblemParagraph }).Select(g => g.First()).ToList();

                updateUI?.Invoke(!isValid ? "Ошибки в оглавлении:" + string.Join("\n", distinctErrors.Select(e => e.ErrorMessage).Take(30))
                                          : "Оглавление полностью соответствует ГОСТу", !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        private (string FontName, bool IsDefined) GetActualFontForToc(Run run, Paragraph paragraph)
        {
            // Пробуем явные свойства Run
            var explicitFont = GetExplicitRunFont(run);
            if (!string.IsNullOrEmpty(explicitFont))
                return (explicitFont, true);

            // Пробуем стиль Run
            var runStyleFont = GetStyleFont(run.RunProperties?.RunStyle?.Val?.Value);
            if (!string.IsNullOrEmpty(runStyleFont))
                return (runStyleFont, true);

            // Пробуем стиль параграфа
            var paraStyleFont = GetParagraphStyleFont(paragraph);
            if (!string.IsNullOrEmpty(paraStyleFont))
                return (paraStyleFont, true);

            return (null, false);
        }

        private string GetExplicitRunFont(Run run)
        {
            // Проверяем все возможные места, где может быть указан шрифт
            var runProps = run.RunProperties;

            if (runProps == null) return null;

            return runProps.RunFonts?.Ascii?.Value ?? runProps.RunFonts?.HighAnsi?.Value ?? 
                   runProps.RunFonts?.ComplexScript?.Value ?? runProps.RunFonts?.EastAsia?.Value;
        }

        private string GetStyleFont(Style style)
        {
            if (style?.StyleRunProperties == null) return null;

            return style.StyleRunProperties.RunFonts?.Ascii?.Value ?? style.StyleRunProperties.RunFonts?.HighAnsi?.Value ?? 
                   style.StyleRunProperties.RunFonts?.ComplexScript?.Value ?? style.StyleRunProperties.RunFonts?.EastAsia?.Value;
        }

        private string GetStyleFont(string styleId)
        {
            if (string.IsNullOrEmpty(styleId)) return null;
            return GetStyleFont(GetStyleById(styleId));
        }

        private string GetParagraphStyleFont(Paragraph paragraph)
        {
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

            if (string.IsNullOrEmpty(styleId)) return null;

            // Проверяем всю цепочку наследования стилей
            var currentStyle = GetStyleById(styleId);
            while (currentStyle != null)
            {
                var font = GetStyleFont(currentStyle.StyleId.Value);

                if (!string.IsNullOrEmpty(font))
                    return font;

                currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
            }

            return null;
        }

        private (double Size, bool IsDefined) GetActualFontSize(Run run, Paragraph paragraph)
        {
            // 1. Проверяем явные свойства Run
            var runSize = run.RunProperties?.FontSize?.Val?.Value;

            if (runSize != null)
                return (double.Parse(runSize) / 2, true);

            // 2. Проверяем стиль Run
            var runStyleId = run.RunProperties?.RunStyle?.Val?.Value;

            if (runStyleId != null)
            {
                var runStyle = GetStyleById(runStyleId);
                if (runStyle?.StyleRunProperties?.FontSize?.Val?.Value != null)
                    return (double.Parse(runStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);
            }

            // 3. Проверяем стиль Paragraph
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

            if (paraStyleId != null)
            {
                var currentStyle = GetStyleById(paraStyleId);
                while (currentStyle != null)
                {
                    if (currentStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
                        return (double.Parse(currentStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
                }
            }

            // 4. Проверяем Normal стиль
            var normalStyle = GetStyleById("Normal");

            if (normalStyle?.StyleRunProperties?.FontSize?.Val?.Value != null)
                return (double.Parse(normalStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);

            return (0, false);
        }

        private (string Alignment, bool IsDefined) GetActualAlignment(Paragraph paragraph, Style paragraphStyle)
        {
            // 1. Проверяем явное выравнивание в параграфе
            if (paragraph.ParagraphProperties?.Justification != null)
            {
                return (GetAlignmentString(paragraph.ParagraphProperties.Justification), true);
            }

            // 2. Проверяем стиль параграфа и его родителей
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null)
            {
                var currentStyle = GetStyleById(paraStyleId);
                while (currentStyle != null)
                {
                    if (currentStyle.StyleParagraphProperties?.Justification != null)
                    {
                        return (GetAlignmentString(currentStyle.StyleParagraphProperties.Justification), true);
                    }
                    currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
                }
            }

            // 3. Проверяем стиль Normal
            var normalStyle = GetStyleById("Normal");
            if (normalStyle?.StyleParagraphProperties?.Justification != null)
            {
                return (GetAlignmentString(normalStyle.StyleParagraphProperties.Justification), true);
            }

            // 4. Если нигде не задано - считаем Left по умолчанию
            return ("Left", true);
        }

        private (string Type, double Value, bool IsDefined) GetActualLineSpacing(Paragraph paragraph, Style paragraphStyle)
        {
            // 1. Явные свойства
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;

            var parsed = ParseLineSpacing(spacing);
            if (parsed.IsDefined)
                return parsed;

            // 2. Стиль абзаца
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null)
            {
                var currentStyle = GetStyleById(paraStyleId);
                while (currentStyle != null)
                {
                    var styleSpacing = currentStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    parsed = ParseLineSpacing(styleSpacing);
                    if (parsed.IsDefined)
                        return parsed;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
                }
            }

            // 3. Normal
            var normalStyle = GetStyleById("Normal");
            parsed = ParseLineSpacing(normalStyle?.StyleParagraphProperties?.SpacingBetweenLines);
            if (parsed.IsDefined)
                return parsed;

            return ("Множитель", 1.0, true);
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

                // По умолчанию
                return ("Множитель", lineValue / 240.0, true);
            }

            return (null, 0, false);
        }

        private (double Before, double After, bool IsDefined) GetActualParagraphSpacing(Paragraph paragraph, Style paragraphStyle)
        {
            double? before = null;
            double? after = null;

            // 1. Явные свойства
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing?.Before?.Value != null) before = ConvertTwipsToPoints(spacing.Before.Value);
            if (spacing?.After?.Value != null) after = ConvertTwipsToPoints(spacing.After.Value);

            // 2. Поиск в стилях по BasedOn цепочке
            if (before == null || after == null)
            {
                var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                var currentStyle = paraStyleId != null ? GetStyleById(paraStyleId) : paragraphStyle;

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

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
                }
            }

            // 3. Fallback: Normal
            if (before == null || after == null)
            {
                var normalStyle = GetStyleById("Normal");
                var spacingNorm = normalStyle?.StyleParagraphProperties?.SpacingBetweenLines;

                if (spacingNorm != null)
                {
                    if (before == null && spacingNorm.Before?.Value != null)
                        before = ConvertTwipsToPoints(spacingNorm.Before.Value);

                    if (after == null && spacingNorm.After?.Value != null)
                        after = ConvertTwipsToPoints(spacingNorm.After.Value);
                }
            }

            var isDefined = before.HasValue || after.HasValue;
            return (before ?? 0, after ?? 0.1, isDefined);
        }

        private (double Left, double Right, double FirstLine, string FirstLineType, bool IsDefined) GetActualIndents(Paragraph paragraph, Style paragraphStyle)
        {
            // Устанавливаем значения по умолчанию
            double left = 0.0;
            double right = 0.0;
            double firstLine = 0.0;
            string firstLineType = "Нет";
            bool isDefined = false;

            // 1. Проверяем явные свойства параграфа
            var indent = paragraph.ParagraphProperties?.Indentation;
            if (indent != null)
            {
                if (indent.Left?.Value != null)
                {
                    left = ConvertTwipsToCm(indent.Left.Value);
                    isDefined = true;
                }

                if (indent.Right?.Value != null)
                {
                    right = ConvertTwipsToCm(indent.Right.Value);
                    isDefined = true;
                }

                if (indent.FirstLine?.Value != null && int.TryParse(indent.FirstLine.Value, out int firstLineValue) && firstLineValue != 0)
                {
                    firstLine = ConvertTwipsToCm(indent.FirstLine.Value);
                    firstLineType = "Отступ";
                    isDefined = true;
                }
                else if (indent.Hanging?.Value != null && int.TryParse(indent.Hanging.Value, out int hangingValue) && hangingValue != 0)
                {
                    firstLine = ConvertTwipsToCm(indent.Hanging.Value);
                    firstLineType = "Выступ";
                    isDefined = true;
                }
            }

            // 2. Проверяем стили (если для каких-то параметров еще не нашли значения)
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var currentStyle = paraStyleId != null ? GetStyleById(paraStyleId) : paragraphStyle;

            while (currentStyle != null)
            {
                var styleIndent = currentStyle.StyleParagraphProperties?.Indentation;
                if (styleIndent != null)
                {
                    if (styleIndent.Left?.Value != null && !isDefined)
                    {
                        left = ConvertTwipsToCm(styleIndent.Left.Value);
                        isDefined = true;
                    }

                    if (styleIndent.Right?.Value != null && !isDefined)
                    {
                        right = ConvertTwipsToCm(styleIndent.Right.Value);
                        isDefined = true;
                    }

                    if (styleIndent.FirstLine?.Value != null && !isDefined)
                    {
                        firstLine = ConvertTwipsToCm(styleIndent.FirstLine.Value);
                        firstLineType = "Отступ";
                        isDefined = true;
                    }
                    else if (styleIndent.Hanging?.Value != null && !isDefined)
                    {
                        firstLine = ConvertTwipsToCm(styleIndent.Hanging.Value);
                        firstLineType = "Выступ";
                        isDefined = true;
                    }
                }
                currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
            }

            return (left, right, firstLine, firstLineType, isDefined);
        }

        /// <summary>
        /// Конвертирует строковое значение twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twipsValue">Значение в twips (строка)</param>
        /// <returns>Значение в сантиметрах</returns>
        private double ConvertTwipsToCm(string twipsValue)
        {
            if (string.IsNullOrEmpty(twipsValue))
                return 0;

            if (double.TryParse(twipsValue, out double twips))
                return twips / 567.0;

            return 0;
        }

        private void ApplyInheritedStyles(Style source, Style target)
        {
            if (source.BasedOn?.Val?.Value != null)
            {
                var parentStyle = GetStyleById(source.BasedOn.Val.Value);
                if (parentStyle != null)
                {
                    ApplyInheritedStyles(parentStyle, target);
                }
            }

            // Копируем свойства, если они не определены в target
            if (source.StyleRunProperties != null)
            {
                target.StyleRunProperties ??= new StyleRunProperties();
                MergeProperties(source.StyleRunProperties, target.StyleRunProperties);
            }

            if (source.StyleParagraphProperties != null)
            {
                target.StyleParagraphProperties ??= new StyleParagraphProperties();
                MergeProperties(source.StyleParagraphProperties, target.StyleParagraphProperties);
            }
        }

        private void MergeProperties(OpenXmlElement source, OpenXmlElement target)
        {
            foreach (var sourceProperty in source.Elements())
            {
                var targetProperty = target.Elements().FirstOrDefault(e => e.GetType() == sourceProperty.GetType());
                if (targetProperty == null)
                {
                    target.AppendChild((OpenXmlElement)sourceProperty.Clone());
                }
            }
        }

        private Style GetEffectiveStyle(Style style)
        {
            if (style == null) return null;
            if (style.BasedOn?.Val?.Value == null) return style;

            var baseStyle = GetStyleById(style.BasedOn.Val.Value);
            return baseStyle != null ? GetEffectiveStyle(baseStyle) : style;
        }

        private Style GetStyleById(string styleId)
        {
            return _wordDoc.MainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId.Value == styleId);
        }

        /// <summary>
        /// Рекурсивная проверка наследования стилей
        /// </summary>
        /// <param name="style"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private Style GetEffectiveStyle(Style style, Dictionary<string, Style> allStyles)
        {
            if (style == null) return null;

            if (style.BasedOn?.Val?.Value != null && allStyles.TryGetValue(style.BasedOn.Val.Value, out var parentStyle))
            {
                return GetEffectiveStyle(parentStyle, allStyles);
            }

            return style;
        }

        /// <summary>
        /// Обрезает текст элемента оглавления до 30 символов
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private string GetShortTocText(Paragraph paragraph)
        {
            string text = paragraph.InnerText.Trim();
            return text.Length > 50 ? text.Substring(0, Math.Min(47, text.Length)) + "..." : text;
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
    }
}