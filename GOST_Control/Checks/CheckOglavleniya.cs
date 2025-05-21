using System;
using System.Collections.Generic;
using System.Linq;
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
        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ОГЛАВЛЕНИЯ =======================
        private const string DefaultTocFont = "Arial";
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

        /// <summary>
        /// Комплексная проверка форматирования оглавления
        /// </summary>
        /// <param name="updateUI"></param>
        /// <returns></returns>
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

                // 5. Получаем стили оглавления
                var tocStyles = _wordDoc.MainDocumentPart?.StyleDefinitionsPart?.Styles.Elements<Style>().Where(s => s.StyleId?.Value?.StartsWith("TOC") == true).ToDictionary(s => s.StyleId.Value);

                // 6. Проверяем каждый параграф оглавления
                foreach (var paragraph in tocContainer.Descendants<Paragraph>())
                {
                    if (_isEmptyParagraph(paragraph)) continue;

                    bool paragraphHasError = false;
                    var errorDetails = new List<string>();

                    var indent = paragraph.ParagraphProperties?.Indentation;
                    var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    var justification = paragraph.ParagraphProperties?.Justification;

                    // Получаем стиль параграфа
                    Style paragraphStyle = null;
                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    if (styleId != null && tocStyles.TryGetValue(styleId, out var style))
                    {
                        paragraphStyle = style;
                    }

                    // Проверка шрифта
                    if (!string.IsNullOrEmpty(_gost.TocFontName))
                    {
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? paragraphStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value;

                            if (font != null && !string.Equals(font, _gost.TocFontName, StringComparison.OrdinalIgnoreCase))
                            {
                                errorDetails.Add($"шрифт: '{font}' (требуется '{_gost.TocFontName}')");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // Проверка размера шрифта
                    if (_gost.TocFontSize.HasValue)
                    {
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value ?? paragraphStyle?.StyleRunProperties?.FontSize?.Val?.Value;

                            double actualSize = fontSizeVal != null ? double.Parse(fontSizeVal) / 2 : DefaultTocSize;

                            if (Math.Abs(actualSize - _gost.TocFontSize.Value) > 0.1)
                            {
                                errorDetails.Add($"размер шрифта: {actualSize:F1} pt (требуется {_gost.TocFontSize.Value:F1} pt)");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // Проверка выравнивания
                    if (!string.IsNullOrEmpty(_gost.TocAlignment))
                    {
                        var actualAlignment = GetAlignmentString(justification);
                        if (actualAlignment != _gost.TocAlignment)
                        {
                            errorDetails.Add($"выравнивание: '{actualAlignment}' (требуется '{_gost.TocAlignment}')");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка отступов

                    // Преобразуем все значения в сантиметры
                    double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                    double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                    double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

                    // Проверка левого отступа
                    if (_gost.TocIndentLeft.HasValue)
                    {
                        double actualTextIndent = leftIndent;
                        if (hangingIndent > 0)
                        {
                            actualTextIndent = leftIndent - hangingIndent;
                        }

                        if (Math.Abs(actualTextIndent - _gost.TocIndentLeft.Value) > 0.05)
                        {
                            errorDetails.Add($"левый отступ: {actualTextIndent:F2} см (требуется {_gost.TocIndentLeft.Value:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка правого отступа
                    if (_gost.TocIndentRight.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTocRightIndent;

                        if (Math.Abs(actualRight - _gost.TocIndentRight.Value) > 0.05)
                        {
                            errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {_gost.TocIndentRight.Value:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка первой строки
                    if (_gost.TocFirstLineIndent.HasValue)
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

                        if (!string.IsNullOrEmpty(_gost.TocIndentOrOutdent))
                        {
                            bool typeError = false;
                            string errorType = string.Empty;

                            if (_gost.TocIndentOrOutdent == "Выступ" && !isHanging)
                            {
                                typeError = true;
                                errorType = "тип первой строки: 'Нет' (требуется 'Выступ')";
                            }
                            else if (_gost.TocIndentOrOutdent == "Отступ" && !isFirstLine)
                            {
                                typeError = true;
                                errorType = "тип первой строки: 'Нет' (требуется 'Отступ')";
                            }
                            else if (_gost.TocIndentOrOutdent == "Нет" && (isHanging || isFirstLine))
                            {
                                typeError = true;
                                errorType = $"тип первой строки: '{(isHanging ? "Выступ" : "Отступ")}' (требуется 'Нет')";
                            }

                            if (typeError)
                            {
                                errorDetails.Add(errorType);
                                paragraphHasError = true;
                            }
                        }

                        double currentValue = isHanging ? hangingIndent : firstLineIndent;
                        if ((isHanging || isFirstLine) && Math.Abs(currentValue - _gost.TocFirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{(isHanging ? "выступ" : "отступ")} первой строки: {currentValue:F2} см (требуется {_gost.TocFirstLineIndent.Value:F2} см)");
                            paragraphHasError = true;
                        }
                        else if (_gost.TocIndentOrOutdent != "Нет" && !isHanging && !isFirstLine)
                        {
                            errorDetails.Add($"отсутствует {_gost.TocIndentOrOutdent.ToLower()} первой строки");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка межстрочного интервала
                    double actualSpacing = DefaultTocLineSpacingValue;
                    string actualSpacingType = DefaultTocLineSpacingType;
                    LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                    if (spacing?.Line != null)
                    {
                        if (spacing.LineRule?.Value == LineSpacingRuleValues.Exact)
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 567.0;
                            actualSpacingType = "точно";
                            actualRule = LineSpacingRuleValues.Exact;
                        }
                        else if (spacing.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 567.0;
                            actualSpacingType = "минимум";
                            actualRule = LineSpacingRuleValues.AtLeast;
                        }
                        else
                        {
                            actualSpacing = double.Parse(spacing.Line.Value) / 240.0;
                            actualSpacingType = "множитель";
                            actualRule = LineSpacingRuleValues.Auto;
                        }
                    }

                    // Проверка типа интервала
                    LineSpacingRuleValues requiredRule = (_gost.TocLineSpacingType ?? DefaultTocLineSpacingType) switch
                    {
                        "Минимум" => LineSpacingRuleValues.AtLeast,
                        "Точно" => LineSpacingRuleValues.Exact,
                        _ => LineSpacingRuleValues.Auto
                    };

                    if (actualRule != requiredRule)
                    {
                        errorDetails.Add($"тип интервала: '{actualSpacingType}' (требуется '{_gost.TocLineSpacingType ?? DefaultTocLineSpacingType.ToLower()}')");
                        paragraphHasError = true;
                    }

                    // Проверка значения интервала
                    if (Math.Abs(actualSpacing - _gost.TocLineSpacing.Value) > 0.01)
                    {
                        errorDetails.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {_gost.TocLineSpacing.Value:F2})");
                        paragraphHasError = true;
                    }


                    // Проверка интервалов перед/после
                    if (_gost.TocLineSpacingBefore.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultTocSpacingBefore;

                        if (Math.Abs(actualBefore - _gost.TocLineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {_gost.TocLineSpacingBefore.Value:F1} pt)");
                            paragraphHasError = true;
                        }
                    }

                    if (_gost.TocLineSpacingAfter.HasValue)
                    {
                        double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultTocSpacingAfter;

                        if (Math.Abs(actualAfter - _gost.TocLineSpacingAfter.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал после: {actualAfter:F1} pt (требуется {_gost.TocLineSpacingAfter.Value:F1} pt)");
                            paragraphHasError = true;
                        }
                    }


                    if (paragraphHasError)
                    {
                        string shortText = GetShortTocText(paragraph);

                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"\nОшибка поля: '{shortText}' - {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });

                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            if (string.IsNullOrWhiteSpace(run.InnerText)) continue;

                            tempErrors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"       • Ошибка в: {run.InnerText}",
                                ProblemParagraph = paragraph,
                                ProblemRun = run
                            });
                        }

                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки в оглавлении:" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(30)) : "Оглавление полностью соответствует ГОСТу",
                                 !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twips"></param>
        /// <returns></returns>
        private double TwipsToCm(double twips) => twips / 567.0;

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