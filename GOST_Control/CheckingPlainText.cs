using Avalonia.Media;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GOST_Control
{
    /// <summary>
    /// Класс для проверки чистого текста
    /// </summary>
    public class CheckingPlainText
    {
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

        private readonly Func<Paragraph, bool> _shouldSkipParagraph;

        public CheckingPlainText(Func<Paragraph, bool> shouldSkipParagraph)
        {
            _shouldSkipParagraph = shouldSkipParagraph;
        }

        /// <summary>
        /// Метод проверки интервалов между абзацами (асинхронная версия)
        /// </summary>
        public async Task<(bool IsValid, List<string> Errors)> CheckParagraphSpacingAsync(bool hasBeforeSpacing, bool hasAfterSpacing, List<Paragraph> paragraphs, Gost gost, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                if (!hasBeforeSpacing && !hasAfterSpacing)
                {
                    updateUI?.Invoke("Интервалы между абзацами не требуются", Brushes.Gray);
                    return (true, tempErrors);
                }

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    bool hasError = false;
                    var errorDetails = new List<string>();

                    if (hasBeforeSpacing && gost.LineSpacingBefore.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultTextSpacingBefore;

                        if (Math.Abs(actualBefore - gost.LineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.LineSpacingBefore.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    if (hasAfterSpacing && gost.LineSpacingAfter.HasValue)
                    {
                        double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultTextSpacingAfter;

                        if (Math.Abs(actualAfter - gost.LineSpacingAfter.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал после: {actualAfter:F1} pt (требуется {gost.LineSpacingAfter.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string shortText = GetShortText(paragraph);
                        tempErrors.Add($"'{shortText}' - {string.Join(", ", errorDetails)}");
                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки в интервалах между абзацами:\n" + string.Join("\n", tempErrors.Take(3)) + (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "")
                                 : "Интервалы между абзацами соответствуют ГОСТу", !isValid ? Brushes.Red : Brushes.Green
                );

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки отступов первой строки (асинхронная версия)
        /// </summary>
        public async Task<(bool IsValid, List<string> Errors)> CheckFirstLineIndentAsync(double requiredFirstLineIndent, List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                if (gost.TextIndentOrOutdent == "Нет")
                {
                    updateUI?.Invoke("Отступ первой строки не требуется", Brushes.Gray);
                    return (true, tempErrors);
                }

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var indent = paragraph.ParagraphProperties?.Indentation;
                    bool hasError = false;
                    var errorDetails = new List<string>();

                    if (gost.IndentLeftText.HasValue)
                    {
                        double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultTextLeftIndent;

                        if (Math.Abs(actualLeft - gost.IndentLeftText.Value) > 0.05)
                        {
                            errorDetails.Add($"левый отступ: {actualLeft:F2} см (требуется {gost.IndentLeftText.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (gost.IndentRightText.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTextRightIndent;

                        if (Math.Abs(actualRight - gost.IndentRightText.Value) > 0.05)
                        {
                            errorDetails.Add($"правый отступ: {actualRight:F2} см (требуется {gost.IndentRightText.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (gost.TextIndentOrOutdent != "Нет")
                    {
                        string currentType = "Нет";
                        double? currentValue = null;

                        if (indent?.Hanging != null)
                        {
                            currentType = "Выступ";
                            currentValue = TwipsToCm(double.Parse(indent.Hanging.Value));
                        }
                        else if (indent?.FirstLine != null)
                        {
                            currentType = "Отступ";
                            currentValue = TwipsToCm(double.Parse(indent.FirstLine.Value));
                        }

                        string requiredType = gost.TextIndentOrOutdent == "Выступ" ? "Выступ" : "Отступ";

                        if (currentType != requiredType)
                        {
                            errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                            hasError = true;
                        }

                        if (currentValue.HasValue)
                        {
                            if (Math.Abs(currentValue.Value - requiredFirstLineIndent) > 0.05)
                            {
                                errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {requiredFirstLineIndent:F2} см)");
                                hasError = true;
                            }
                        }
                        else
                        {
                            errorDetails.Add($"Отсутствует {gost.TextIndentOrOutdent} первой строки");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string shortText = GetShortText(paragraph);
                        tempErrors.Add($"'{shortText}' - {string.Join(", ", errorDetails)}");
                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки в отступах:\n" + string.Join("\n", tempErrors.Take(3)) + (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "")
                                 : "Отступы соответствуют ГОСТу", !isValid ? Brushes.Red : Brushes.Green
                );

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки межстрочного интервала для простого текста (асинхронная версия)
        /// </summary>
        public async Task<(bool IsValid, List<string> Errors)> CheckLineSpacingAsync(double requiredLineSpacing, string requiredLineSpacingType, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;
                string requiredSpacingType = requiredLineSpacingType;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    foreach (var run in paragraph.Elements<Run>())
                    {
                        if (ShouldSkipRun(run))
                            continue;

                        var paragraphProperties = paragraph.ParagraphProperties;
                        var spacing = paragraphProperties?.SpacingBetweenLines;

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

                        if (actualSpacingType != requiredSpacingType)
                        {
                            var runText = run.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(runText))
                            {
                                tempErrors.Add($"{runText} (тип интервала: '{actualSpacingType}') - требуется ('{requiredSpacingType}')");
                                isValid = false;
                            }
                        }

                        if (Math.Abs(actualSpacing - requiredLineSpacing) > 0.01)
                        {
                            var runText = run.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(runText))
                            {
                                tempErrors.Add($"{runText} (интервал: {actualSpacing:F2}) - требуется ({requiredLineSpacing:F2})");
                                isValid = false;
                            }
                        }
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки в межстрочном интервале:\n" + string.Join("\n", tempErrors.Take(3)) + (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "")
                                        : "Межстрочный интервал соответствует ГОСТу",
                                !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки выравнивания текста (асинхронная версия)
        /// </summary>
        public async Task<(bool IsValid, List<string> Errors)> CheckTextAlignmentAsync(string requiredAlignment, List<Paragraph> paragraphs, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification) ?? DefaultTextAlignment;

                    if (currentAlignment != requiredAlignment)
                    {
                        var shortText = GetShortText(paragraph);
                        tempErrors.Add($"'{shortText}' (выравнивание: '{currentAlignment}', требуется: '{requiredAlignment}')");
                        isValid = false;
                    }
                }

                updateUI?.Invoke(isValid ? "Выравнивание простого текста соответствует ГОСТу" : "Ошибки в выравнивании текста: " + string.Join("\n", tempErrors), isValid ? Brushes.Green : Brushes.Red);
                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки размера шрифта (асинхронная версия)
        /// </summary>
        public async Task<(bool IsValid, List<string> Errors)> CheckFontSizeAsync(double requiredFontSize, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    foreach (var run in paragraph.Elements<Run>())
                    {
                        if (ShouldSkipRun(run)) continue;

                        var fontSize = run.RunProperties?.FontSize?.Val?.Value;
                        double actualSize = fontSize != null ? double.Parse(fontSize.ToString()) / 2 : DefaultTextSize;

                        if (Math.Abs(actualSize - requiredFontSize) > 0.1)
                        {
                            var runText = run.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(runText))
                            {
                                tempErrors.Add($"{runText} (размер: {actualSize:F1} pt) - требуется ({requiredFontSize:F1} pt)");
                                isValid = false;
                            }
                        }
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки в размере шрифта у простого текста:\n" + string.Join("\n", tempErrors.Take(3)) + (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") :
                                            "Размер шрифта простого текста соответствует ГОСТу", !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки шрифта (асинхронная версия)
        /// </summary>
        public async Task<(bool IsValid, List<string> Errors)> CheckFontNameAsync(string requiredFontName, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                var defaultStyle = GetDefaultStyle(doc);
                string defaultFont = defaultStyle?.StyleRunProperties?.RunFonts?.Ascii?.Value ?? DefaultTextFont;
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    foreach (var run in paragraph.Elements<Run>())
                    {
                        if (ShouldSkipRun(run)) continue;

                        var fontName = run.RunProperties?.RunFonts?.Ascii?.Value ?? defaultFont;

                        if (fontName != requiredFontName)
                        {
                            var shortText = GetShortText(paragraph);
                            tempErrors.Add($"'{shortText}' (шрифт: '{fontName}', требуется: '{requiredFontName}')");
                            isValid = false;
                        }
                    }
                }

                updateUI?.Invoke(isValid ? "Шрифт простого текста соответствует ГОСТу" : "Ошибки в шрифте: " + string.Join("\n", tempErrors), isValid ? Brushes.Green : Brushes.Red);
                return (isValid, tempErrors);
            });
        }

        // ======================= ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ =======================

        /// <summary>
        /// Вспомогательный метод который преобразует объект выравнивания в строковое представление
        /// </summary>
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
        /// Определяет тип межстрочного интервала
        /// </summary>
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
        private double CalculateActualSpacing(SpacingBetweenLines spacing)
        {
            if (spacing.Line == null) return 0;

            if (spacing.LineRule == LineSpacingRuleValues.Exact)
            {
                return double.Parse(spacing.Line.Value) / 20.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                return double.Parse(spacing.Line.Value) / 20.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                return double.Parse(spacing.Line.Value) / 240.0;
            }
            else
            {
                return double.Parse(spacing.Line.Value) / 240.0;
            }
        }

        /// <summary>
        /// Обрезает текст параграфа до 50 символов с добавлением многоточия
        /// </summary>
        private string GetShortText(Paragraph paragraph)
        {
            string text = paragraph.InnerText.Trim();
            return text.Length > 50 ? text.Substring(0, 47) + "..." : text;
        }

        /// <summary>
        /// Конвертирует строковое значение в twips в пункты 
        /// </summary>
        private double ConvertTwipsToPoints(string twipsValue)
        {
            if (string.IsNullOrEmpty(twipsValue))
                return 0;

            double value = double.Parse(twipsValue);
            return value / 567.0;
        }

        /// <summary>
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        private double TwipsToCm(double twips) => twips / 567.0;

        private Style GetDefaultStyle(WordprocessingDocument doc)
        {
            var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
            if (stylesPart == null) return null;

            return stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.Type == StyleValues.Paragraph && (s.Default?.Value ?? false));
        }

        /// <summary>
        /// Определяет нужно ли пропускать Run при проверке
        /// </summary>
        private bool ShouldSkipRun(Run run)
        {
            if (string.IsNullOrWhiteSpace(run.InnerText))
                return true;

            if (run.Elements<Break>().Any() || run.Elements<TabChar>().Any())
                return true;

            if (run.Descendants<Hyperlink>().Any())
                return true;

            return false;
        }
    }
}