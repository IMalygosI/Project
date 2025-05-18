using Avalonia.Media;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
        /// Метод проверки интервалов между абзацами
        /// </summary>
        /// <param name="hasBeforeSpacing"></param>
        /// <param name="hasAfterSpacing"></param>
        /// <param name="paragraphs"></param>
        /// <param name="gost"></param>
        /// <param name="doc"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckParagraphSpacingAsync(bool hasBeforeSpacing, bool hasAfterSpacing, List<Paragraph> paragraphs, Gost gost, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
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

                    // Проверка интервала перед абзацем
                    if (hasBeforeSpacing && gost.LineSpacingBefore.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultTextSpacingBefore;
                        if (Math.Abs(actualBefore - gost.LineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.LineSpacingBefore.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    // Проверка интервала после абзаца
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
                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"'{shortText}' - {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph, 
                            ProblemRun = null
                        });
                        isValid = false;
                    }

                }
                updateUI?.Invoke(!isValid ? "Ошибки в интервалах между абзацами:\n" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3)) : 
                                                                 "Интервалы между абзацами соответствуют ГОСТу", !isValid ? Brushes.Red : Brushes.Green);
                return (isValid, tempErrors);
            });
        }

        /// <summary>
        ///  Метод проверки отступов первой строки
        /// </summary>
        /// <param name="requiredFirstLineIndent"></param>
        /// <param name="paragraphs"></param>
        /// <param name="gost"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckFirstLineIndentAsync(double requiredFirstLineIndent, List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var indent = paragraph.ParagraphProperties?.Indentation;
                    bool hasError = false;
                    var errorDetails = new List<string>();

                    // Преобразуем все значения в сантиметры
                    double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                    double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                    double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

                    // 1. ПРОВЕРКА ВИЗУАЛЬНОГО ЛЕВОГО ОТСТУПА
                    if (gost.IndentLeftText.HasValue)
                    {
                        double actualTextIndent = leftIndent; // Базовый отступ

                        // Корректировка если есть выступ (hanging)
                        if (hangingIndent > 0)
                        {
                            actualTextIndent = leftIndent - hangingIndent;
                        }

                        if (Math.Abs(actualTextIndent - gost.IndentLeftText.Value) > 0.05)
                        {
                            errorDetails.Add($"Фактический левый отступ текста: {actualTextIndent:F2} см (требуется {gost.IndentLeftText.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // 2. ПРОВЕРКА ПЕРВОЙ СТРОКИ
                    if (gost.FirstLineIndent.HasValue)
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

                        // Проверка типа (выступ/отступ)
                        if (!string.IsNullOrEmpty(gost.TextIndentOrOutdent))
                        {
                            bool typeError = false;

                            if (gost.TextIndentOrOutdent == "Выступ" && !isHanging)
                                typeError = true;
                            else if (gost.TextIndentOrOutdent == "Отступ" && !isFirstLine)
                                typeError = true;
                            else if (gost.TextIndentOrOutdent == "Нет" && (isHanging || isFirstLine))
                                typeError = true;

                            if (typeError)
                            {
                                errorDetails.Add($"Тип первой строки: {(isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет")} (требуется {gost.TextIndentOrOutdent})");
                                hasError = true;
                            }
                        }

                        // Проверка значения
                        double currentValue = isHanging ? hangingIndent : firstLineIndent;
                        if ((isHanging || isFirstLine) && Math.Abs(currentValue - gost.FirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"{(isHanging ? "Выступ" : "Отступ")} первой строки: {currentValue:F2} см (требуется {gost.FirstLineIndent.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    // 3. ПРОВЕРКА ПРАВОГО ОТСТУПА (без изменений)
                    if (gost.IndentRightText.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTextRightIndent;

                        if (Math.Abs(actualRight - gost.IndentRightText.Value) > 0.05)
                        {
                            errorDetails.Add($"Правый отступ: {actualRight:F2} см (требуется {gost.IndentRightText.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string shortText = GetShortText(paragraph);
                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"'{shortText}' - {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }

                // 5. ОБНОВЛЕНИЕ UI
                updateUI?.Invoke( !isValid ? "Ошибки в отступах:\n" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3)) : "Отступы соответствуют ГОСТу", 
                                  !isValid ? Brushes.Red : Brushes.Green);
                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки межстрочного интервала для простого текста
        /// </summary>
        /// <param name="requiredLineSpacing"></param>
        /// <param name="requiredLineSpacingType"></param>
        /// <param name="paragraphs"></param>
        /// <param name="doc"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckLineSpacingAsync(double requiredLineSpacing, string requiredLineSpacingType, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
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

                        if (actualSpacingType != requiredLineSpacingType)
                        {
                            tempErrors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"Тип интервала: '{actualSpacingType}' (требуется '{requiredLineSpacingType}')",
                                ProblemRun = run,
                                ProblemParagraph = null
                            });
                            isValid = false;
                        }

                        if (Math.Abs(actualSpacing - requiredLineSpacing) > 0.01)
                        {
                            tempErrors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"Интервал: {actualSpacing:F2} (требуется {requiredLineSpacing:F2})",
                                ProblemRun = run,
                                ProblemParagraph = null
                            });
                            isValid = false;
                        }
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки интервала:\n" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3)) : "Интервал соответствует ГОСТу", 
                                                                                                                               !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки выравнивания текста
        /// </summary>
        /// <param name="requiredAlignment"></param>
        /// <param name="paragraphs"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckTextAlignmentAsync(string requiredAlignment, List<Paragraph> paragraphs, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification) ?? DefaultTextAlignment;

                    if (currentAlignment != requiredAlignment)
                    {
                        var shortText = GetShortText(paragraph);
                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"'{shortText}' (выравнивание: '{currentAlignment}', требуется: '{requiredAlignment}')",
                            ProblemParagraph = paragraph, 
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }
                updateUI?.Invoke(isValid ? "Выравнивание соответствует ГОСТу" : "Ошибки в выравнивании:\n" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3)),
                                                                                                                                        isValid ? Brushes.Green : Brushes.Red);
                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки размера шрифта
        /// </summary>
        /// <param name="requiredFontSize"></param>
        /// <param name="paragraphs"></param>
        /// <param name="doc"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckFontSizeAsync(double requiredFontSize, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
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
                                tempErrors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"{runText} (размер: {actualSize:F1} pt) - требуется ({requiredFontSize:F1} pt)",
                                    ProblemRun = run,
                                    ProblemParagraph = null 
                                });
                                isValid = false;
                            }
                        }
                    }
                }

                updateUI?.Invoke(!isValid ? "Ошибки в размере шрифта:\n" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3)) : "Размер шрифта соответствует ГОСТу",
                                                                                                                                           !isValid ? Brushes.Red : Brushes.Green);
                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки шрифта
        /// </summary>
        /// <param name="requiredFontName"></param>
        /// <param name="paragraphs"></param>
        /// <param name="doc"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckFontNameAsync(string requiredFontName, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
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
                            tempErrors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"Ошибка шрифта: '{fontName}' (требуется '{requiredFontName}')",
                                ProblemRun = run,       // Указываем проблемный Run
                                ProblemParagraph = null // Ошибка не на уровне абзаца
                            });
                            isValid = false;
                        }
                    }
                }
                updateUI?.Invoke(isValid ? "Шрифт корректен" : "Ошибки шрифта:\n" + string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3)), isValid ? Brushes.Green :
                                                                                                                                                                 Brushes.Red);
                return (isValid, tempErrors);
            });
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
        /// Определяет тип межстрочного интервала
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
                return double.Parse(spacing.Line.Value) / 567.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                return double.Parse(spacing.Line.Value) / 567.0;
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
            return value / 567.0;
        }

        /// <summary>
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twips"></param>
        /// <returns></returns>
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
        /// <param name="run"></param>
        /// <returns></returns>
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