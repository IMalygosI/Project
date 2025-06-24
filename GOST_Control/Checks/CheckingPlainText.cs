using Avalonia.Media;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверки чистого текста
    /// </summary>
    public class CheckingPlainText
    {
        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ПРОСТОГО ТЕКСТА =======================
        private const string DefaultTextFont = "Не обнаружен!";
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
        private readonly WordprocessingDocument _wordDoc;

        public CheckingPlainText(WordprocessingDocument wordDoc, Func<Paragraph, bool> shouldSkipParagraph)
        {
            _wordDoc = wordDoc;
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

                var normalStyle = GetStyleById("Normal");
                var defaultSpacing = normalStyle?.StyleParagraphProperties?.SpacingBetweenLines;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var explicitSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    var styleSpacing = GetParagraphStyleSpacing(paragraph, doc);
                    var spacing = explicitSpacing ?? styleSpacing ?? defaultSpacing;

                    bool hasError = false;
                    var errorDetails = new List<string>();

                    if (hasBeforeSpacing && gost.LineSpacingBefore.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : 0;

                        if (Math.Abs(actualBefore - gost.LineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"\n              • интервал перед: {actualBefore:F1} pt (требуется {gost.LineSpacingBefore.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    if (hasAfterSpacing && gost.LineSpacingAfter.HasValue)
                    {
                        double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : 0;

                        if (Math.Abs(actualAfter - gost.LineSpacingAfter.Value) > 0.01)
                        {
                            errorDetails.Add($"\n              • интервал после: {actualAfter:F1} pt (требуется {gost.LineSpacingAfter.Value:F1} pt)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string shortText = GetShortText(paragraph);
                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Абзац '{shortText}': {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? $"Ошибки в интервалах между абзацами:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3))}" +
                                (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") : "Интервалы между абзацами соответствуют ГОСТу",
                                 !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод проверки отступов первой строки
        /// </summary>
        /// <param name="requiredFirstLineIndent"></param>
        /// <param name="paragraphs"></param>
        /// <param name="gost"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckFirstLineIndentAsync(double requiredFirstLineIndent, List<Paragraph> paragraphs, WordprocessingDocument doc, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var explicitIndent = paragraph.ParagraphProperties?.Indentation;
                    var styleIndent = GetParagraphStyleIndentation(paragraph, doc);
                    var indent = explicitIndent ?? styleIndent;

                    bool hasError = false;
                    var errorDetails = new List<string>();

                    double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                    double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                    double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

                    if (gost.IndentLeftText.HasValue)
                    {
                        double actualTextIndent = leftIndent;
                        if (hangingIndent > 0) actualTextIndent = leftIndent - hangingIndent;

                        if (Math.Abs(actualTextIndent - gost.IndentLeftText.Value) > 0.05)
                        {
                            errorDetails.Add($"\n              • левый отступ текста: {actualTextIndent:F2} см (требуется {gost.IndentLeftText.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (gost.FirstLineIndent.HasValue || !string.IsNullOrEmpty(gost.TextIndentOrOutdent))
                    {
                        bool isHanging = hangingIndent > 0;
                        bool isFirstLine = firstLineIndent > 0;

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
                                errorDetails.Add($"\n              • тип первой строки: {(isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет")} (требуется {gost.TextIndentOrOutdent})");
                                hasError = true;
                            }
                        }

                        double currentValue = isHanging ? hangingIndent : firstLineIndent;
                        if ((isHanging || isFirstLine) && Math.Abs(currentValue - gost.FirstLineIndent.Value) > 0.05)
                        {
                            errorDetails.Add($"\n              • {(isHanging ? "Выступ" : "Отступ")} первой строки: {currentValue:F2} см (требуется {gost.FirstLineIndent.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (gost.IndentRightText.HasValue)
                    {
                        double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTextRightIndent;

                        if (Math.Abs(actualRight - gost.IndentRightText.Value) > 0.05)
                        {
                            errorDetails.Add($"\n              • правый отступ: {actualRight:F2} см (требуется {gost.IndentRightText.Value:F2} см)");
                            hasError = true;
                        }
                    }

                    if (hasError)
                    {
                        string shortText = GetShortText(paragraph);
                        tempErrors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Абзац '{shortText}': {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? $"Ошибки в отступах:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3))}" +
                                 (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") : "Отступы соответствуют ГОСТу",
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

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacing(paragraph, doc);

                    if (!isSpacingDefined)
                    {
                        AddSpacingError(tempErrors, paragraph, "не удалось определить междустрочный интервал");
                        isValid = false;
                        continue;
                    }

                    bool hasError = false;
                    var errorDetails = new List<string>();

                    if (actualSpacingType != requiredLineSpacingType)
                    {
                        errorDetails.Add($"\n              • тип интервала: '{actualSpacingType}' (требуется '{requiredLineSpacingType}')");
                        hasError = true;
                    }

                    if (Math.Abs(actualSpacingValue - requiredLineSpacing) > 0.01)
                    {
                        errorDetails.Add($"\n              • интервал: {actualSpacingValue:F2} (требуется {requiredLineSpacing:F2})");
                        hasError = true;
                    }

                    if (hasError)
                    {
                        AddSpacingError(tempErrors, paragraph, string.Join(", ", errorDetails));
                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? $"Ошибки интервала:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3))}" +
                                 (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") : "Интервал соответствует ГОСТу",
                                 !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        private (string Type, double Value, bool IsDefined) GetActualLineSpacing(Paragraph paragraph, WordprocessingDocument doc)
        {
            // 1. Проверяем явные свойства абзаца
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            var parsed = ParseLineSpacing(spacing);
            if (parsed.IsDefined)
                return parsed;

            // 2. Проверяем стиль абзаца и его родительские стили
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null)
            {
                var currentStyle = GetStyleById( paraStyleId);
                while (currentStyle != null)
                {
                    var styleSpacing = currentStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    parsed = ParseLineSpacing(styleSpacing);
                    if (parsed.IsDefined)
                        return parsed;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
                }
            }

            // 3. Проверяем Normal стиль
            var normalStyle = GetStyleById("Normal");
            parsed = ParseLineSpacing(normalStyle?.StyleParagraphProperties?.SpacingBetweenLines);
            if (parsed.IsDefined)
                return parsed;

            return (DefaultTextLineSpacingType, DefaultTextLineSpacingValue, false);
        }

        private (string Type, double Value, bool IsDefined) ParseLineSpacing(SpacingBetweenLines spacing)
        {
            if (spacing?.Line != null)
            {
                double value = double.Parse(spacing.Line.Value.ToString()) / 240.0; // Конвертация из twips в множитель
                string type = ConvertSpacingRuleToName(spacing.LineRule?.Value);
                return (type, value, true);
            }
            return (null, 0, false);
        }

        private void AddSpacingError(List<TextErrorInfo> errors, Paragraph paragraph, string errorMessage)
        {
            var runText = paragraph.InnerText.Trim();
            if (!string.IsNullOrWhiteSpace(runText))
            {
                string shortText = runText.Length > 30 ? runText.Substring(0, 27) + "..." : runText;
                errors.Add(new TextErrorInfo
                {
                    ErrorMessage = $"       • Текст '{shortText}': {errorMessage}",
                    ProblemRun = null,
                    ProblemParagraph = paragraph
                });
            }
        }


        /// <summary>
        /// Метод проверки выравнивания текста
        /// </summary>
        /// <param name="requiredAlignment"></param>
        /// <param name="paragraphs"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckTextAlignmentAsync(string requiredAlignment, List<Paragraph> paragraphs, WordprocessingDocument doc, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<TextErrorInfo>();
                bool isValid = true;

                foreach (var paragraph in paragraphs)
                {
                    if (_shouldSkipParagraph(paragraph))
                        continue;

                    var (actualAlignment, isAlignmentDefined) = GetActualAlignment(paragraph, doc);

                    if (!isAlignmentDefined)
                    {
                        AddAlignmentError(tempErrors, paragraph, "не удалось определить выравнивание");
                        isValid = false;
                    }
                    else if (actualAlignment != requiredAlignment)
                    {
                        AddAlignmentError(tempErrors, paragraph,
                            $"выравнивание '{actualAlignment}' (требуется '{requiredAlignment}')");
                        isValid = false;
                    }
                }

                updateUI?.Invoke(!isValid ? $"Ошибки в выравнивании:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3))}" +
                                 (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") : "Выравнивание соответствует ГОСТу",
                                 !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        private (string Alignment, bool IsDefined) GetActualAlignment(Paragraph paragraph, WordprocessingDocument doc)
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

            return (DefaultTextAlignment, false);
        }

        private void AddAlignmentError(List<TextErrorInfo> errors, Paragraph paragraph, string errorMessage)
        {
            string shortText = GetShortText(paragraph);
            if (!string.IsNullOrWhiteSpace(shortText))
            {
                errors.Add(new TextErrorInfo
                {
                    ErrorMessage = $"       • Абзац '{shortText}':\n              • {errorMessage}",
                    ProblemParagraph = paragraph,
                    ProblemRun = null
                });
            }
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

                        var (actualSize, isSizeDefined) = GetActualFontSize(run, paragraph);

                        if (!isSizeDefined)
                        {
                            var runText = run.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(runText))
                            {
                                string shortText = runText.Length > 30 ? runText.Substring(0, 27) + "..." : runText;
                                tempErrors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"       • Текст '{shortText}':\n              • не удалось определить размер шрифта",
                                    ProblemRun = run,
                                    ProblemParagraph = null
                                });
                                isValid = false;
                            }
                        }
                        else if (Math.Abs(actualSize - requiredFontSize) > 0.1)
                        {
                            var runText = run.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(runText))
                            {
                                string shortText = runText.Length > 30 ? runText.Substring(0, 27) + "..." : runText;
                                tempErrors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"       • Текст '{shortText}':\n              • размер {actualSize:F1} pt (требуется {requiredFontSize:F1} pt)",
                                    ProblemRun = run,
                                    ProblemParagraph = null
                                });
                                isValid = false;
                            }
                        }
                    }
                }

                updateUI?.Invoke(!isValid ? $"Ошибки в размере шрифта:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3))}" +
                                (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") : "Размер шрифта соответствует ГОСТу",
                                 !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }

        private (double Size, bool IsDefined) GetActualFontSize(Run run, Paragraph paragraph)
        {
            // 1. Проверяем явные свойства Run
            var runSize = run.RunProperties?.FontSize?.Val?.Value;

            if (runSize != null)
                return (double.Parse(runSize.ToString()) / 2, true);

            // 2. Проверяем стиль Run
            var runStyleId = run.RunProperties?.RunStyle?.Val?.Value;

            if (runStyleId != null)
            {
                var runStyle = GetStyleById(runStyleId);
                if (runStyle?.StyleRunProperties?.FontSize?.Val?.Value != null)
                    return (double.Parse(runStyle.StyleRunProperties.FontSize.Val.Value.ToString()) / 2, true);
            }

            // 3. Проверяем стиль Paragraph
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

            if (paraStyleId != null)
            {
                var currentStyle = GetStyleById(paraStyleId);
                while (currentStyle != null)
                {
                    if (currentStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
                        return (double.Parse(currentStyle.StyleRunProperties.FontSize.Val.Value.ToString()) / 2, true);

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null ? GetStyleById(currentStyle.BasedOn.Val.Value) : null;
                }
            }

            // 4. Проверяем Normal стиль
            var normalStyle = GetStyleById("Normal");

            if (normalStyle?.StyleRunProperties?.FontSize?.Val?.Value != null)
                return (double.Parse(normalStyle.StyleRunProperties.FontSize.Val.Value.ToString()) / 2, true);

            return (0, false);
        }

        private Style GetStyleById(string styleId)
        {
            return _wordDoc.MainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId.Value == styleId);
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
                            var runText = run.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(runText))
                            {
                                string shortText = runText.Length > 30 ? runText.Substring(0, 27) + "..." : runText;
                                tempErrors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"       • Текст '{shortText}':\n              • шрифт '{fontName}' (требуется '{requiredFontName}')",
                                    ProblemRun = run,
                                    ProblemParagraph = null
                                });
                                isValid = false;
                            }
                        }
                    }
                }

                updateUI?.Invoke(!isValid ? $"Ошибки шрифта:\n{string.Join("\n", tempErrors.Select(e => e.ErrorMessage).Take(3))}" +
                                (tempErrors.Count > 3 ? $"\n...и ещё {tempErrors.Count - 3} ошибок" : "") : "Шрифт соответствует ГОСТу",
                                 !isValid ? Brushes.Red : Brushes.Green);

                return (isValid, tempErrors);
            });
        }
        
        private Indentation GetParagraphStyleIndentation(Paragraph paragraph, WordprocessingDocument doc)
        {
            if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value != null)
            {
                var styleId = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value;
                var style = GetStyleById(styleId);
                return style?.StyleParagraphProperties?.Indentation;
            }
            return null;
        }

        private SpacingBetweenLines GetParagraphStyleSpacing(Paragraph paragraph, WordprocessingDocument doc)
        {
            if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value != null)
            {
                var styleId = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value;
                var style = GetStyleById(styleId);
                return style?.StyleParagraphProperties?.SpacingBetweenLines;
            }
            return null;
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

        /// <summary>
        /// Поиск стиля
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        private Style GetDefaultStyle(WordprocessingDocument doc)
        {
            var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
            if (stylesPart == null) return null;

            // 1. Сначала ищем стиль "Normal" по имени
            var normalStyle = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.Type == StyleValues.Paragraph && 
                                                                                (s.StyleId?.Value?.Equals("Normal", StringComparison.OrdinalIgnoreCase) ?? false));

            // 2. Если не нашли, ищем стиль с атрибутом Default
            if (normalStyle == null)
            {
                normalStyle = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.Type == StyleValues.Paragraph && (s.Default?.Value ?? false));
            }

            // 3. Если стиль всё ещё не найден, берём первый параграфный стиль
            if (normalStyle == null)
            {
                normalStyle = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.Type == StyleValues.Paragraph);
            }

            return normalStyle;
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