using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Avalonia.Media;
using Avalonia.Threading;
using System.Threading.Tasks;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок текста подписи картинок
    /// </summary>
    public class CheckingImageDoc
    {
        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ПОДПИСЕЙ К КАРТИНКАМ =======================
        private const string DefaultImageCaptionFont = "Arial";
        private const double DefaultImageCaptionFontSize = 11.0;
        private const string DefaultImageCaptionIndentOrOutdent = "Нет";
        private const double DefaultImageCaptionFirstLineIndent = 1.25;
        private const double DefaultImageCaptionIndentLeft = 0.0;
        private const double DefaultImageCaptionIndentRight = 0.0;
        private const string DefaultImageCaptionAlignment = "Left";
        private const string DefaultImageCaptionLineSpacingType = "Множитель";
        private const double DefaultImageCaptionLineSpacingValue = 1.15;
        private const double DefaultImageCaptionLineSpacingBefore = 0.0;
        private const double DefaultImageCaptionLineSpacingAfter = 0.35;

        private readonly WordprocessingDocument _wordDoc;
        private readonly Gost _gost;
        private readonly Func<Run, bool> _shouldSkipRun;

        public CheckingImageDoc(WordprocessingDocument wordDoc, Gost gost, Func<Run, bool> shouldSkipRun)
        {
            _wordDoc = wordDoc;
            _gost = gost;
            _shouldSkipRun = shouldSkipRun;
        }

        /// <summary>
        /// Проверяет изображения и их подписи на соответствие ГОСТу
        /// </summary>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckImagesAsync(List<Paragraph> paragraphs, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var body = _wordDoc.MainDocumentPart.Document.Body;
                var errors = new List<TextErrorInfo>();
                bool allImagesValid = true;
                bool hasAtLeastOneImage = false;

                // Проверка наличия изображений
                foreach (var paragraph in paragraphs)
                {
                    var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                  paragraph.Descendants<Picture>().Any();

                    if (hasImage)
                    {
                        hasAtLeastOneImage = true;
                        break;
                    }
                }

                if (!hasAtLeastOneImage)
                {
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke("Рисунки не обнаружены — проверка не требуется.", Brushes.Green);
                    });
                    return (true, errors);
                }

                // Проверка шрифта подписей
                if (!string.IsNullOrEmpty(_gost.ImageCaptionFontName))
                {
                    bool fontNameValid = true;
                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                      paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            // Ищем следующий непустой параграф с текстом
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                // Проверка формата подписи
                                if (!CheckImageCaptionFormat(captionParagraph, errors))
                                {
                                    allImagesValid = false;
                                }

                                foreach (var run in captionParagraph.Elements<Run>())
                                {
                                    if (_shouldSkipRun(run)) continue;

                                    var font = run.RunProperties?.RunFonts?.Ascii?.Value;
                                    if (font != null && font != _gost.ImageCaptionFontName)
                                    {
                                        fontNameValid = false;
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"Шрифт подписи под рисунком должен быть: {_gost.ImageCaptionFontName}, а не {font}",
                                            ProblemRun = run,
                                            ProblemParagraph = captionParagraph
                                        });
                                    }
                                }
                            }
                        }
                    }

                    allImagesValid &= fontNameValid;
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(fontNameValid ? "Шрифт подписей соответствует ГОСТу." : "Ошибки в шрифте подписей.",
                                          fontNameValid ? Brushes.Green : Brushes.Red);
                    });
                }

                // Проверка размера шрифта подписей
                if (_gost.ImageCaptionFontSize.HasValue)
                {
                    bool fontSizeValid = true;
                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                      paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (var run in captionParagraph.Elements<Run>())
                                {
                                    if (_shouldSkipRun(run)) continue;

                                    var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value;
                                    double actualFontSize = fontSizeVal != null ? double.Parse(fontSizeVal) / 2 : DefaultImageCaptionFontSize;

                                    if (Math.Abs(actualFontSize - _gost.ImageCaptionFontSize.Value) > 0.1)
                                    {
                                        fontSizeValid = false;
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"Размер шрифта подписи должен быть {_gost.ImageCaptionFontSize.Value}, а не {actualFontSize}",
                                            ProblemRun = run,
                                            ProblemParagraph = captionParagraph
                                        });
                                    }
                                }
                            }
                        }
                    }

                    allImagesValid &= fontSizeValid;
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(fontSizeValid ? "Размер шрифта подписей соответствует ГОСТу." : "Ошибки в размере шрифта подписей.",
                                          fontSizeValid ? Brushes.Green : Brushes.Red);
                    });
                }

                // Проверка выравнивания подписей
                if (!string.IsNullOrEmpty(_gost.ImageCaptionAlignment))
                {
                    bool alignmentValid = true;
                    string requiredAlignment = _gost.ImageCaptionAlignment.ToLowerInvariant();

                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                      paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                string currentAlignment = GetAlignmentString(captionParagraph.ParagraphProperties?.Justification)?.ToLowerInvariant() ??
                                                        DefaultImageCaptionAlignment.ToLowerInvariant();

                                if (currentAlignment != requiredAlignment)
                                {
                                    alignmentValid = false;
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"Подпись под рисунком должна быть выровнена: {requiredAlignment}, а не {currentAlignment}",
                                        ProblemRun = null,
                                        ProblemParagraph = captionParagraph
                                    });
                                }
                            }
                        }
                    }

                    allImagesValid &= alignmentValid;
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(alignmentValid ? "Выравнивание подписей соответствует ГОСТу." : "Ошибки в выравнивании подписей.",
                                        alignmentValid ? Brushes.Green : Brushes.Red);
                    });
                }

                // Проверка отступов подписей изображений
                if (_gost.ImageCaptionFirstLineIndent.HasValue || _gost.ImageCaptionIndentLeft.HasValue || _gost.ImageCaptionIndentRight.HasValue)
                {
                    bool indentsValid = true;
                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                      paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                var indent = captionParagraph.ParagraphProperties?.Indentation;
                                var errorDetails = new List<string>();

                                // Преобразуем все значения в сантиметры
                                double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                                double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                                double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

                                // 1. Проверка визуального левого отступа подписи
                                if (_gost.ImageCaptionIndentLeft.HasValue)
                                {
                                    double actualTextIndent = leftIndent;

                                    if (hangingIndent > 0)
                                    {
                                        actualTextIndent = leftIndent - hangingIndent;
                                    }

                                    if (Math.Abs(actualTextIndent - _gost.ImageCaptionIndentLeft.Value) > 0.05)
                                    {
                                        errorDetails.Add($"Левый отступ подписи: {actualTextIndent:F2} см (требуется {_gost.ImageCaptionIndentLeft.Value:F2} см)");
                                        indentsValid = false;
                                    }
                                }

                                // 2. Проверка правого отступа подписи
                                if (_gost.ImageCaptionIndentRight.HasValue)
                                {
                                    double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultImageCaptionIndentRight;

                                    if (Math.Abs(actualRight - _gost.ImageCaptionIndentRight.Value) > 0.05)
                                    {
                                        errorDetails.Add($"Правый отступ подписи: {actualRight:F2} см (требуется {_gost.ImageCaptionIndentRight.Value:F2} см)");
                                        indentsValid = false;
                                    }
                                }

                                // 3. Проверка первой строки подписи
                                if (_gost.ImageCaptionFirstLineIndent.HasValue)
                                {
                                    bool isHanging = hangingIndent > 0;
                                    bool isFirstLine = firstLineIndent > 0;

                                    if (!string.IsNullOrEmpty(_gost.ImageCaptionIndentOrOutdent))
                                    {
                                        bool typeError = false;

                                        if (_gost.ImageCaptionIndentOrOutdent == "Выступ" && !isHanging)
                                            typeError = true;
                                        else if (_gost.ImageCaptionIndentOrOutdent == "Отступ" && !isFirstLine)
                                            typeError = true;
                                        else if (_gost.ImageCaptionIndentOrOutdent == "Нет" && (isHanging || isFirstLine))
                                            typeError = true;

                                        if (typeError)
                                        {
                                            errorDetails.Add($"Тип первой строки подписи: {(isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет")} (требуется {_gost.ImageCaptionIndentOrOutdent})");
                                            indentsValid = false;
                                        }
                                    }

                                    double currentValue = isHanging ? hangingIndent : firstLineIndent;
                                    if ((isHanging || isFirstLine) && Math.Abs(currentValue - _gost.ImageCaptionFirstLineIndent.Value) > 0.05)
                                    {
                                        errorDetails.Add($"{(isHanging ? "Выступ" : "Отступ")} первой строки подписи: {currentValue:F2} см (требуется {_gost.ImageCaptionFirstLineIndent.Value:F2} см)");
                                        indentsValid = false;
                                    }
                                    else if (_gost.ImageCaptionIndentOrOutdent != "Нет" && !isHanging && !isFirstLine)
                                    {
                                        errorDetails.Add($"Отсутствует {_gost.ImageCaptionIndentOrOutdent} первой строки подписи");
                                        indentsValid = false;
                                    }
                                }

                                if (errorDetails.Any())
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"Подпись под рисунком: {string.Join(", ", errorDetails)}",
                                        ProblemRun = null,
                                        ProblemParagraph = captionParagraph
                                    });
                                }
                            }
                        }
                    }

                    allImagesValid &= indentsValid;

                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(indentsValid ? "Отступы подписей соответствуют ГОСТу." : "Ошибки в отступах подписей.",
                                        indentsValid ? Brushes.Green : Brushes.Red);
                    });
                }

                // Проверка межстрочных интервалов подписей
                if (_gost.ImageCaptionLineSpacingValue.HasValue || _gost.ImageCaptionLineSpacingBefore.HasValue ||
                    _gost.ImageCaptionLineSpacingAfter.HasValue || !string.IsNullOrEmpty(_gost.ImageCaptionLineSpacingType))
                {
                    bool spacingValid = true;
                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                      paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                var spacing = captionParagraph.ParagraphProperties?.SpacingBetweenLines;
                                var errorDetails = new List<string>();

                                if (!string.IsNullOrEmpty(_gost.ImageCaptionLineSpacingType))
                                {
                                    string currentSpacingType = ConvertSpacingRuleToName(spacing?.LineRule?.Value);

                                    if (currentSpacingType != _gost.ImageCaptionLineSpacingType)
                                    {
                                        errorDetails.Add($"Тип межстрочного интервала должен быть: {_gost.ImageCaptionLineSpacingType}, а не {currentSpacingType}");
                                        spacingValid = false;
                                    }
                                }

                                if (_gost.ImageCaptionLineSpacingValue.HasValue)
                                {
                                    double actualSpacing = spacing?.Line != null ?
                                        CalculateActualSpacing(spacing) : DefaultImageCaptionLineSpacingValue;

                                    if (Math.Abs(actualSpacing - _gost.ImageCaptionLineSpacingValue.Value) > 0.1)
                                    {
                                        errorDetails.Add($"Межстрочный интервал подписи должен быть {_gost.ImageCaptionLineSpacingValue.Value}, а не {actualSpacing}");
                                        spacingValid = false;
                                    }
                                }

                                if (_gost.ImageCaptionLineSpacingBefore.HasValue)
                                {
                                    double actualBefore = spacing?.Before?.Value != null ?
                                        ConvertTwipsToPoints(spacing.Before.Value) : DefaultImageCaptionLineSpacingBefore;
                                    if (Math.Abs(actualBefore - _gost.ImageCaptionLineSpacingBefore.Value) > 0.1)
                                    {
                                        errorDetails.Add($"Интервал перед подписью должен быть {_gost.ImageCaptionLineSpacingBefore.Value}, а не {actualBefore}");
                                        spacingValid = false;
                                    }
                                }

                                if (_gost.ImageCaptionLineSpacingAfter.HasValue)
                                {
                                    double actualAfter = spacing?.After?.Value != null ?
                                        ConvertTwipsToPoints(spacing.After.Value) : DefaultImageCaptionLineSpacingAfter;
                                    if (Math.Abs(actualAfter - _gost.ImageCaptionLineSpacingAfter.Value) > 0.1)
                                    {
                                        errorDetails.Add($"Интервал после подписи должен быть {_gost.ImageCaptionLineSpacingAfter.Value}, а не {actualAfter}");
                                        spacingValid = false;
                                    }
                                }

                                if (errorDetails.Any())
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"Подпись под рисунком: {string.Join(", ", errorDetails)}",
                                        ProblemRun = null,
                                        ProblemParagraph = captionParagraph
                                    });
                                }
                            }
                        }
                    }

                    allImagesValid &= spacingValid;
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(spacingValid ? "Интервалы подписей соответствуют ГОСТу." : "Ошибки в интервалах подписей.",
                                        spacingValid ? Brushes.Green : Brushes.Red);
                    });
                }

                if (!allImagesValid)
                {
                    Dispatcher.UIThread.Post(() => {
                        var msg = $"Ошибки в подписях под изображениями:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3) msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Brushes.Red);
                    });
                }
                return (allImagesValid, errors);
            });
        }

        /// <summary>
        /// Вспомогательный метод, предназначен для поиска первого непустого параграфа. 
        /// Этот метод используется для корректного определения подписи под изображением
        /// </summary>
        /// <param name="startParagraph"></param>
        /// <returns></returns>
        private Paragraph FindNextNonEmptyParagraph(Paragraph startParagraph)
        {
            var nextElement = startParagraph.NextSibling();
            while (nextElement != null)
            {
                if (nextElement is Paragraph nextParagraph &&
                    !string.IsNullOrWhiteSpace(nextParagraph.InnerText))
                {
                    return nextParagraph;
                }
                nextElement = nextElement.NextSibling();
            }
            return null;
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
        /// Проверка формата подписи рисунков
        /// </summary>
        /// <param name="captionParagraph"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private bool CheckImageCaptionFormat(Paragraph captionParagraph, List<TextErrorInfo> errors)
        {
            // Сначала проверяем, что параграф вообще содержит текст
            if (string.IsNullOrWhiteSpace(captionParagraph.InnerText))
            {
                return true;
            }
            string pattern = @"^Рисунок\s+\d+\s*[-–—]\s*.+";
            string text = captionParagraph.InnerText.Trim();
            bool isValid = Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);

            if (!isValid)
            {
                // Проверяем, действительно ли это подпись (может содержать слово "Рисунок")
                if (text.StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = $"Неверный формат подписи: '{GetShortText(text)}' (требуется 'Рисунок X - Описание')",
                        ProblemRun = null,
                        ProblemParagraph = captionParagraph
                    });
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Обрезает текст параграфа до 30 символов с добавлением многоточия
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string GetShortText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "[пустой элемент]";
            return text.Length > 30 ? text.Substring(0, 27) + "..." : text;
        }

        /// <summary>
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twips"></param>
        /// <returns></returns>
        private double TwipsToCm(double twips) => twips / 567.0;

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
    }
}