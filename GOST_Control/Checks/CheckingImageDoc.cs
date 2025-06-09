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
        private const double DefaultImageCaptionFontSize = 11.0;
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

                var allStyles = _wordDoc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                // Проверка наличия изображений
                foreach (var paragraph in paragraphs)
                {
                    var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() || paragraph.Descendants<Picture>().Any();

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
                    var fontErrors = new List<TextErrorInfo>();

                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() ||
                                      paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                var (isValidFormat, captionText) = CheckImageCaptionFormat(captionParagraph, errors);
                                if (!isValidFormat)
                                {
                                    allImagesValid = false;
                                }

                                bool hasFontError = false;
                                Run problematicRun = null;

                                foreach (var run in captionParagraph.Elements<Run>())
                                {
                                    if (_shouldSkipRun(run)) continue;

                                    var (actualFont, isFontDefined) = GetActualFontForRun(run, captionParagraph, allStyles);

                                    if (!isFontDefined && !hasFontError)
                                    {
                                        problematicRun = run;
                                        hasFontError = true;
                                        fontNameValid = false;
                                    }
                                    else if (!string.Equals(actualFont, _gost.ImageCaptionFontName, StringComparison.OrdinalIgnoreCase) && !hasFontError)
                                    {
                                        problematicRun = run;
                                        hasFontError = true;
                                        fontNameValid = false;
                                    }
                                }

                                if (hasFontError)
                                {
                                    var (actualFont, _) = GetActualFontForRun(problematicRun, captionParagraph, allStyles);

                                    fontErrors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = actualFont == null ? "       • Не удалось определить шрифт подписи рисунка"
                                        : $"      • Шрифт подписи под рисунком должен быть: {_gost.ImageCaptionFontName}, а не {actualFont}",
                                        ProblemRun = problematicRun,
                                        ProblemParagraph = captionParagraph
                                    });
                                }
                            }
                        }
                    }

                    errors.AddRange(fontErrors);
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
                    var fontSizeErrors = new List<TextErrorInfo>();

                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() || paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                bool hasFontSizeError = false;
                                Run problematicRun = null;
                                double? actualSize = null;

                                foreach (var run in captionParagraph.Elements<Run>())
                                {
                                    if (_shouldSkipRun(run)) continue;

                                    var (size, isSizeDefined) = GetActualFontSizeForRun(run, captionParagraph, allStyles);

                                    if (!isSizeDefined && !hasFontSizeError)
                                    {
                                        problematicRun = run;
                                        hasFontSizeError = true;
                                        fontSizeValid = false;
                                    }
                                    else if (isSizeDefined && Math.Abs(size - _gost.ImageCaptionFontSize.Value) > 0.1 && !hasFontSizeError)
                                    {
                                        problematicRun = run;
                                        actualSize = size;
                                        hasFontSizeError = true;
                                        fontSizeValid = false;
                                    }
                                }

                                if (hasFontSizeError)
                                {
                                    fontSizeErrors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = actualSize == null ? "       • Не удалось определить размер шрифта подписи рисунка"
                                        : $"       • Размер шрифта подписи должен быть {_gost.ImageCaptionFontSize.Value:F1} pt, а не {actualSize.Value:F1} pt",
                                        ProblemRun = problematicRun,
                                        ProblemParagraph = captionParagraph
                                    });
                                }
                            }
                        }
                    }

                    errors.AddRange(fontSizeErrors);
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
                    string requiredAlignment = _gost.ImageCaptionAlignment;

                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() || paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                var (actualAlignment, isAlignmentDefined) = GetActualAlignmentForParagraph(captionParagraph, allStyles);

                                if (!isAlignmentDefined)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = "\n       • Не удалось определить выравнивание подписи рисунка",
                                        ProblemParagraph = captionParagraph,
                                        ProblemRun = null
                                    });
                                    alignmentValid = false;
                                }
                                else if (actualAlignment != requiredAlignment)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Подпись под рисунком должна быть выровнена: {requiredAlignment}, а не {actualAlignment}",
                                        ProblemParagraph = captionParagraph,
                                        ProblemRun = null
                                    });
                                    alignmentValid = false;
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
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() || paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                // Получаем отступы с учетом стилей
                                var indent = captionParagraph.ParagraphProperties?.Indentation;
                                var styleIndent = GetStyleIndentationForParagraph(captionParagraph, allStyles);
                                indent ??= styleIndent;

                                var errorDetails = new List<string>();

                                // Преобразуем все значения в сантиметры
                                double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                                double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                                double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;
                                double rightIndent = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultImageCaptionIndentRight;

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
                                        errorDetails.Add($"       • Левый отступ подписи: {actualTextIndent:F2} см (требуется {_gost.ImageCaptionIndentLeft.Value:F2} см)");
                                        indentsValid = false;
                                    }
                                }

                                // 2. Проверка правого отступа подписи
                                if (_gost.ImageCaptionIndentRight.HasValue)
                                {
                                    if (Math.Abs(rightIndent - _gost.ImageCaptionIndentRight.Value) > 0.05)
                                    {
                                        errorDetails.Add($"\n       • Правый отступ подписи: {rightIndent:F2} см (требуется {_gost.ImageCaptionIndentRight.Value:F2} см)");
                                        indentsValid = false;
                                    }
                                }

                                // 3. Проверка первой строки подписи
                                if (_gost.ImageCaptionFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.ImageCaptionIndentOrOutdent))
                                {
                                    bool isHanging = hangingIndent > 0;
                                    bool isFirstLine = firstLineIndent > 0;

                                    // Проверка типа (выступ/отступ)
                                    if (!string.IsNullOrEmpty(_gost.ImageCaptionIndentOrOutdent))
                                    {
                                        string actualType = isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет";
                                        if (actualType != _gost.ImageCaptionIndentOrOutdent)
                                        {
                                            errorDetails.Add($"\n       • Тип первой строки подписи: '{actualType}' (требуется '{_gost.ImageCaptionIndentOrOutdent}')");
                                            indentsValid = false;
                                        }
                                    }

                                    // Проверка значения отступа/выступа
                                    if (_gost.ImageCaptionFirstLineIndent.HasValue)
                                    {
                                        double actualValue = isHanging ? hangingIndent : firstLineIndent;
                                        if ((isHanging || isFirstLine) && Math.Abs(actualValue - _gost.ImageCaptionFirstLineIndent.Value) > 0.05)
                                        {
                                            errorDetails.Add($"\n       • {(isHanging ? "Выступ" : "Отступ")} первой строки подписи: {actualValue:F2} см (требуется {_gost.ImageCaptionFirstLineIndent.Value:F2} см)");
                                            indentsValid = false;
                                        }
                                        else if (!isHanging && !isFirstLine && _gost.ImageCaptionIndentOrOutdent != "Нет")
                                        {
                                            errorDetails.Add($"\n       • Отсутствует {_gost.ImageCaptionIndentOrOutdent} первой строки подписи");
                                            indentsValid = false;
                                        }
                                    }
                                }

                                if (errorDetails.Any())
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = string.Join("", errorDetails),
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
                if (_gost.ImageCaptionLineSpacingValue.HasValue || !string.IsNullOrEmpty(_gost.ImageCaptionLineSpacingType))
                {
                    bool spacingValid = true;
                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any() || paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacingForParagraph(captionParagraph, allStyles);

                                if (!isSpacingDefined)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = "       • Не удалось определить межстрочный интервал подписи рисунка",
                                        ProblemParagraph = captionParagraph,
                                        ProblemRun = null
                                    });
                                    spacingValid = false;
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(_gost.ImageCaptionLineSpacingType))
                                    {
                                        if (actualSpacingType != _gost.ImageCaptionLineSpacingType)
                                        {
                                            errors.Add(new TextErrorInfo
                                            {
                                                ErrorMessage = $"       • Тип межстрочного интервала подписи: '{actualSpacingType}' (требуется '{_gost.ImageCaptionLineSpacingType}')",
                                                ProblemParagraph = captionParagraph,
                                                ProblemRun = null
                                            });
                                            spacingValid = false;
                                        }
                                    }

                                    // Проверка значения интервала
                                    if (_gost.ImageCaptionLineSpacingValue.HasValue &&
                                        Math.Abs(actualSpacingValue - _gost.ImageCaptionLineSpacingValue.Value) > 0.1)
                                    {
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"       • Межстрочный интервал подписи: {actualSpacingValue:F2} (требуется {_gost.ImageCaptionLineSpacingValue.Value:F2})",
                                            ProblemParagraph = captionParagraph,
                                            ProblemRun = null
                                        });
                                        spacingValid = false;
                                    }
                                }
                            }
                        }
                    }

                    allImagesValid &= spacingValid;
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(spacingValid ? "Межстрочные интервалы подписей соответствуют ГОСТу." : "Ошибки в межстрочных интервалах подписей.",
                                         spacingValid ? Brushes.Green : Brushes.Red);
                    });
                }

                // Проверка интервалов перед/после подписей
                if (_gost.ImageCaptionLineSpacingBefore.HasValue || _gost.ImageCaptionLineSpacingAfter.HasValue)
                {
                    bool spacingBeforeAfterValid = true;
                    foreach (var paragraph in paragraphs)
                    {
                        var hasImage = paragraph.Descendants<DocumentFormat.OpenXml.Office.Drawing.Drawing>().Any()
                                     || paragraph.Descendants<Picture>().Any();

                        if (hasImage)
                        {
                            Paragraph captionParagraph = FindNextNonEmptyParagraph(paragraph);

                            if (captionParagraph != null && captionParagraph.InnerText.Trim().StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
                            {
                                var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacingForCaption(captionParagraph, allStyles);

                                if (!isSpacingDefined)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = "       • Не удалось определить интервалы перед/после подписи рисунка",
                                        ProblemParagraph = captionParagraph,
                                        ProblemRun = null
                                    });
                                    spacingBeforeAfterValid = false;
                                }
                                else
                                {
                                    // Проверка интервала перед абзацем
                                    if (_gost.ImageCaptionLineSpacingBefore.HasValue &&
                                        Math.Abs(actualBefore - _gost.ImageCaptionLineSpacingBefore.Value) > 0.1)
                                    {
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"       • Интервал перед подписью: {actualBefore:F1} pt (требуется {_gost.ImageCaptionLineSpacingBefore.Value:F1} pt)",
                                            ProblemParagraph = captionParagraph,
                                            ProblemRun = null
                                        });
                                        spacingBeforeAfterValid = false;
                                    }

                                    // Проверка интервала после абзаца
                                    if (_gost.ImageCaptionLineSpacingAfter.HasValue &&
                                        Math.Abs(actualAfter - _gost.ImageCaptionLineSpacingAfter.Value) > 0.1)
                                    {
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"       • Интервал после подписи: {actualAfter:F1} pt (требуется {_gost.ImageCaptionLineSpacingAfter.Value:F1} pt)",
                                            ProblemParagraph = captionParagraph,
                                            ProblemRun = null
                                        });
                                        spacingBeforeAfterValid = false;
                                    }
                                }
                            }
                        }
                    }

                    allImagesValid &= spacingBeforeAfterValid;
                    Dispatcher.UIThread.Post(() => {
                        updateUI?.Invoke(spacingBeforeAfterValid ? "Интервалы перед/после подписей соответствуют ГОСТу." : "Ошибки в интервалах перед/после подписей.",
                                         spacingBeforeAfterValid ? Brushes.Green : Brushes.Red);
                    });
                }

                if (!allImagesValid)
                {
                    Dispatcher.UIThread.Post(() => {
                        var groupedErrors = errors.GroupBy(e => e.ProblemParagraph).Select(g => new {
                                Paragraph = g.Key,
                                Caption = g.Key?.InnerText?.Trim() ?? "Неизвестный рисунок",
                                Errors = g.ToList()
                            });

                        var errorMessages = new List<string>();

                        foreach (var group in groupedErrors)
                        {
                            // Ограничиваем длину названия рисунка
                            var shortCaption = group.Caption.Length > 50 ? group.Caption.Substring(0, 47) + "..." : group.Caption;

                            errorMessages.Add($"Ошибки в подписях под изображение '{shortCaption}':");

                            // Добавляем первые 3 ошибки для этого рисунка
                            errorMessages.AddRange(group.Errors.Take(300).Select(e => e.ErrorMessage));

                            // Если ошибок больше 3, добавляем сообщение об этом
                            if (group.Errors.Count > 3)
                            {
                                errorMessages.Add($"...и ещё {group.Errors.Count - 3} ошибок");
                            }

                            errorMessages.Add(""); // Пустая строка между группами
                        }

                        var msg = string.Join("\n", errorMessages);
                        updateUI?.Invoke(msg, Brushes.Red);
                    });
                }

                return (allImagesValid, errors);
            });
        }

        private (string Type, double Value, bool IsDefined) GetActualLineSpacingForParagraph(Paragraph paragraph, Dictionary<string, Style> allStyles)
        {
            // 1. Проверяем явные свойства абзаца
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            var parsed = ParseLineSpacing(spacing);
            if (parsed.IsDefined)
                return parsed;

            // 2. Проверяем стиль абзаца и его родителей
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var currentStyle))
            {
                while (currentStyle != null)
                {
                    var styleSpacing = currentStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    parsed = ParseLineSpacing(styleSpacing);
                    if (parsed.IsDefined)
                        return parsed;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null
                                 && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle)
                                 ? basedOnStyle : null;
                }
            }

            // 3. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle))
            {
                parsed = ParseLineSpacing(normalStyle.StyleParagraphProperties?.SpacingBetweenLines);
                if (parsed.IsDefined)
                    return parsed;
            }

            return ("Множитель", 1.0, true);
        }

        private (string Type, double Value, bool IsDefined) ParseLineSpacing(SpacingBetweenLines spacing)
        {
            if (spacing == null)
                return (null, 0, false);

            if (spacing.LineRule != null && spacing.Line == null)
                return (null, 0, false);

            if (spacing.Line != null && spacing.LineRule == null)
            {
                double lineValue = double.Parse(spacing.Line.Value);
                return ("Множитель", lineValue / 240.0, true);
            }

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

            return (null, 0, false);
        }

        /// <summary>
        /// Определение стиля интервалов для подписи
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private (double Before, double After, bool IsDefined) GetActualParagraphSpacingForCaption(Paragraph paragraph, Dictionary<string, Style> allStyles)
        {
            double? before = null;
            double? after = null;

            // 1. Проверяем явные свойства абзаца
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing != null)
            {
                before = spacing.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : 0;
                after = spacing.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : 0;
                return (before.Value, after.Value, true);
            }

            // 2. Проверяем стиль абзаца и его родителей
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var currentStyle))
            {
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

                    // Если нашли оба значения, прерываем цикл
                    if (before != null && after != null)
                        break;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            // 3. Проверяем стиль по умолчанию для заголовков таблиц (если есть)
            if ((before == null || after == null) && allStyles.TryGetValue("TableCaption", out var captionStyle))
            {
                var styleSpacing = captionStyle.StyleParagraphProperties?.SpacingBetweenLines;
                if (styleSpacing != null)
                {
                    if (before == null && styleSpacing.Before?.Value != null)
                        before = ConvertTwipsToPoints(styleSpacing.Before.Value);
                    if (after == null && styleSpacing.After?.Value != null)
                        after = ConvertTwipsToPoints(styleSpacing.After.Value);
                }
            }

            // 4. Проверяем Normal стиль
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

        private Indentation GetStyleIndentationForParagraph(Paragraph paragraph, Dictionary<string, Style> allStyles)
        {
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var currentStyle))
            {
                while (currentStyle != null)
                {
                    var indent = currentStyle.StyleParagraphProperties?.Indentation;
                    if (indent != null)
                        return indent;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                }
            }

            // Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle))
            {
                return normalStyle.StyleParagraphProperties?.Indentation;
            }

            return null;
        }

        private (string Alignment, bool IsDefined) GetActualAlignmentForParagraph(Paragraph paragraph, Dictionary<string, Style> allStyles)
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

            // 3. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle)
                && normalStyle.StyleParagraphProperties?.Justification?.Val?.Value != null)
            {
                return (GetAlignmentString(normalStyle.StyleParagraphProperties.Justification), true);
            }

            return (null, false);
        }

        private (double Size, bool IsDefined) GetActualFontSizeForRun(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles)
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
            if (allStyles.TryGetValue("Normal", out var normalStyle) &&
                normalStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
            {
                return (double.Parse(normalStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);
            }

            return (0, false);
        }

        private (string FontName, bool IsDefined) GetActualFontForRun(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles)
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

        private string GetExplicitRunFont(Run run)
        {
            var runProps = run.RunProperties;
            if (runProps == null) return null;

            return runProps.RunFonts?.Ascii?.Value ?? runProps.RunFonts?.HighAnsi?.Value
                   ?? runProps.RunFonts?.ComplexScript?.Value ?? runProps.RunFonts?.EastAsia?.Value;
        }

        private string GetStyleFont(Style style)
        {
            if (style?.StyleRunProperties == null) return null;
            return style.StyleRunProperties.RunFonts?.Ascii?.Value ?? style.StyleRunProperties.RunFonts?.HighAnsi?.Value
                   ?? style.StyleRunProperties.RunFonts?.ComplexScript?.Value ?? style.StyleRunProperties.RunFonts?.EastAsia?.Value;
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
        /// Проверка формата подписи рисунков
        /// </summary>
        /// <param name="captionParagraph"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private (bool IsValid, string CaptionText) CheckImageCaptionFormat(Paragraph captionParagraph, List<TextErrorInfo> errors)
        {
            if (string.IsNullOrWhiteSpace(captionParagraph.InnerText))
            {
                return (true, null);
            }

            string text = captionParagraph.InnerText.Trim();
            string pattern = @"^Рисунок\s+\d+\s*[-–—]\s*.+";
            bool isValid = Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);

            if (!isValid && text.StartsWith("Рисунок", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(new TextErrorInfo
                {
                    ErrorMessage = $"Неверный формат подписи: '{GetShortText(text)}' (требуется 'Рисунок X - Описание')",
                    ProblemRun = null,
                    ProblemParagraph = captionParagraph
                });
                return (false, text);
            }

            return (isValid, text);
        }

        /// <summary>
        /// Обрезает текст параграфа до 50 символов с добавлением многоточия
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string GetShortText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "[пустой элемент]";
            return text.Length > 50 ? text.Substring(0, 47) + "..." : text;
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
    }
}