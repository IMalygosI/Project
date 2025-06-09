using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia.Media;
using Avalonia.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Brushes = Avalonia.Media.Brushes;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок списков
    /// </summary>
    public class CheckListDocText
    {
        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ СПИСКОВ =======================
        private const string DefaultListFont = "Arial";
        private const double DefaultListSize = 11.0;
        private const string DefaultListLineSpacingType = "Множитель";
        private const double DefaultListLineSpacingValue = 1.15;
        private const double DefaultListSpacingBefore = 0.0;
        private const double DefaultListSpacingAfter = 0.35;
        private const string BulletAlignment = "Left";

        // для многоуровневых
        private const double DefaultListLevel1BulletIndentLeft = 1.87;
        private const double DefaultListLevel2BulletIndentLeft = 2.5;
        private const double DefaultListLevel3BulletIndentLeft = 3.14;
        private const double DefaultListLevel4BulletIndentLeft = 3.77;
        private const double DefaultListLevel5BulletIndentLeft = 4.41;
        private const double DefaultListLevel6BulletIndentLeft = 5.04;
        private const double DefaultListLevel7BulletIndentLeft = 5.68;
        private const double DefaultListLevel8BulletIndentLeft = 6.31;
        private const double DefaultListLevel9BulletIndentLeft = 6.95;
        private const double DefaultListLevel1BulletIndentRight = 0;
        private const double DefaultListLevel2BulletIndentRight = 0;
        private const double DefaultListLevel3BulletIndentRight = 0;
        private const double DefaultListLevel4BulletIndentRight = 0;
        private const double DefaultListLevel5BulletIndentRight = 0;
        private const double DefaultListLevel6BulletIndentRight = 0;
        private const double DefaultListLevel7BulletIndentRight = 0;
        private const double DefaultListLevel8BulletIndentRight = 0;
        private const double DefaultListLevel9BulletIndentRight = 0;
        private const double DefaultListLevel1Indent = 0.64;
        private const double DefaultListLevel2Indent = 0.76;
        private const double DefaultListLevel3Indent = 0.89;
        private const double DefaultListLevel4Indent = 1.14;
        private const double DefaultListLevel5Indent = 1.4;
        private const double DefaultListLevel6Indent = 1.65;
        private const double DefaultListLevel7Indent = 1.91;
        private const double DefaultListLevel8Indent = 2.16;
        private const double DefaultListLevel9Indent = 2.54;
        private const string DefaultListLevel1NumberFormat = "1.";
        private const string DefaultListLevel2NumberFormat = "1.1";
        private const string DefaultListLevel3NumberFormat = "1.1.1";
        private const string DefaultListLevel4NumberFormat = "1.1.1.1";
        private const string DefaultListLevel5NumberFormat = "1.1.1.1.1";
        private const string DefaultListLevel6NumberFormat = "1.1.1.1.1.1";
        private const string DefaultListLevel7NumberFormat = "1.1.1.1.1.1.1";
        private const string DefaultListLevel8NumberFormat = "1.1.1.1.1.1.1.1";
        private const string DefaultListLevel9NumberFormat = "1.1.1.1.1.1.1.1.1";
        private const string DefaultListLevel1IndentOrOutdent = "Выступ";
        private const string DefaultListLevel2IndentOrOutdent = "Выступ";
        private const string DefaultListLevel3IndentOrOutdent = "Выступ";
        private const string DefaultListLevel4IndentOrOutdent = "Выступ";
        private const string DefaultListLevel5IndentOrOutdent = "Выступ";
        private const string DefaultListLevel6IndentOrOutdent = "Выступ";
        private const string DefaultListLevel7IndentOrOutdent = "Выступ";
        private const string DefaultListLevel8IndentOrOutdent = "Выступ";
        private const string DefaultListLevel9IndentOrOutdent = "Выступ";

        /// <summary>
        /// Вспомогательный метод для получения требуемого типа отступа в списках
        /// </summary>
        /// <param name="gost"></param>
        /// <param name="level"></param>
        /// <returns></returns>
        private string GetRequiredIndentType(Gost gost, int level)
        {
            return level switch
            {
                1 => gost.ListLevel1IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                2 => gost.ListLevel2IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                3 => gost.ListLevel3IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                4 => gost.ListLevel4IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                5 => gost.ListLevel5IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                6 => gost.ListLevel6IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                7 => gost.ListLevel7IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                8 => gost.ListLevel8IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                9 => gost.ListLevel9IndentOrOutdent ?? gost.ListLevel1IndentOrOutdent,
                _ => gost.ListLevel1IndentOrOutdent
            };
        }

        /// <summary>
        /// Вспомогательный метод для получения требуемого левого отступа в списках
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        private double GetListLevelIndentLeft(int level)
        {
            switch (level)
            {
                case 1: return DefaultListLevel1BulletIndentLeft;
                case 2: return DefaultListLevel2BulletIndentLeft;
                case 3: return DefaultListLevel3BulletIndentLeft;
                case 4: return DefaultListLevel4BulletIndentLeft;
                case 5: return DefaultListLevel5BulletIndentLeft;
                case 6: return DefaultListLevel6BulletIndentLeft;
                case 7: return DefaultListLevel7BulletIndentLeft;
                case 8: return DefaultListLevel8BulletIndentLeft;
                case 9: return DefaultListLevel9BulletIndentLeft;
                default: return DefaultListLevel1BulletIndentLeft; // по умолчанию уровень 1
            }
        }

        /// <summary>
        /// Вспомогательный метод для получения требуемого правого отступа в списках
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        private double GetListLevelIndentRight(int level)
        {
            switch (level)
            {
                case 1: return DefaultListLevel1BulletIndentRight;
                case 2: return DefaultListLevel2BulletIndentRight;
                case 3: return DefaultListLevel3BulletIndentRight;
                case 4: return DefaultListLevel4BulletIndentRight;
                case 5: return DefaultListLevel5BulletIndentRight;
                case 6: return DefaultListLevel6BulletIndentRight;
                case 7: return DefaultListLevel7BulletIndentRight;
                case 8: return DefaultListLevel8BulletIndentRight;
                case 9: return DefaultListLevel9BulletIndentRight;
                default: return DefaultListLevel1BulletIndentRight; // по умолчанию уровень 1
            }
        }

        /// <summary>
        /// Вспомогательный метод для получения требуемого отступа первой строки в списках
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        private double GetListLevelIndent(int level)
        {
            switch (level)
            {
                case 1: return DefaultListLevel1Indent;
                case 2: return DefaultListLevel2Indent;
                case 3: return DefaultListLevel3Indent;
                case 4: return DefaultListLevel4Indent;
                case 5: return DefaultListLevel5Indent;
                case 6: return DefaultListLevel6Indent;
                case 7: return DefaultListLevel7Indent;
                case 8: return DefaultListLevel8Indent;
                case 9: return DefaultListLevel9Indent;
                default: return DefaultListLevel1Indent; // по умолчанию уровень 1
            }
        }

        private readonly Func<Paragraph, Gost, bool> _isAdditionalHeader;
        private readonly Gost _gost;

        public CheckListDocText(Gost gost, Func<Paragraph, Gost, bool> isAdditionalHeader)
        {
            _gost = gost;
            _isAdditionalHeader = isAdditionalHeader;
        }

        /// <summary>
        /// Проверка базовых параметров списков
        /// </summary>
        /// <param name="paragraphs">Список параграфов для проверки</param>
        /// <param name="gost">Объект с параметрами ГОСТ</param>
        /// <param name="updateUI">Делегат для обновления UI</param>
        /// <returns>True если проверка пройдена успешно</returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckBulletedListsAsync(WordprocessingDocument doc, List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var errors = new List<TextErrorInfo>();
                bool hasErrors = false;
                var allStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                foreach (var paragraph in paragraphs)
                {
                    if (!IsListItem(paragraph)) continue;

                    // ============ ПРОПУСКАЕМ ПУСТЫЕ АБЗАЦЫ ============
                    if (string.IsNullOrWhiteSpace(paragraph.InnerText?.Trim()))
                        continue;

                    var errorDetails = new List<string>();
                    bool paragraphHasError = false;
                    var runsWithText = paragraph.Elements<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)).ToList();

                    // 1. Проверка выравнивания
                    if (!string.IsNullOrEmpty(gost.BulletAlignment))
                    {
                        var (actualAlignment, isAlignmentDefined) = GetActualAlignmentForList(paragraph, allStyles);

                        if (!isAlignmentDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить выравнивание");
                            paragraphHasError = true;
                        }
                        else if (actualAlignment != gost.BulletAlignment)
                        {
                            errorDetails.Add($"\n       • выравнивание: '{actualAlignment}' (требуется '{gost.BulletAlignment}')");
                            paragraphHasError = true;
                        }
                    }

                    // 2. Проверка формата нумерации
                    if (IsNumberedList(paragraph))
                    {
                        int level = GetListLevel(paragraph, gost);
                        string? requiredFormat = level switch
                        {
                            1 => gost.ListLevel1NumberFormat,
                            2 => gost.ListLevel2NumberFormat,
                            3 => gost.ListLevel3NumberFormat,
                            4 => gost.ListLevel4NumberFormat,
                            5 => gost.ListLevel5NumberFormat,
                            6 => gost.ListLevel6NumberFormat,
                            7 => gost.ListLevel7NumberFormat,
                            8 => gost.ListLevel8NumberFormat,
                            9 => gost.ListLevel9NumberFormat,
                            _ => null
                        };

                        if (!string.IsNullOrEmpty(requiredFormat))
                        {
                            var firstRunText = runsWithText.FirstOrDefault()?.InnerText?.Trim();
                            if (firstRunText != null && !CheckNumberFormat(firstRunText, requiredFormat))
                            {
                                errorDetails.Add($"Неверный формат нумерации уровня {level}: '{firstRunText}' (требуется '{requiredFormat}')");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // 3. Проверка типа шрифта
                    if (!string.IsNullOrEmpty(gost.BulletFontName))
                    {
                        foreach (var run in runsWithText)
                        {
                            var (actualFont, isFontDefined) = GetActualFontForList(run, paragraph, allStyles, doc);

                            if (!isFontDefined)
                            {
                                errorDetails.Add($"\n       • не удалось определить шрифт");
                                paragraphHasError = true;
                                break;
                            }
                            else if (!string.Equals(actualFont, gost.BulletFontName, StringComparison.OrdinalIgnoreCase))
                            {
                                errorDetails.Add($"\n       • шрифт: '{actualFont}' (требуется '{gost.BulletFontName}')");
                                paragraphHasError = true;
                                break;
                            }
                        }
                    }

                    // 4. Проверка размера шрифта
                    if (gost.BulletFontSize.HasValue)
                    {
                        foreach (var run in runsWithText)
                        {
                            var (actualSize, isSizeDefined) = GetActualFontSizeForList(run, paragraph, allStyles, doc);

                            if (!isSizeDefined)
                            {
                                errorDetails.Add($"\n       • не удалось определить размер шрифта");
                                paragraphHasError = true;
                                break;
                            }
                            else if (Math.Abs(actualSize - gost.BulletFontSize.Value) > 0.1)
                            {
                                errorDetails.Add($"\n       • размер шрифта: {actualSize:F1} pt (требуется {gost.BulletFontSize.Value:F1} pt)");
                                paragraphHasError = true;
                                break;
                            }
                        }
                    }

                    if (paragraphHasError)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"Элемент списка '{GetShortText(paragraph)}': {string.Join(", ", errorDetails)}",
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
                        var msg = $"Ошибки в списках:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3)
                            msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Avalonia.Media.Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Списки соответствуют ГОСТу", Avalonia.Media.Brushes.Green);
                    }
                });

                return (!hasErrors, errors);
            });
        }

        /// <summary>
        /// Проверка соответствия интервалов в списках требованиям ГОСТа
        /// </summary>
        /// <param name="paragraphs">Список параграфов для проверки</param>
        /// <param name="gost">Объект с параметрами ГОСТ</param>
        /// <param name="updateUI">Делегат для обновления UI</param>
        /// <returns>Кортеж с результатом проверки и списком ошибок</returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckListParagraphSpacingAsync(WordprocessingDocument doc, List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var errors = new List<TextErrorInfo>();
                bool hasErrors = false;

                // Получаем все стили документа
                var allStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>() ?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                foreach (var paragraph in paragraphs)
                {
                    if (!IsListItem(paragraph)) continue;

                    if (string.IsNullOrWhiteSpace(paragraph.InnerText?.Trim()))
                        continue;

                    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    var paragraphStyle = styleId != null && allStyles.TryGetValue(styleId, out var style) ? style : null;

                    bool paragraphHasError = false;
                    var errorDetails = new List<string>();

                    // 1. Проверка междустрочного интервала
                    if (gost.BulletLineSpacingValue.HasValue || !string.IsNullOrEmpty(gost.BulletLineSpacingType))
                    {
                        var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacingForList(paragraph, paragraphStyle, allStyles);

                        if (!isSpacingDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить междустрочный интервал");
                            paragraphHasError = true;
                        }
                        else
                        {
                            // Проверка типа интервала
                            if (!string.IsNullOrEmpty(gost.BulletLineSpacingType))
                            {
                                string requiredType = gost.BulletLineSpacingType ?? DefaultListLineSpacingType;
                                if (actualSpacingType != requiredType)
                                {
                                    errorDetails.Add($"\n       • тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                                    paragraphHasError = true;
                                }
                            }

                            // Проверка значения интервала
                            if (gost.BulletLineSpacingValue.HasValue && Math.Abs(actualSpacingValue - gost.BulletLineSpacingValue.Value) > 0.01)
                            {
                                errorDetails.Add($"\n       • межстрочный интервал: {actualSpacingValue:F2} (требуется {gost.BulletLineSpacingValue.Value:F2})");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // 2. Проверка интервалов перед/после
                    if (gost.BulletLineSpacingBefore.HasValue || gost.BulletLineSpacingAfter.HasValue)
                    {
                        var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacingForList(paragraph, paragraphStyle, allStyles);

                        if (!isSpacingDefined)
                        {
                            errorDetails.Add($"\n       • не удалось определить интервалы перед/после");
                            paragraphHasError = true;
                        }
                        else
                        {
                            if (gost.BulletLineSpacingBefore.HasValue && Math.Abs(actualBefore - gost.BulletLineSpacingBefore.Value) > 0.1)
                            {
                                errorDetails.Add($"\n       • интервал перед: {actualBefore:F1} pt (требуется {gost.BulletLineSpacingBefore.Value:F1} pt)");
                                paragraphHasError = true;
                            }

                            if (gost.BulletLineSpacingAfter.HasValue && Math.Abs(actualAfter - gost.BulletLineSpacingAfter.Value) > 0.1)
                            {
                                errorDetails.Add($"\n       • интервал после: {actualAfter:F1} pt (требуется {gost.BulletLineSpacingAfter.Value:F1} pt)");
                                paragraphHasError = true;
                            }
                        }
                    }

                    if (paragraphHasError)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"\n       • Элемент списка '{GetShortText(paragraph)}': {string.Join(", ", errorDetails)}",
                            ProblemParagraph = paragraph,
                            ProblemRun = null
                        });
                        hasErrors = true;
                    }
                }

                // Вывод результатов (оставьте ваш существующий код)
                Dispatcher.UIThread.Post(() =>
                {
                    if (hasErrors)
                    {
                        var msg = $"Ошибки в интервалах списков:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3) msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Интервалы списков соответствуют ГОСТу", Brushes.Green);
                    }
                });

                return (!hasErrors, errors);
            });
        }

        /// <summary>
        /// Проверка отступов списков
        /// </summary>
        /// <param name="paragraphs">Список параграфов для проверки</param>
        /// <param name="gost">Объект с параметрами ГОСТ</param>
        /// <param name="updateUI">Делегат для обновления UI</param>
        /// <returns>Кортеж с результатом проверки и списком ошибок</returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckListIndentsAsync(List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                bool hasIndentRequirements = gost.ListLevel1Indent.HasValue || gost.ListLevel2Indent.HasValue || gost.ListLevel3Indent.HasValue || gost.ListLevel4Indent.HasValue ||
                gost.ListLevel5Indent.HasValue || gost.ListLevel6Indent.HasValue || gost.ListLevel7Indent.HasValue || gost.ListLevel8Indent.HasValue || gost.ListLevel9Indent.HasValue;

                bool hasLeftIndentRequirements = gost.ListLevel1BulletIndentLeft.HasValue || gost.ListLevel2BulletIndentLeft.HasValue || gost.ListLevel3BulletIndentLeft.HasValue ||
                gost.ListLevel4BulletIndentLeft.HasValue || gost.ListLevel5BulletIndentLeft.HasValue || gost.ListLevel6BulletIndentLeft.HasValue || gost.ListLevel7BulletIndentLeft.HasValue ||
                gost.ListLevel8BulletIndentLeft.HasValue || gost.ListLevel9BulletIndentLeft.HasValue;

                bool hasRightIndentRequirements = gost.ListLevel1BulletIndentRight.HasValue || gost.ListLevel2BulletIndentRight.HasValue || gost.ListLevel3BulletIndentRight.HasValue ||
                gost.ListLevel4BulletIndentRight.HasValue || gost.ListLevel5BulletIndentRight.HasValue || gost.ListLevel6BulletIndentRight.HasValue || gost.ListLevel7BulletIndentRight.HasValue ||
                gost.ListLevel8BulletIndentRight.HasValue || gost.ListLevel9BulletIndentRight.HasValue;

                if (!hasIndentRequirements && !hasLeftIndentRequirements && !hasRightIndentRequirements)
                {
                    Dispatcher.UIThread.Post(() =>
                    {
                        updateUI?.Invoke("Проверка отступов списков не требуется", Avalonia.Media.Brushes.Gray);
                    });
                    return (true, new List<TextErrorInfo>());
                }

                bool hasErrors = false;
                var errors = new List<TextErrorInfo>();

                foreach (var paragraph in paragraphs)
                {
                    if (!IsStrictListItem(paragraph)) continue;

                    // ============ ПРОПУСКАЕМ ПУСТЫЕ АБЗАЦЫ ============
                    if (string.IsNullOrWhiteSpace(paragraph.InnerText?.Trim()))
                        continue;

                    int level = GetListLevel(paragraph, gost);
                    var indent = paragraph.ParagraphProperties?.Indentation;
                    bool paragraphHasError = false;
                    var errorDetails = new List<string>();

                    // 1. Получаем ТРЕБУЕМЫЕ значения из ГОСТа для текущего уровня
                    double? gostRequiredIndent = level switch
                    {
                        1 => gost.ListLevel1Indent,
                        2 => gost.ListLevel2Indent,
                        3 => gost.ListLevel3Indent,
                        4 => gost.ListLevel4Indent,
                        5 => gost.ListLevel5Indent,
                        6 => gost.ListLevel6Indent,
                        7 => gost.ListLevel7Indent,
                        8 => gost.ListLevel8Indent,
                        9 => gost.ListLevel9Indent,
                        _ => null
                    };

                    // Получаем требуемый отступ слева для текущего уровня
                    double? gostRequiredLeftIndent = level switch
                    {
                        1 => gost.ListLevel1BulletIndentLeft,
                        2 => gost.ListLevel2BulletIndentLeft,
                        3 => gost.ListLevel3BulletIndentLeft,
                        4 => gost.ListLevel4BulletIndentLeft,
                        5 => gost.ListLevel5BulletIndentLeft,
                        6 => gost.ListLevel6BulletIndentLeft,
                        7 => gost.ListLevel7BulletIndentLeft,
                        8 => gost.ListLevel8BulletIndentLeft,
                        9 => gost.ListLevel9BulletIndentLeft,
                        _ => null
                    };

                    // Получаем требуемый отступ справа для текущего уровня
                    double? gostRequiredRightIndent = level switch
                    {
                        1 => gost.ListLevel1BulletIndentRight,
                        2 => gost.ListLevel2BulletIndentRight,
                        3 => gost.ListLevel3BulletIndentRight,
                        4 => gost.ListLevel4BulletIndentRight,
                        5 => gost.ListLevel5BulletIndentRight,
                        6 => gost.ListLevel6BulletIndentRight,
                        7 => gost.ListLevel7BulletIndentRight,
                        8 => gost.ListLevel8BulletIndentRight,
                        9 => gost.ListLevel9BulletIndentRight,
                        _ => null
                    };

                    // Если для уровня нет специфичного требования, используем общее значение
                    gostRequiredIndent ??= gost.ListLevel1Indent;
                    gostRequiredLeftIndent ??= gost.ListLevel1BulletIndentLeft;
                    gostRequiredRightIndent ??= gost.ListLevel1BulletIndentRight;

                    // 2. Получаем ФАКТИЧЕСКИЕ значения из документа
                    double? leftIndentValue = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : null;
                    double? firstLineIndentValue = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : null;
                    double? hangingIndentValue = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : null;
                    double? rightIndentValue = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : null;

                    // Определяем текущий тип и значение отступа
                    string currentType = "Нет";
                    double? currentValue = null;

                    if (hangingIndentValue != null && hangingIndentValue != 0)
                    {
                        currentType = "Выступ";
                        currentValue = hangingIndentValue;
                    }
                    else if (firstLineIndentValue != null && firstLineIndentValue != 0)
                    {
                        currentType = "Отступ";
                        currentValue = firstLineIndentValue;
                    }
                    else // Если в документе не заданы отступы, используем значения по умолчанию
                    {
                        currentType = GetRequiredIndentType(gost, level);
                        currentValue = GetListLevelIndent(level);
                    }

                    // 3. Проверяем только если в ГОСТе есть требования для отступов
                    if (gostRequiredIndent.HasValue)
                    {
                        string requiredType = GetRequiredIndentType(gost, level);

                        // Проверка типа отступа
                        if (!string.IsNullOrEmpty(requiredType))
                        {
                            bool typeMatches = string.Equals(currentType, requiredType, StringComparison.OrdinalIgnoreCase);
                            if (!typeMatches)
                            {
                                errorDetails.Add($"\n       • тип первой строки: {currentType} (требуется {requiredType})");
                                paragraphHasError = true;
                            }
                        }

                        // Проверка значения отступа
                        if (currentValue.HasValue)
                        {
                            double deviation = Math.Abs(currentValue.Value - gostRequiredIndent.Value);
                            if (deviation > 0.05)
                            {
                                errorDetails.Add($"\n       • {currentType} первой строки: {currentValue.Value:F2} см (требуется {gostRequiredIndent.Value:F2} см)");
                                paragraphHasError = true;
                            }
                        }
                    }

                    // 4. Проверка левого отступа (теперь для каждого уровня)
                    if (gostRequiredLeftIndent.HasValue)
                    {
                        double actualLeft = leftIndentValue ?? GetListLevelIndentLeft(level);
                        double requiredLeft = gostRequiredLeftIndent.Value;

                        if (Math.Abs(actualLeft - requiredLeft) > 0.05)
                        {
                            errorDetails.Add($"\n       • Левый отступ: {actualLeft:F2} см (требуется {requiredLeft:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    // 5. Проверка правого отступа (теперь для каждого уровня)
                    if (gostRequiredRightIndent.HasValue)
                    {
                        double actualRight = rightIndentValue ?? GetListLevelIndentRight(level);
                        double requiredRight = gostRequiredRightIndent.Value;

                        if (Math.Abs(actualRight - requiredRight) > 0.05)
                        {
                            errorDetails.Add($"\n       • Правый отступ: {actualRight:F2} см (требуется {requiredRight:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    if (paragraphHasError)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"\n       • Список ур. {level} '{GetShortText(paragraph)}': {string.Join(", ", errorDetails)}",
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
                        var msg = $"Ошибки в отступах списков:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3)
                            msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Avalonia.Media.Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Отступы списков соответствуют ГОСТу", Avalonia.Media.Brushes.Green);
                    }
                });

                return (!hasErrors, errors);
            });
        }

        private (string Type, double Value, bool IsDefined) GetActualLineSpacingForList(Paragraph paragraph, Style paragraphStyle, Dictionary<string, Style> allStyles)
        {
            // 1. Явные свойства
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            var parsed = ParseLineSpacing(spacing);
            if (parsed.IsDefined) return parsed;

            // 2. Стиль абзаца и его родители
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var currentStyle))
            {
                while (currentStyle != null)
                {
                    var styleSpacing = currentStyle.StyleParagraphProperties?.SpacingBetweenLines;
                    parsed = ParseLineSpacing(styleSpacing);
                    if (parsed.IsDefined) return parsed;

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null &&
                                 allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle)
                        ? basedOnStyle
                        : null;
                }
            }

            // 3. Проверяем переданный стиль (если есть)
            if (paragraphStyle != null)
            {
                var styleSpacing = paragraphStyle.StyleParagraphProperties?.SpacingBetweenLines;
                parsed = ParseLineSpacing(styleSpacing);
                if (parsed.IsDefined) return parsed;
            }

            // 4. Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle))
            {
                parsed = ParseLineSpacing(normalStyle.StyleParagraphProperties?.SpacingBetweenLines);
                if (parsed.IsDefined) return parsed;
            }

            // 5. Стандартные значения для списков
            return (DefaultListLineSpacingType, DefaultListLineSpacingValue, true);
        }

        private (double Before, double After, bool IsDefined) GetActualParagraphSpacingForList(Paragraph paragraph, Style paragraphStyle, Dictionary<string, Style> allStyles)
        {
            double? before = null;
            double? after = null;

            // 1. Явные свойства
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing?.Before?.Value != null) before = ConvertTwipsToPoints(spacing.Before.Value);
            if (spacing?.After?.Value != null) after = ConvertTwipsToPoints(spacing.After.Value);

            // 2. Поиск в стилях
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

            // 3. Normal стиль
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

            return (before ?? DefaultListSpacingBefore, after ?? DefaultListSpacingAfter, true);
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

        private (string Alignment, bool IsDefined) GetActualAlignmentForList(Paragraph paragraph, Dictionary<string, Style> allStyles)
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
            if (allStyles.TryGetValue("Normal", out var normalStyle) && normalStyle.StyleParagraphProperties?.Justification?.Val?.Value != null)
            {
                return (GetAlignmentString(normalStyle.StyleParagraphProperties.Justification), true);
            }

            // 4. Если нигде не задано - считаем Left по умолчанию
            return ("Left", true);
        }

        private (string FontName, bool IsDefined) GetActualFontForList(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles, WordprocessingDocument doc)
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

            // 3. Проверяем стиль списка (если есть)
            var listStyle = GetListStyle(paragraph, doc); 
            if (listStyle != null)
            {
                var listStyleFont = GetStyleFont(listStyle);
                if (!string.IsNullOrEmpty(listStyleFont))
                    return (listStyleFont, true);
            }

            // 4. Проверяем стиль Paragraph с учетом наследования
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null)
            {
                var currentStyle = allStyles.TryGetValue(paraStyleId, out var style) ? style : null;
                while (currentStyle != null)
                {
                    var font = GetStyleFont(currentStyle);
                    if (!string.IsNullOrEmpty(font))
                        return (font, true);

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle)
                        ? basedOnStyle
                        : null;
                }
            }

            // 5. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle))
            {
                var font = GetStyleFont(normalStyle);
                if (!string.IsNullOrEmpty(font))
                    return (font, true);
            }

            return (DefaultListFont, false);
        }

        private (double Size, bool IsDefined) GetActualFontSizeForList(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles, WordprocessingDocument doc)
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

            // 3. Проверяем стиль списка (если есть)
            var listStyle = GetListStyle(paragraph, doc);
            if (listStyle != null && listStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
            {
                return (double.Parse(listStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);
            }

            // 4. Проверяем стиль Paragraph с учетом наследования
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null && allStyles.TryGetValue(paraStyleId, out var paraStyle))
            {
                var currentStyle = paraStyle;
                while (currentStyle != null)
                {
                    if (currentStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
                        return (double.Parse(currentStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle)
                        ? basedOnStyle
                        : null;
                }
            }

            // 5. Проверяем Normal стиль
            if (allStyles.TryGetValue("Normal", out var normalStyle) && normalStyle.StyleRunProperties?.FontSize?.Val?.Value != null)
            {
                return (double.Parse(normalStyle.StyleRunProperties.FontSize.Val.Value) / 2, true);
            }

            return (DefaultListSize, false);
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

        private Style GetListStyle(Paragraph paragraph, WordprocessingDocument doc)
        {
            if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val == null)
                return null;

            var styleId = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value;

            // Проверяем, относится ли стиль к спискам
            if (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering"))
            {
                return doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            }

            return null;
        }

        /// <summary>
        /// Строгая проверка определяющая что параграф является элементом списка
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsStrictListItem(Paragraph paragraph)
        {
            // 1. Сначала проверяем, не является ли это заголовком
            if (_isAdditionalHeader != null && _isAdditionalHeader(paragraph, _gost))
                return false;

            // 2. Проверяем явные свойства нумерации
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // 3. Проверяем стили списка
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) &&
                (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering")))
                return true;

            // 4. Проверяем по содержимому (маркеры или нумерация)
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                string text = firstRun.InnerText?.Trim() ?? "";

                // Маркированные списки
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—"))
                    return true;

                // Нумерованные списки (более строгая проверка)
                if (Regex.IsMatch(text, @"^\d+[\.\)]\s") ||   // "1. Текст" или "1) Текст"
                    Regex.IsMatch(text, @"^[a-z]\)\s") ||     // "a) Текст"
                    Regex.IsMatch(text, @"^[IVXLCDM]+\.\s", RegexOptions.IgnoreCase))  // "I. Текст"
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверяет, является ли параграф элементом списка
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsListItem(Paragraph paragraph)
        {
            // Используем строгую версию, но без проверки стилей
            if (_isAdditionalHeader != null && _isAdditionalHeader(paragraph, _gost))
                return false;

            // Проверяем свойства нумерации
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // Проверяем по форматированию
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                var text = firstRun.InnerText.Trim();

                // Маркированные списки
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—"))
                    return true;

                // Нумерованные списки (более строгая проверка)
                if (Regex.IsMatch(text, @"^\d+[\.\)]\s") || Regex.IsMatch(text, @"^[a-z]\)\s"))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверяет соответствие формата нумерации требуемому
        /// </summary>
        /// <param name="text"></param>
        /// <param name="requiredFormat"></param>
        /// <returns></returns>
        private bool CheckNumberFormat(string text, string requiredFormat)
        {
            if (requiredFormat.EndsWith(".") && text.EndsWith("."))
                return true;
            if (requiredFormat.EndsWith(")") && text.EndsWith(")"))
                return true;
            return false;
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
        /// Конвертирует twips в сантиметры (1 см = 567 twips)
        /// </summary>
        /// <param name="twips"></param>
        /// <returns></returns>
        private double TwipsToCm(double twips) => twips / 567.0;

        /// <summary>
        /// Вспомогательный метод для получения сокращенного текста
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string GetShortText(Paragraph paragraph)
        {
            string text = paragraph.InnerText.Trim();
            return GetShortText2(text);
        }

        private string GetShortText2(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "[пустой элемент]";

            return text.Length > 30 ? text.Substring(0, 27) + "..." : text;
        }

        /// <summary>
        /// Определяет уровень вложенности списка (1-9)
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="gost"></param>
        /// <returns></returns> 
        private int GetListLevel(Paragraph paragraph, Gost gost)
        {
            var numberingProps = paragraph.ParagraphProperties?.NumberingProperties;

            if (numberingProps?.NumberingLevelReference?.Val?.Value != null)
            {
                return numberingProps.NumberingLevelReference.Val.Value + 1;
            }

            var indent = paragraph.ParagraphProperties?.Indentation;
            if (indent?.Left != null)
            {
                double leftIndent = double.Parse(indent.Left.Value) / 567.0; // в см

                if (gost.ListLevel3Indent.HasValue && leftIndent >= gost.ListLevel3Indent.Value - 0.5)
                    return 3;
                if (gost.ListLevel2Indent.HasValue && leftIndent >= gost.ListLevel2Indent.Value - 0.5)
                    return 2;
            }

            return 1; // По умолчанию 
        }

        /// <summary>
        /// Проверяет является ли параграф нумерованным списком по формату первого символа
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsNumberedList(Paragraph paragraph)
        {
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun == null) return false;

            var text = firstRun.InnerText.Trim();

            return Regex.IsMatch(text, @"^(\d+[\.\)]|[a-z]\)|[A-Z]\.|I+\.|V+\.|X+\.)");// Форматы нумерации: 1., 1), a., a), I., и т.д.
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
