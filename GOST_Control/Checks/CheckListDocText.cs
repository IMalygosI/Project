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
using DocumentFormat.OpenXml.Wordprocessing;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок списков
    /// </summary>
    public class CheckListDocText
    {

        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ОГЛАВЛЕНИЯ =======================
        private const double DefaultTocSize = 11.0;
        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ СПИСКОВ =======================
        private const string DefaultListFont = "Arial";
        private const double DefaultListSize = 11.0;
        private const string DefaultListLineSpacingType = "Множитель";
        private const double DefaultListLineSpacingValue = 1.15;
        private const double DefaultListSpacingBefore = 0.0;
        private const double DefaultListSpacingAfter = 0.35;
        private const string DefaultListFirstLineType = "Выступ";
        private const double DefaultListHangingIndent = 0.64;
        private const double DefaultListLeftIndent = 0.62;
        private const double DefaultListRightIndent = 0.0;
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

        /// <summary>
        /// Проверка базовых параметров списков
        /// </summary>
        /// <param name="paragraphs">Список параграфов для проверки</param>
        /// <param name="gost">Объект с параметрами ГОСТ</param>
        /// <param name="updateUI">Делегат для обновления UI</param>
        /// <returns>True если проверка пройдена успешно</returns>
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckBulletedListsAsync(List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var errors = new List<TextErrorInfo>();
                bool hasErrors = false;

                foreach (var paragraph in paragraphs)
                {
                    if (!IsListItem(paragraph)) continue;

                    var errorDetails = new List<string>();
                    bool paragraphHasError = false;
                    var runsWithText = paragraph.Elements<Run>().Where(r => !string.IsNullOrWhiteSpace(r.InnerText)).ToList();

                    // Проверка выравнивания
                    if (!string.IsNullOrEmpty(gost.BulletAlignment))
                    {
                        var currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification) ?? BulletAlignment;
                        if (currentAlignment != gost.BulletAlignment)
                        {
                            errorDetails.Add($"выравнивание: '{currentAlignment}' (требуется '{gost.BulletAlignment}')");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка формата нумерации
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

                    // Проверка типа шрифта
                    if (!string.IsNullOrEmpty(gost.BulletFontName))
                    {
                        foreach (var run in runsWithText)
                        {
                            var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? DefaultListFont;
                            if (font != null && font != gost.BulletFontName)
                            {
                                errorDetails.Add($"Неверный шрифт списка: '{font}' (требуется '{gost.BulletFontName}')");
                                paragraphHasError = true;
                                break;
                            }
                        }
                    }

                    // Проверка размера шрифта
                    if (gost.BulletFontSize.HasValue)
                    {
                        foreach (var run in runsWithText)
                        {
                            var fontSize = run.RunProperties?.FontSize?.Val?.Value ?? DefaultListSize.ToString();

                            if (fontSize != null)
                            {
                                double actualSize = -1;
                                var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value;

                                if (fontSizeVal != null)
                                {
                                    actualSize = double.Parse(fontSizeVal) / 2;
                                }
                                else
                                {
                                    actualSize = DefaultTocSize;
                                }

                                if (Math.Abs(actualSize - gost.BulletFontSize.Value) > 0.1)
                                {
                                    errorDetails.Add($"Неверный размер шрифта: {actualSize}pt (требуется {gost.BulletFontSize.Value}pt) в параграфе: '{GetShortText(paragraph)}'");
                                    paragraphHasError = true;
                                    break;
                                }
                            }
                            else if (gost.BulletFontSize.Value != 0) // 0 - значение по умолчанию
                            {
                                errorDetails.Add("Отсутствует размер шрифта");
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
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckListParagraphSpacingAsync(List<Paragraph> paragraphs, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var lineSpacingTypeNames = new Dictionary<LineSpacingRuleValues, string>
        {
            { LineSpacingRuleValues.Auto, "Множитель" },
            { LineSpacingRuleValues.AtLeast, "Минимум" },
            { LineSpacingRuleValues.Exact, "Точно" }
        };

                bool hasErrors = false;
                var errors = new List<TextErrorInfo>();

                foreach (var paragraph in paragraphs)
                {
                    if (!IsListItem(paragraph)) continue;

                    var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
                    bool paragraphHasError = false;
                    var errorDetails = new List<string>();

                    // Проверка межстрочного интервала
                    if (gost.BulletLineSpacingValue.HasValue)
                    {
                        double actualSpacing = DefaultListLineSpacingValue;
                        string actualSpacingType = DefaultListLineSpacingType;
                        LineSpacingRuleValues? actualRule = LineSpacingRuleValues.Auto;

                        if (spacing?.Line != null)
                        {
                            if (spacing.LineRule?.Value == LineSpacingRuleValues.Exact)
                            {
                                actualSpacing = double.Parse(spacing.Line.Value) / 567.0;
                                actualSpacingType = "Точно";
                                actualRule = LineSpacingRuleValues.Exact;
                            }
                            else if (spacing.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                            {
                                actualSpacing = double.Parse(spacing.Line.Value) / 567.0;
                                actualSpacingType = "Минимум";
                                actualRule = LineSpacingRuleValues.AtLeast;
                            }
                            else
                            {
                                actualSpacing = double.Parse(spacing.Line.Value) / 240.0;
                                actualSpacingType = "Множитель";
                                actualRule = LineSpacingRuleValues.Auto;
                            }
                        }

                        // Определяем требуемый тип интервала
                        LineSpacingRuleValues requiredRule = (gost.BulletLineSpacingType ?? DefaultListLineSpacingType) switch
                        {
                            "Минимум" => LineSpacingRuleValues.AtLeast,
                            "Точно" => LineSpacingRuleValues.Exact,
                            _ => LineSpacingRuleValues.Auto
                        };

                        string requiredType = gost.BulletLineSpacingType ?? DefaultListLineSpacingType;

                        // Проверка типа интервала
                        if (actualRule != requiredRule)
                        {
                            errorDetails.Add($"тип интервала: '{actualSpacingType}' (требуется '{requiredType}')");
                            paragraphHasError = true;
                        }

                        // Проверка значения интервала
                        double requiredSpacingValue = gost.BulletLineSpacingValue ?? DefaultListLineSpacingValue;
                        if (Math.Abs(actualSpacing - requiredSpacingValue) > 0.01)
                        {
                            errorDetails.Add($"межстрочный интервал: {actualSpacing:F2} (требуется {requiredSpacingValue:F2})");
                            paragraphHasError = true;
                        }
                    }

                    // Проверка интервалов перед/после
                    if (gost.BulletLineSpacingBefore.HasValue || gost.BulletLineSpacingAfter.HasValue)
                    {
                        double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultListSpacingBefore;
                        double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultListSpacingAfter;

                        if (gost.BulletLineSpacingBefore.HasValue && Math.Abs(actualBefore - gost.BulletLineSpacingBefore.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал перед: {actualBefore:F1} pt (требуется {gost.BulletLineSpacingBefore.Value:F1} pt)");
                            paragraphHasError = true;
                        }

                        if (gost.BulletLineSpacingAfter.HasValue && Math.Abs(actualAfter - gost.BulletLineSpacingAfter.Value) > 0.01)
                        {
                            errorDetails.Add($"интервал после: {actualAfter:F1} pt (требуется {gost.BulletLineSpacingAfter.Value:F1} pt)");
                            paragraphHasError = true;
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
                        var msg = $"Ошибки в интервалах списков:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(3))}";
                        if (errors.Count > 3)
                            msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Avalonia.Media.Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke("Интервалы списков соответствуют ГОСТу", Avalonia.Media.Brushes.Green);
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

                    if (hangingIndentValue != null)
                    {
                        currentType = "Выступ";
                        currentValue = hangingIndentValue;
                    }
                    else if (firstLineIndentValue != null)
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
                                errorDetails.Add($"тип первой строки: {currentType} (требуется {requiredType})");
                                paragraphHasError = true;
                            }
                        }

                        // Проверка значения отступа
                        if (currentValue.HasValue)
                        {
                            double deviation = Math.Abs(currentValue.Value - gostRequiredIndent.Value);
                            if (deviation > 0.05)
                            {
                                errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {gostRequiredIndent.Value:F2} см)");
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
                            errorDetails.Add($"Левый отступ: {actualLeft:F2} см (требуется {requiredLeft:F2} см)");
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
                            errorDetails.Add($"Правый отступ: {actualRight:F2} см (требуется {requiredRight:F2} см)");
                            paragraphHasError = true;
                        }
                    }

                    if (paragraphHasError)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"Список ур. {level} '{GetShortText(paragraph)}': {string.Join(", ", errorDetails)}",
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

        /// <summary>
        /// Строгая проверка определяющая что параграф является элементом списка
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private bool IsStrictListItem(Paragraph paragraph)
        {
            // 1. Проверка явных свойств нумерации
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // 2. Проверка стилей списка
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) && (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering")))
                return true;

            // 3. Проверка по содержимому (маркеры или нумерация)
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                string text = firstRun.InnerText?.Trim() ?? "";

                // Маркированные списки
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—"))
                    return true;

                // Нумерованные списки
                if (Regex.IsMatch(text, @"^\d+[\.\)]") ||   // 1. 1) 
                    Regex.IsMatch(text, @"^[a-z]\)") ||     // a) b)
                    Regex.IsMatch(text, @"^[IVXLCDM]+\.", RegexOptions.IgnoreCase))  // I. II.
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Проверяет, является ли параграф элементом списка
        /// </summary>ё
        private bool IsListItem(Paragraph paragraph)
        {
            // 1. Проверка нумерации
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
                return true;

            // 2. Проверка стиля списка
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId) && (styleId.Contains("List") || styleId.Contains("Bullet") || styleId.Contains("Numbering")))
                return true;

            // 3. Проверка по форматированию
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                var text = firstRun.InnerText.Trim();

                // Маркированные списки
                if (text.StartsWith("•") || text.StartsWith("-") || text.StartsWith("—"))
                    return true;

                // Нумерованные списки
                if (Regex.IsMatch(text, @"^\d+[\.\)]") || Regex.IsMatch(text, @"^[a-z]\)"))
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
