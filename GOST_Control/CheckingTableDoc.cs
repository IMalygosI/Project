using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avalonia.Media;
using Avalonia.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace GOST_Control
{
    public class CheckingTableDoc
    {
        // ======================= СТАНДАРТНЫЕ ЗНАЧЕНИЯ ДЛЯ ПОДПИСЕЙ К ТАБЛИЦАМ =======================
        private const string DefaultTableCaptionFont = "Arial";  // Стандартный шрифт подписей к таблицам
        private const double DefaultTableCaptionFontSize = 11.0; // Стандартный размер шрифта
        private const string DefaultTableCaptionIndentOrOutdent = "Нет"; // Тип первой строки — "Отступ" или "Выступ"
        private const double DefaultTableCaptionFirstLineIndent = 1.25; // Отступ первой строки подписи (в см)
        private const double DefaultTableCaptionIndentLeft = 0.0; // Левый отступ подписи
        private const double DefaultTableCaptionIndentRight = 0.0; // Правый отступ подписи
        private const string DefaultTableCaptionAlignment = "Left"; // Выравнивание подписи
        private const string DefaultTableCaptionLineSpacingType = "Множитель"; // Тип межстрочного интервала (например, "Множитель")
        private const double DefaultTableCaptionLineSpacingValue = 1.15; // Значение межстрочного интервала
        private const double DefaultTableCaptionLineSpacingBefore = 0.0; // Интервал перед подписью
        private const double DefaultTableCaptionLineSpacingAfter = 0.35; // Интервал после подписи

        // ======================= ТЕКСТ В ТАБЛИЦЕ =======================
        private const double DefaultTableFontSize = 11.0; // Стандартный размер шрифта
        private const string DefaultTableIndentOrOutdent = "Нет"; // Тип первой строки — "Отступ" или "Выступ"
        private const double DefaultTableFirstLineIndent = 1.25; // Отступ первой строки подписи (в см)
        private const double DefaultTableIndentLeft = 0.0; // Левый отступ подписи
        private const double DefaultTableIndentRight = 0.0; // Правый отступ подписи
        private const string DefaultTableAlignment = "Left"; // Выравнивание подписи
        private const string DefaultTableLineSpacingType = "Множитель"; // Тип межстрочного интервала (например, "Множитель")
        private const double DefaultTableLineSpacingValue = 1.15; // Значение межстрочного интервала
        private const double DefaultTableLineSpacingBefore = 0.0; // Интервал перед подписью
        private const double DefaultTableLineSpacingAfter = 0.35; // Интервал после подписи

        private readonly WordprocessingDocument _wordDoc;
        private readonly Gost _gost;
        private readonly Func<Run, bool> _shouldSkipRun;

        public CheckingTableDoc(WordprocessingDocument wordDoc, Gost gost, Func<Run, bool> shouldSkipRun)
        {
            _wordDoc = wordDoc;
            _gost = gost;
            _shouldSkipRun = shouldSkipRun;
        }

        /// <summary>
        /// Проверяет таблицы и их подписи на соответствие ГОСТу
        /// </summary>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<string> Errors)> CheckTablesAsync(Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var body = _wordDoc.MainDocumentPart.Document.Body;
                var tables = body.Elements<Table>().ToList();
                var errors = new List<string>();
                bool allTablesValid = true;

                foreach (var table in tables)
                {
                    bool tableValid = true;

                    // 1. Проверка подписи таблицы (должна быть над таблицей)
                    var tableCaption = GetTableCaption(table);

                    if (tableCaption == null)
                    {
                        continue;
                    }

                    // Проверка формата подписи
                    if (!CheckTableCaptionFormat(tableCaption, errors))
                    {
                        tableValid = false;
                    }

                    // Проверка стиля подписи таблицы
                    if (!CheckTableCaptionStyle(tableCaption, errors))
                    {
                        tableValid = false;
                    }

                    // Проверка содержимого таблицы
                    if (!CheckTableContent(table, errors))
                    {
                        tableValid = false;
                    }

                    if (!tableValid)
                    {
                        allTablesValid = false;
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (!allTablesValid)
                    {
                        var msg = $"Ошибки в таблицах:\n{string.Join("\n", errors.Take(3))}";
                        if (errors.Count > 3) msg += $"\n...и ещё {errors.Count - 3} ошибок";
                        updateUI?.Invoke(msg, Brushes.Red);
                    }
                    else
                    {
                        updateUI?.Invoke(tables.Any() ? "Все таблицы соответствуют ГОСТу" : "Таблицы не обнаружены - проверка не требуется", Brushes.Green);
                    }
                });

                return (allTablesValid, errors);
            });
        }

        /// <summary>
        /// Проверка формата подписи таблицы (например "Таблица 1 - Название")
        /// </summary>
        /// <param name="captionParagraph"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private bool CheckTableCaptionFormat(Paragraph captionParagraph, List<string> errors)
        {
            if (captionParagraph == null) return false;

            string pattern = @"^Таблица\s\d+\s*[-–]\s*\D.*";
            string text = captionParagraph.InnerText.Trim();

            if (!Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
            {
                errors.Add($"Неверный формат подписи таблицы: '{GetShortText(text)}'. Требуется формат: 'Таблица N - Название'");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Проверяет стиль подписи таблицы
        /// </summary>
        /// <param name="captionParagraph"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private bool CheckTableCaptionStyle(Paragraph captionParagraph, List<string> errors)
        {
            bool isValid = true;
            var errorDetails = new List<string>();

            // Проверка шрифта
            if (!string.IsNullOrEmpty(_gost.TableCaptionFontName))
            {
                foreach (var run in captionParagraph.Elements<Run>())
                {
                    if (_shouldSkipRun(run)) continue;

                    var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? DefaultTableCaptionFont;
                    if (font != null && font != _gost.TableCaptionFontName)
                    {
                        errorDetails.Add($"шрифт подписи таблицы должен быть: {_gost.TableCaptionFontName}, а не {font}");
                        isValid = false;
                    }
                }
            }

            // Проверка размера шрифта
            if (_gost.TableCaptionFontSize.HasValue)
            {
                foreach (var run in captionParagraph.Elements<Run>())
                {
                    if (_shouldSkipRun(run)) continue;

                    var fontSizeVal = run.RunProperties?.FontSize?.Val?.Value;
                    double actualFontSize = fontSizeVal != null ? double.Parse(fontSizeVal) / 2 : DefaultTableCaptionFontSize;

                    if (Math.Abs(actualFontSize - _gost.TableCaptionFontSize.Value) > 0.1)
                    {
                        errorDetails.Add($"размер шрифта подписи должен быть {_gost.TableCaptionFontSize.Value}, а не {actualFontSize}");
                        isValid = false;
                    }
                }
            }

            // Проверка выравнивания
            if (!string.IsNullOrEmpty(_gost.TableCaptionAlignment))
            {
                string currentAlignment = GetAlignmentString(captionParagraph.ParagraphProperties?.Justification) ?? DefaultTableCaptionAlignment;
                if (currentAlignment != _gost.TableCaptionAlignment)
                {
                    errorDetails.Add($"выравнивание подписи должно быть: {_gost.TableCaptionAlignment}, а не {currentAlignment}");
                    isValid = false;
                }
            }

            // Проверка отступов
            var indent = captionParagraph.ParagraphProperties?.Indentation;

            // Проверка левого отступа
            if (_gost.TableCaptionIndentLeft.HasValue)
            {
                double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultTableCaptionIndentLeft;
                if (Math.Abs(actualLeft - _gost.TableCaptionIndentLeft.Value) > 0.05)
                {
                    errorDetails.Add($"левый отступ подписи: {actualLeft:F2} см (требуется {_gost.TableCaptionIndentLeft.Value:F2} см)");
                    isValid = false;
                }
            }

            // Проверка правого отступа
            if (_gost.TableCaptionIndentRight.HasValue)
            {
                double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTableCaptionIndentRight;
                if (Math.Abs(actualRight - _gost.TableCaptionIndentRight.Value) > 0.05)
                {
                    errorDetails.Add($"правый отступ подписи: {actualRight:F2} см (требуется {_gost.TableCaptionIndentRight.Value:F2} см)");
                    isValid = false;
                }
            }

            // Проверка отступа первой строки
            if (_gost.TableCaptionFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TableCaptionIndentOrOutdent))
            {
                string currentType = DefaultTableCaptionIndentOrOutdent;
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

                if (!string.IsNullOrEmpty(_gost.TableCaptionIndentOrOutdent) && _gost.TableCaptionIndentOrOutdent != "Нет")
                {
                    if (currentType != _gost.TableCaptionIndentOrOutdent)
                    {
                        errorDetails.Add($"тип первой строки: {currentType} (требуется {_gost.TableCaptionIndentOrOutdent})");
                        isValid = false;
                    }

                    if (currentValue.HasValue && Math.Abs(currentValue.Value - _gost.TableCaptionFirstLineIndent.Value) > 0.05)
                    {
                        errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {_gost.TableCaptionFirstLineIndent.Value:F2} см)");
                        isValid = false;
                    }
                }
            }

            // Проверка межстрочных интервалов
            var spacing = captionParagraph.ParagraphProperties?.SpacingBetweenLines;

            // Проверка типа межстрочного интервала 
            if (!string.IsNullOrEmpty(_gost.TableCaptionLineSpacingType))
            {
                string currentSpacingType = ConvertSpacingRuleToName(spacing?.LineRule);

                if (currentSpacingType != _gost.TableCaptionLineSpacingType)
                {
                    errorDetails.Add($"тип межстрочного интервала: {currentSpacingType} (требуется {_gost.TableCaptionLineSpacingType})");
                    isValid = false;
                }
            }

            // Проверка межстрочного интервала
            if (_gost.TableCaptionLineSpacingValue.HasValue)
            {
                double actualSpacing = spacing?.Line != null ? CalculateActualSpacing(spacing) : DefaultTableCaptionLineSpacingValue;
                if (Math.Abs(actualSpacing - _gost.TableCaptionLineSpacingValue.Value) > 0.1)
                {
                    errorDetails.Add($"межстрочный интервал подписи должен быть {_gost.TableCaptionLineSpacingValue.Value}, а не {actualSpacing}");
                    isValid = false;
                }
            }

            // Проверка интервала перед
            if (_gost.TableCaptionLineSpacingBefore.HasValue)
            {
                double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultTableCaptionLineSpacingBefore;
                if (Math.Abs(actualBefore - _gost.TableCaptionLineSpacingBefore.Value) > 0.1)
                {
                    errorDetails.Add($"интервал перед подписью должен быть {_gost.TableCaptionLineSpacingBefore.Value}, а не {actualBefore}");
                    isValid = false;
                }
            }

            // Проверка интервала после
            if (_gost.TableCaptionLineSpacingAfter.HasValue)
            {
                double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultTableCaptionLineSpacingAfter;
                if (Math.Abs(actualAfter - _gost.TableCaptionLineSpacingAfter.Value) > 0.1)
                {
                    errorDetails.Add($"интервал после подписи должен быть {_gost.TableCaptionLineSpacingAfter.Value}, а не {actualAfter}");
                    isValid = false;
                }
            }

            if (!isValid)
            {
                errors.Add($"Подпись таблицы '{GetShortText(captionParagraph.InnerText)}':");
                errors.AddRange(errorDetails.Select(d => $"  • {d}"));
            }

            return isValid;
        }

        /// <summary>
        /// Проверяет содержимое таблицы (шрифт, выравнивание, отступы)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private bool CheckTableContent(Table table, List<string> errors)
        {
            bool isValid = true;
            var errorDetails = new List<string>();

            foreach (var row in table.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        // Проверка шрифта
                        if (!string.IsNullOrEmpty(_gost.FontName))
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                if (_shouldSkipRun(run)) continue;

                                var font = run.RunProperties?.RunFonts?.Ascii?.Value ?? DefaultTableCaptionFont;
                                if (font != null && font != _gost.FontName)
                                {
                                    errorDetails.Add($"шрифт в таблице должен быть: {_gost.FontName}, а не {font}");
                                    isValid = false;
                                }
                            }
                        }

                        // Проверка размера шрифта
                        if (_gost.TableFontSize.HasValue)
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                if (_shouldSkipRun(run)) continue;

                                var fontSize = run.RunProperties?.FontSize?.Val;
                                double actualFontSize = fontSize == null ? DefaultTableFontSize : double.Parse(fontSize) / 2;

                                if (Math.Abs(actualFontSize - _gost.TableFontSize.Value) > 0.1)
                                {
                                    errorDetails.Add($"размер шрифта в таблице должен быть {_gost.TableFontSize.Value:F2} pt, а не {actualFontSize:F2} pt");
                                    isValid = false;
                                }
                            }
                        }

                        // Проверка выравнивания
                        if (!string.IsNullOrEmpty(_gost.TableAlignment))
                        {
                            string currentAlignment = GetAlignmentString(paragraph.ParagraphProperties?.Justification) ?? DefaultTableAlignment;
                            if (currentAlignment != _gost.TableAlignment)
                            {
                                errorDetails.Add($"выравнивание в таблице должно быть: {_gost.TableAlignment}, а не {currentAlignment}");
                                isValid = false;
                            }
                        }

                        // Проверка отступов
                        var indent = paragraph.ParagraphProperties?.Indentation;

                        // Проверка левого отступа
                        if (_gost.TableIndentLeft.HasValue)
                        {
                            double actualLeft = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : DefaultTableIndentLeft;
                            if (Math.Abs(actualLeft - _gost.TableIndentLeft.Value) > 0.05)
                            {
                                errorDetails.Add($"левый отступ в таблице: {actualLeft:F2} см (требуется {_gost.TableIndentLeft.Value:F2} см)");
                                isValid = false;
                            }
                        }

                        // Проверка правого отступа
                        if (_gost.TableIndentRight.HasValue)
                        {
                            double actualRight = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : DefaultTableIndentRight;
                            if (Math.Abs(actualRight - _gost.TableIndentRight.Value) > 0.05)
                            {
                                errorDetails.Add($"правый отступ в таблице: {actualRight:F2} см (требуется {_gost.TableIndentRight.Value:F2} см)");
                                isValid = false;
                            }
                        }

                        // Проверка отступа первой строки
                        if (_gost.TableFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TableIndentOrOutdent))
                        {
                            string currentType = DefaultTableIndentOrOutdent;
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

                            if (!string.IsNullOrEmpty(_gost.TableIndentOrOutdent) && _gost.TableIndentOrOutdent != "Нет")
                            {
                                if (currentType != _gost.TableIndentOrOutdent)
                                {
                                    errorDetails.Add($"тип первой строки: {currentType} (требуется {_gost.TableIndentOrOutdent})");
                                    isValid = false;
                                }

                                if (currentValue.HasValue && Math.Abs(currentValue.Value - _gost.TableFirstLineIndent.Value) > 0.05)
                                {
                                    errorDetails.Add($"{currentType} первой строки: {currentValue.Value:F2} см (требуется {_gost.TableFirstLineIndent.Value:F2} см)");
                                    isValid = false;
                                }
                            }
                        }

                        // Проверка межстрочных интервалов
                        var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;

                        // Проверка типа межстрочного интервала
                        if (!string.IsNullOrEmpty(_gost.TableSpacingType))
                        {
                            string currentSpacingType = ConvertSpacingRuleToName(spacing?.LineRule);

                            if (currentSpacingType != _gost.TableSpacingType)
                            {
                                errorDetails.Add($"тип межстрочного интервала в таблице: {currentSpacingType} (требуется {_gost.TableSpacingType})");
                                isValid = false;
                            }
                        }

                        // Проверка межстрочного интервала
                        if (_gost.TableLineSpacingValue.HasValue)
                        {
                            double actualSpacing = spacing?.Line != null ? CalculateActualSpacing(spacing) : DefaultTableLineSpacingValue;
                            if (Math.Abs(actualSpacing - _gost.TableLineSpacingValue.Value) > 0.1)
                            {
                                errorDetails.Add($"межстрочный интервал в таблице должен быть {_gost.TableLineSpacingValue.Value}, а не {actualSpacing}");
                                isValid = false;
                            }
                        }

                        // Проверка интервала перед
                        if (_gost.TableLineSpacingBefore.HasValue)
                        {
                            double actualBefore = spacing?.Before?.Value != null ? ConvertTwipsToPoints(spacing.Before.Value) : DefaultTableLineSpacingBefore;
                            if (Math.Abs(actualBefore - _gost.TableLineSpacingBefore.Value) > 0.1)
                            {
                                errorDetails.Add($"интервал перед в таблице должен быть {_gost.TableLineSpacingBefore.Value}, а не {actualBefore}");
                                isValid = false;
                            }
                        }

                        // Проверка интервала после
                        if (_gost.TableLineSpacingAfter.HasValue)
                        {
                            double actualAfter = spacing?.After?.Value != null ? ConvertTwipsToPoints(spacing.After.Value) : DefaultTableLineSpacingAfter;
                            if (Math.Abs(actualAfter - _gost.TableLineSpacingAfter.Value) > 0.1)
                            {
                                errorDetails.Add($"интервал после в таблице должен быть {_gost.TableLineSpacingAfter.Value}, а не {actualAfter}");
                                isValid = false;
                            }
                        }
                    }
                }
            }

            if (!isValid)
            {
                errors.Add("Содержимое таблицы:");
                errors.AddRange(errorDetails.Distinct().Select(d => $"  • {d}"));
            }

            return isValid;
        }

        /// <summary>
        /// Получает подпись таблицы (должна быть непосредственно перед таблицей)
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private Paragraph GetTableCaption(Table table)
        {
            var previousElement = table.PreviousSibling();
            while (previousElement != null)
            {
                if (previousElement is Paragraph paragraph && !string.IsNullOrWhiteSpace(paragraph.InnerText))
                {
                    return paragraph;
                }
                previousElement = previousElement.PreviousSibling();
            }
            return null;
        }

        /// <summary>
        /// Обрезает текст параграфа до 50 символов с добавлением многоточия
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string GetShortText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "[пустой элемент]";
            return text.Length > 30 ? text.Substring(0, 27) + "..." : text;
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
            return value / 567.0; // Стандартное преобразование twips -> pt
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
                // Точно
                return double.Parse(spacing.Line.Value) / 20.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.AtLeast)
            {
                // Минимум
                return double.Parse(spacing.Line.Value) / 20.0;
            }
            else if (spacing.LineRule == LineSpacingRuleValues.Auto)
            {
                // Множитель
                return double.Parse(spacing.Line.Value) / 240.0;
            }
            else
            {
                // По умолчанию множитель
                return double.Parse(spacing.Line.Value) / 240.0;
            }
        }
    }
}