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
using DocumentFormat.OpenXml.Spreadsheet;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using GOST_Control;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок подписей к таблицами и текста в таблице
    /// </summary>
    public class CheckingTableDoc
    {
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
        public async Task<(bool IsValid, List<TextErrorInfo> Errors)> CheckTablesAsync(Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var body = _wordDoc.MainDocumentPart.Document.Body;
                var tables = body.Elements<Table>().ToList();
                var errors = new List<TextErrorInfo>();
                bool allTablesValid = true;

                // Получаем все стили документа
                var allStyles = _wordDoc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()?.ToDictionary(s => s.StyleId.Value) ?? new Dictionary<string, Style>();

                foreach (var table in tables)
                {
                    bool tableValid = true;
                    var tableErrors = new List<TextErrorInfo>();

                    // 1. Проверка подписи таблицы (должна быть над таблицей)
                    var tableCaption = GetTableCaption(table);

                    if (tableCaption == null)
                    {
                        continue;
                    }

                    // Проверка формата подписи
                    if (!CheckTableCaptionFormat(tableCaption, tableErrors))
                    {
                        tableValid = false;
                    }

                    // Проверка стиля подписи таблицы
                    if (!CheckTableCaptionStyle(tableCaption, tableErrors, allStyles))
                    {
                        tableValid = false;
                    }

                    // Проверка содержимого таблицы 
                    if (!CheckTableContent(table, tableErrors, allStyles))
                    {
                        tableValid = false;
                    }

                    if (!tableValid)
                    {
                        allTablesValid = false;
                        tableCaption = GetTableCaption(table);
                        string captionText = tableCaption != null ? GetShortText(tableCaption.InnerText) : "Таблица без подписи";

                        // Добавляем заголовок с названием таблицы
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"Таблица \"{captionText}\":",
                            ProblemParagraph = tableCaption,
                            ProblemRun = null
                        });

                        // Добавляем все ошибки для этой таблицы
                        errors.AddRange(tableErrors);

                        // Добавляем пустую строку для разделения
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = "",
                            ProblemParagraph = null,
                            ProblemRun = null
                        });
                    }
                }

                Dispatcher.UIThread.Post(() =>
                {
                    if (!allTablesValid)
                    {
                        var msg = $"Ошибки в таблицах:\n{string.Join("\n", errors.Select(e => e.ErrorMessage).Take(10))}";
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
        private bool CheckTableCaptionFormat(Paragraph captionParagraph, List<TextErrorInfo> errors)
        {
            if (captionParagraph == null) return false;

            string pattern = @"^Таблица\s\d+\s*[-–]\s*\D.*";
            string text = captionParagraph.InnerText.Trim();

            if (!Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
            {
                errors.Add(new TextErrorInfo
                {
                    ErrorMessage = $"Неверный формат подписи таблицы: '{GetShortText(text)}'. Требуется формат: 'Таблица N - Название'",
                    ProblemParagraph = captionParagraph,
                    ProblemRun = null
                });
                return false;
            }

            return true;
        }

        /// <summary>
        /// Проверяет стиль подписи таблицы с учетом наследования стилей
        /// </summary>
        /// <param name="captionParagraph">Параграф подписи таблицы</param>
        /// <param name="errors">Список для записи ошибок</param>
        /// <returns>True если стиль соответствует требованиям, иначе False</returns>
        private bool CheckTableCaptionStyle(Paragraph captionParagraph, List<TextErrorInfo> errors, Dictionary<string, Style> allStyles)
        {
            bool isValid = true;

            bool hasFontError = false;
            bool hasFontSizeError = false;

            foreach (var run in captionParagraph.Elements<Run>())
            {
                if (_shouldSkipRun(run)) continue;

                // 1. Проверка шрифта подписи таблицы
                if (!string.IsNullOrEmpty(_gost.TableCaptionFontName) && !hasFontError)
                {
                    var (actualFont, isFontDefined) = GetActualFontForTable(run, captionParagraph, allStyles);

                    if (!isFontDefined)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = "       • Не удалось определить шрифт подписи таблицы",
                            ProblemRun = run, 
                            ProblemParagraph = captionParagraph
                        });
                        isValid = false;
                        hasFontError = true; 
                    }
                    else if (!string.Equals(actualFont, _gost.TableCaptionFontName, StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Шрифт подписи таблицы должен быть: {_gost.TableCaptionFontName}, а не {actualFont}",
                            ProblemRun = run, 
                            ProblemParagraph = captionParagraph
                        });
                        isValid = false;
                        hasFontError = true;
                    }
                }

                // 2. Проверка размера шрифта подписи таблицы
                if (_gost.TableCaptionFontSize.HasValue && !hasFontSizeError)
                {
                    var (actualSize, isSizeDefined) = GetActualFontSizeForTable(run, captionParagraph, allStyles);

                    if (!isSizeDefined)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = "       • Не удалось определить размер шрифта подписи таблицы",
                            ProblemRun = run, 
                            ProblemParagraph = captionParagraph
                        });
                        isValid = false;
                        hasFontSizeError = true;
                    }
                    else if (Math.Abs(actualSize - _gost.TableCaptionFontSize.Value) > 0.1)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Размер шрифта подписи таблицы должен быть {_gost.TableCaptionFontSize.Value:F1} pt, а не {actualSize:F1} pt",
                            ProblemRun = run, 
                            ProblemParagraph = captionParagraph
                        });
                        isValid = false;
                        hasFontSizeError = true; 
                    }
                }

                if (hasFontError && hasFontSizeError)
                    break;
            }

            // 3. Проверка выравнивания подписи таблицы
            if (!string.IsNullOrEmpty(_gost.TableCaptionAlignment))
            {
                var (actualAlignment, isAlignmentDefined) = GetActualAlignmentForTable(captionParagraph, allStyles);

                if (!isAlignmentDefined)
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = "       • Не удалось определить выравнивание подписи таблицы",
                        ProblemParagraph = captionParagraph,
                        ProblemRun = null
                    });
                    isValid = false;
                }
                else if (actualAlignment != _gost.TableCaptionAlignment)
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = $"       • Выравнивание подписи таблицы должно быть: {_gost.TableCaptionAlignment}, а не {actualAlignment}",
                        ProblemParagraph = captionParagraph,
                        ProblemRun = null
                    });
                    isValid = false;
                }
            }

            // 4. Проверка отступов подписи таблицы
            var indent = captionParagraph.ParagraphProperties?.Indentation;
            var styleIndent = GetStyleIndentationForTable(captionParagraph, allStyles);
            indent ??= styleIndent;

            // Преобразуем все значения в сантиметры
            double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
            double rightIndent = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : 0;
            double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
            double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

            // 4.1. Проверка левого отступа подписи
            if (_gost.TableCaptionIndentLeft.HasValue)
            {
                double actualTextIndent = leftIndent;

                if (hangingIndent > 0)
                    actualTextIndent = leftIndent - hangingIndent;

                if (Math.Abs(actualTextIndent - _gost.TableCaptionIndentLeft.Value) > 0.05)
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = $"       • Левый отступ подписи таблицы: {actualTextIndent:F2} см (требуется {_gost.TableCaptionIndentLeft.Value:F2} см)",
                        ProblemParagraph = captionParagraph,
                        ProblemRun = null
                    });
                    isValid = false;
                }
            }

            // 4.2. Проверка правого отступа подписи
            if (_gost.TableCaptionIndentRight.HasValue)
            {
                if (Math.Abs(rightIndent - _gost.TableCaptionIndentRight.Value) > 0.05)
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = $"       • Правый отступ подписи таблицы: {rightIndent:F2} см (требуется {_gost.TableCaptionIndentRight.Value:F2} см)",
                        ProblemParagraph = captionParagraph,
                        ProblemRun = null
                    });
                    isValid = false;
                }
            }

            // 4.3. Проверка первой строки подписи
            if (_gost.TableCaptionFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TableCaptionIndentOrOutdent))
            {
                bool isHanging = hangingIndent > 0;
                bool isFirstLine = firstLineIndent > 0;

                // Проверка типа (выступ/отступ)
                if (!string.IsNullOrEmpty(_gost.TableCaptionIndentOrOutdent))
                {
                    string actualType = isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет";
                    if (actualType != _gost.TableCaptionIndentOrOutdent)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Тип первой строки подписи: '{actualType}' (требуется '{_gost.TableCaptionIndentOrOutdent}')",
                            ProblemParagraph = captionParagraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }

                // Проверка значения отступа/выступа
                if (_gost.TableCaptionFirstLineIndent.HasValue)
                {
                    double actualValue = isHanging ? hangingIndent : firstLineIndent;
                    if ((isHanging || isFirstLine) && Math.Abs(actualValue - _gost.TableCaptionFirstLineIndent.Value) > 0.05)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • {(isHanging ? "Выступ" : "Отступ")} первой строки подписи: {actualValue:F2} см (требуется {_gost.TableCaptionFirstLineIndent.Value:F2} см)",
                            ProblemParagraph = captionParagraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                    else if (!isHanging && !isFirstLine && _gost.TableCaptionIndentOrOutdent != "Нет")
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Отсутствует {_gost.TableCaptionIndentOrOutdent} первой строки подписи",
                            ProblemParagraph = captionParagraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }
            }

            // 5. Проверка межстрочных интервалов подписи
            if (_gost.TableCaptionLineSpacingValue.HasValue || !string.IsNullOrEmpty(_gost.TableCaptionLineSpacingType))
            {
                var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacingForTable(captionParagraph, allStyles);

                if (!isSpacingDefined)
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = "       • Не удалось определить межстрочный интервал подписи таблицы",
                        ProblemParagraph = captionParagraph,
                        ProblemRun = null
                    });
                    isValid = false;
                }
                else
                {
                    // Проверка типа интервала
                    if (!string.IsNullOrEmpty(_gost.TableCaptionLineSpacingType))
                    {
                        if (actualSpacingType != _gost.TableCaptionLineSpacingType)
                        {
                            errors.Add(new TextErrorInfo
                            {
                                ErrorMessage = $"       • Тип межстрочного интервала подписи: '{actualSpacingType}' (требуется '{_gost.TableCaptionLineSpacingType}')",
                                ProblemParagraph = captionParagraph,
                                ProblemRun = null
                            });
                            isValid = false;
                        }
                    }

                    // Проверка значения интервала
                    if (_gost.TableCaptionLineSpacingValue.HasValue && Math.Abs(actualSpacingValue - _gost.TableCaptionLineSpacingValue.Value) > 0.1)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Межстрочный интервал подписи: {actualSpacingValue:F2} (требуется {_gost.TableCaptionLineSpacingValue.Value:F2})",
                            ProblemParagraph = captionParagraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }
            }

            // 6. Проверка интервалов перед/после подписи
            if (_gost.TableCaptionLineSpacingBefore.HasValue || _gost.TableCaptionLineSpacingAfter.HasValue)
            {
                var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacingForCaption(captionParagraph, allStyles);

                if (!isSpacingDefined)
                {
                    errors.Add(new TextErrorInfo
                    {
                        ErrorMessage = "       • Не удалось определить интервалы перед/после подписи таблицы",
                        ProblemParagraph = captionParagraph,
                        ProblemRun = null
                    });
                    isValid = false;
                }
                else
                {
                    // Проверка интервала перед абзацем
                    if (_gost.TableCaptionLineSpacingBefore.HasValue && Math.Abs(actualBefore - _gost.TableCaptionLineSpacingBefore.Value) > 0.1)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Интервал перед подписью: {actualBefore:F1} pt (требуется {_gost.TableCaptionLineSpacingBefore.Value:F1} pt)",
                            ProblemParagraph = captionParagraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }

                    // Проверка интервала после абзаца
                    if (_gost.TableCaptionLineSpacingAfter.HasValue && Math.Abs(actualAfter - _gost.TableCaptionLineSpacingAfter.Value) > 0.1)
                    {
                        errors.Add(new TextErrorInfo
                        {
                            ErrorMessage = $"       • Интервал после подписи: {actualAfter:F1} pt (требуется {_gost.TableCaptionLineSpacingAfter.Value:F1} pt)",
                            ProblemParagraph = captionParagraph,
                            ProblemRun = null
                        });
                        isValid = false;
                    }
                }
            }

            return isValid;
        }

        /// <summary>
        /// Проверяет содержимое таблицы (шрифт, выравнивание, отступы)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private bool CheckTableContent(Table table, List<TextErrorInfo> errors, Dictionary<string, Style> allStyles)
        {
            bool isValid = true;

            foreach (var row in table.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        // 1. Проверка шрифта текста в таблице
                        if (!string.IsNullOrEmpty(_gost.FontName))
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                if (_shouldSkipRun(run)) continue;

                                var (actualFont, isFontDefined) = GetActualFontForTable(run, paragraph, allStyles);

                                if (!isFontDefined)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Не удалось определить шрифт в таблице",
                                        ProblemRun = run,
                                        ProblemParagraph = paragraph
                                    });
                                    isValid = false;
                                }
                                else if (!string.Equals(actualFont, _gost.FontName, StringComparison.OrdinalIgnoreCase))
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Шрифт в таблице должен быть: {_gost.FontName}, а не {actualFont}",
                                        ProblemRun = run,
                                        ProblemParagraph = paragraph
                                    });
                                    isValid = false;
                                }
                            }
                        }

                        // 2. Проверка размера шрифта текста в таблице
                        if (_gost.TableFontSize.HasValue)
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                if (_shouldSkipRun(run)) continue;

                                var (actualSize, isSizeDefined) = GetActualFontSizeForTable(run, paragraph, allStyles);

                                if (!isSizeDefined)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = "       • Не удалось определить размер шрифта в таблице",
                                        ProblemRun = run,
                                        ProblemParagraph = paragraph
                                    });
                                    isValid = false;
                                }
                                else if (Math.Abs(actualSize - _gost.TableFontSize.Value) > 0.1)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Размер шрифта в таблице должен быть {_gost.TableFontSize.Value:F1} pt, а не {actualSize:F1} pt",
                                        ProblemRun = run,
                                        ProblemParagraph = paragraph
                                    });
                                    isValid = false;
                                }
                            }
                        }

                        // 3. Проверка выравнивания текста в таблице
                        if (!string.IsNullOrEmpty(_gost.TableAlignment))
                        {
                            var (actualAlignment, isAlignmentDefined) = GetActualAlignmentForTable(paragraph, allStyles);

                            if (!isAlignmentDefined)
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = "       • Не удалось определить выравнивание текста в таблице",
                                    ProblemParagraph = paragraph,
                                    ProblemRun = null
                                });
                                isValid = false;
                            }
                            else if (actualAlignment != _gost.TableAlignment)
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"       • Выравнивание в таблице должно быть: {_gost.TableAlignment}, а не {actualAlignment}",
                                    ProblemParagraph = paragraph,
                                    ProblemRun = null
                                });
                                isValid = false;
                            }
                        }

                        // 4. Проверка отступов в таблице
                        var indent = paragraph.ParagraphProperties?.Indentation;
                        var styleIndent = GetStyleIndentationForTable(paragraph, allStyles);
                        indent ??= styleIndent;

                        // Преобразуем все значения в сантиметры
                        double leftIndent = indent?.Left?.Value != null ? TwipsToCm(double.Parse(indent.Left.Value)) : 0;
                        double rightIndent = indent?.Right?.Value != null ? TwipsToCm(double.Parse(indent.Right.Value)) : 0;
                        double firstLineIndent = indent?.FirstLine?.Value != null ? TwipsToCm(double.Parse(indent.FirstLine.Value)) : 0;
                        double hangingIndent = indent?.Hanging?.Value != null ? TwipsToCm(double.Parse(indent.Hanging.Value)) : 0;

                        // 4.1. Проверка левого отступа в таблице
                        if (_gost.TableIndentLeft.HasValue)
                        {
                            double actualTextIndent = leftIndent;

                            if (hangingIndent > 0)
                                actualTextIndent = leftIndent - hangingIndent;

                            if (Math.Abs(actualTextIndent - _gost.TableIndentLeft.Value) > 0.05)
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"       • Левый отступ в таблице: {actualTextIndent:F2} см (требуется {_gost.TableIndentLeft.Value:F2} см)",
                                    ProblemParagraph = paragraph,
                                    ProblemRun = null
                                });
                                isValid = false;
                            }
                        }

                        // 4.2. Проверка правого отступа в таблице
                        if (_gost.TableIndentRight.HasValue)
                        {
                            if (Math.Abs(rightIndent - _gost.TableIndentRight.Value) > 0.05)
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = $"       • Правый отступ в таблице: {rightIndent:F2} см (требуется {_gost.TableIndentRight.Value:F2} см)",
                                    ProblemParagraph = paragraph,
                                    ProblemRun = null
                                });
                                isValid = false;
                            }
                        }

                        // 4.3. Проверка первой строки в таблице
                        if (_gost.TableFirstLineIndent.HasValue || !string.IsNullOrEmpty(_gost.TableIndentOrOutdent))
                        {
                            bool isHanging = hangingIndent > 0;
                            bool isFirstLine = firstLineIndent > 0;

                            // Проверка типа (выступ/отступ)
                            if (!string.IsNullOrEmpty(_gost.TableIndentOrOutdent))
                            {
                                string actualType = isHanging ? "Выступ" : isFirstLine ? "Отступ" : "Нет";
                                if (actualType != _gost.TableIndentOrOutdent)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Тип первой строки в таблице: '{actualType}' (требуется '{_gost.TableIndentOrOutdent}')",
                                        ProblemParagraph = paragraph,
                                        ProblemRun = null
                                    });
                                    isValid = false;
                                }
                            }

                            // Проверка значения отступа/выступа
                            if (_gost.TableFirstLineIndent.HasValue)
                            {
                                double actualValue = isHanging ? hangingIndent : firstLineIndent;
                                if ((isHanging || isFirstLine) && Math.Abs(actualValue - _gost.TableFirstLineIndent.Value) > 0.05)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • {(isHanging ? "Выступ" : "Отступ")} первой строки в таблице: {actualValue:F2} см (требуется {_gost.TableFirstLineIndent.Value:F2} см)",
                                        ProblemParagraph = paragraph,
                                        ProblemRun = null
                                    });
                                    isValid = false;
                                }
                                else if (!isHanging && !isFirstLine && _gost.TableIndentOrOutdent != "Нет")
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Отсутствует {_gost.TableIndentOrOutdent} первой строки в таблице",
                                        ProblemParagraph = paragraph,
                                        ProblemRun = null
                                    });
                                    isValid = false;
                                }
                            }
                        }

                        // 5. Проверка межстрочных интервалов в таблице
                        if (_gost.TableLineSpacingValue.HasValue || !string.IsNullOrEmpty(_gost.TableSpacingType))
                        {
                            var (actualSpacingType, actualSpacingValue, isSpacingDefined) = GetActualLineSpacingForTable(paragraph, allStyles);

                            if (!isSpacingDefined)
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = "       • Не удалось определить межстрочный интервал в таблице",
                                    ProblemParagraph = paragraph,
                                    ProblemRun = null
                                });
                                isValid = false;
                            }
                            else
                            {
                                // Проверка типа интервала
                                if (!string.IsNullOrEmpty(_gost.TableSpacingType))
                                {
                                    if (actualSpacingType != _gost.TableSpacingType)
                                    {
                                        errors.Add(new TextErrorInfo
                                        {
                                            ErrorMessage = $"       • Тип межстрочного интервала в таблице: '{actualSpacingType}' (требуется '{_gost.TableSpacingType}')",
                                            ProblemParagraph = paragraph,
                                            ProblemRun = null
                                        });
                                        isValid = false;
                                    }
                                }

                                // Проверка значения интервала
                                if (_gost.TableLineSpacingValue.HasValue && Math.Abs(actualSpacingValue - _gost.TableLineSpacingValue.Value) > 0.1)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Межстрочный интервал в таблице: {actualSpacingValue:F2} (требуется {_gost.TableLineSpacingValue.Value:F2})",
                                        ProblemParagraph = paragraph,
                                        ProblemRun = null
                                    });
                                    isValid = false;
                                }
                            }
                        }

                        // 6. Проверка интервалов перед/после абзаца в таблице
                        if (_gost.TableLineSpacingBefore.HasValue || _gost.TableLineSpacingAfter.HasValue)
                        {
                            var (actualBefore, actualAfter, isSpacingDefined) = GetActualParagraphSpacingForTable(paragraph, allStyles);

                            if (!isSpacingDefined)
                            {
                                errors.Add(new TextErrorInfo
                                {
                                    ErrorMessage = "       • Не удалось определить интервалы перед/после абзаца в таблице",
                                    ProblemParagraph = paragraph,
                                    ProblemRun = null
                                });
                                isValid = false;
                            }
                            else
                            {
                                // Проверка интервала перед абзацем
                                if (_gost.TableLineSpacingBefore.HasValue && Math.Abs(actualBefore - _gost.TableLineSpacingBefore.Value) > 0.1)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Интервал перед абзацем в таблице: {actualBefore:F1} pt (требуется {_gost.TableLineSpacingBefore.Value:F1} pt)",
                                        ProblemParagraph = paragraph,
                                        ProblemRun = null
                                    });
                                    isValid = false;
                                }

                                // Проверка интервала после абзаца
                                if (_gost.TableLineSpacingAfter.HasValue && Math.Abs(actualAfter - _gost.TableLineSpacingAfter.Value) > 0.1)
                                {
                                    errors.Add(new TextErrorInfo
                                    {
                                        ErrorMessage = $"       • Интервал после абзаца в таблице: {actualAfter:F1} pt (требуется {_gost.TableLineSpacingAfter.Value:F1} pt)",
                                        ProblemParagraph = paragraph,
                                        ProblemRun = null
                                    });
                                    isValid = false;
                                }
                            }
                        }
                    }
                }
            }

            return isValid;
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

        /// <summary>
        /// Определение стиля интервалов
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private (double Before, double After, bool IsDefined) GetActualParagraphSpacingForTable(Paragraph paragraph, Dictionary<string, Style> allStyles)
        {
            double? before = null;
            double? after = null;

            // 1. Проверяем явные свойства абзаца
            var spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            if (spacing != null)
            {
                if (spacing.Before?.Value != null)
                    before = ConvertTwipsToPoints(spacing.Before.Value);
                if (spacing.After?.Value != null)
                    after = ConvertTwipsToPoints(spacing.After.Value);
            }

            // 2. Проверяем стиль абзаца и его родителей
            if (before == null || after == null)
            {
                var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
                if (paraStyleId != null && allStyles.TryGetValue(paraStyleId.Value, out var currentStyle))
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

                        currentStyle = currentStyle.BasedOn?.Val != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
                    }
                }
            }

            // 3. Проверяем свойства таблицы и ячейки
            if (before == null || after == null)
            {
                var tableCell = paragraph.Ancestors<TableCell>().FirstOrDefault();
                if (tableCell != null)
                {
                    var tableCellProps = tableCell.TableCellProperties;
                    var cellSpacing = tableCellProps?.TableCellMargin;

                    if (cellSpacing != null)
                    {
                        if (before == null && cellSpacing.TopMargin?.Width?.Value != null)
                            before = ConvertTwipsToPoints(cellSpacing.TopMargin.Width.Value);
                        if (after == null && cellSpacing.BottomMargin?.Width?.Value != null)
                            after = ConvertTwipsToPoints(cellSpacing.BottomMargin.Width.Value);
                    }

                    // Проверяем стиль таблицы
                    var table = tableCell.Ancestors<Table>().FirstOrDefault();
                    if (table != null)
                    {
                        var tableProps = table.Elements<TableProperties>().FirstOrDefault();
                        var tableStyle = tableProps?.TableStyle;

                        if (tableStyle != null && tableStyle.Val != null && allStyles.TryGetValue(tableStyle.Val.Value, out var tblStyle))
                        {
                            var tblStyleProps = tblStyle.StyleTableProperties;
                            var tblCellProps = tblStyle.StyleTableCellProperties;

                            if (tblCellProps?.TableCellMargin != null)
                            {
                                if (before == null && tblCellProps.TableCellMargin.TopMargin?.Width?.Value != null)
                                    before = ConvertTwipsToPoints(tblCellProps.TableCellMargin.TopMargin.Width.Value);
                                if (after == null && tblCellProps.TableCellMargin.BottomMargin?.Width?.Value != null)
                                    after = ConvertTwipsToPoints(tblCellProps.TableCellMargin.BottomMargin.Width.Value);
                            }
                        }
                    }
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

        /// <summary>
        /// Поиск стиля для интервалов "межстрочного"
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private (string Type, double Value, bool IsDefined) GetActualLineSpacingForTable(Paragraph paragraph, Dictionary<string, Style> allStyles)
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

                    currentStyle = currentStyle.BasedOn?.Val?.Value != null && allStyles.TryGetValue(currentStyle.BasedOn.Val.Value, out var basedOnStyle) ? basedOnStyle : null;
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
        /// Опредеелние стиля для отступов
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private Indentation GetStyleIndentationForTable(Paragraph paragraph, Dictionary<string, Style> allStyles)
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

        /// <summary>
        /// Поиск стиля для выравнивания
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private (string Alignment, bool IsDefined) GetActualAlignmentForTable(Paragraph paragraph, Dictionary<string, Style> allStyles)
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

            return ("Left", false);
        }

        /// <summary>
        /// Поиск стиля для размера шрифта
        /// </summary>
        /// <param name="run"></param>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private (double Size, bool IsDefined) GetActualFontSizeForTable(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles)
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

        /// <summary>
        /// Поиск стиля для шрифта
        /// </summary>
        /// <param name="run"></param>
        /// <param name="paragraph"></param>
        /// <param name="allStyles"></param>
        /// <returns></returns>
        private (string FontName, bool IsDefined) GetActualFontForTable(Run run, Paragraph paragraph, Dictionary<string, Style> allStyles)
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

            return runProps.RunFonts?.Ascii?.Value ?? runProps.RunFonts?.HighAnsi?.Value ?? runProps.RunFonts?.ComplexScript?.Value ?? runProps.RunFonts?.EastAsia?.Value;
        }

        private string GetStyleFont(Style style)
        {
            if (style?.StyleRunProperties == null) return null;
            return style.StyleRunProperties.RunFonts?.Ascii?.Value ?? style.StyleRunProperties.RunFonts?.HighAnsi?.Value ?? style.StyleRunProperties.RunFonts?.ComplexScript?.Value ?? style.StyleRunProperties.RunFonts?.EastAsia?.Value;
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
    }
}