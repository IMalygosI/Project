﻿using Avalonia.Media;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GOST_Control
{
    /// <summary>
    /// Класс проверок настроек документа "Формат"
    /// </summary>
    public class CheckingSettingDoc
    {
        private readonly WordprocessingDocument _wordDoc;
        private readonly Gost _gost;

        public CheckingSettingDoc(WordprocessingDocument wordDoc, Gost gost)
        {
            _wordDoc = wordDoc;
            _gost = gost;
        }

        /// <summary>
        /// Проверка нумерации страниц и её расположения
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="requiredNumbering"></param>
        /// <param name="requiredAlignment"></param>
        /// <param name="requiredPosition"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<string> Errors)> CheckPageNumberingAsync(WordprocessingDocument wordDoc, bool requiredNumbering, 
                                                           string requiredAlignment, string requiredPosition, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                var foundNumberings = new List<(string Position, string Alignment)>();

                void CheckNumberingInParagraph(Paragraph paragraph, string position)
                {
                    // Проверяем наличие нумерации
                    bool hasNumbering = paragraph.Descendants<SimpleField>().Any(f => f.Instruction?.Value?.Contains("PAGE") == true) ||
                                        paragraph.Descendants<FieldCode>().Any(f => f.Text.Contains("PAGE")) ||
                                        paragraph.Descendants<Run>().Any(r => int.TryParse(r.InnerText?.Trim(), out _));

                    if (hasNumbering)
                    {
                        var (alignment, isDefined) = GetActualAlignment(paragraph, null);
                        foundNumberings.Add((position, alignment));
                    }
                }

                // Проверка верхних колонтитулов
                if (wordDoc.MainDocumentPart.HeaderParts != null)
                {
                    foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                    {
                        foreach (var paragraph in headerPart.Header.Descendants<Paragraph>())
                        {
                            CheckNumberingInParagraph(paragraph, "Top");
                        }
                    }
                }

                // Проверка нижних колонтитулов
                if (wordDoc.MainDocumentPart.FooterParts != null)
                {
                    foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                    {
                        foreach (var paragraph in footerPart.Footer.Descendants<Paragraph>())
                        {
                            CheckNumberingInParagraph(paragraph, "Bottom");
                        }
                    }
                }

                // Анализ результатов
                if (!requiredNumbering)
                {
                    updateUI?.Invoke(foundNumberings.Any() ? "⚠ Найдена нумерация, но не требуется" : 
                                                             "Нумерация не требуется", foundNumberings.Any() ? Brushes.Orange : Brushes.Green);
                    return (!foundNumberings.Any(), tempErrors);
                }

                if (!foundNumberings.Any())
                {
                    updateUI?.Invoke("❌ Нумерация не найдена", Brushes.Red);
                    tempErrors.Add("Нумерация страниц отсутствует");
                    return (false, tempErrors);
                }

                // Проверяем соответствие требованиям
                var correctNumbering = foundNumberings .Where(n => (string.IsNullOrEmpty(requiredPosition) ||
                                                              n.Position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase)) && (string.IsNullOrEmpty(requiredAlignment) ||
                                                              n.Alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase))).ToList();

                if (correctNumbering.Any())
                {
                    var incorrectNumbering = foundNumberings.Except(correctNumbering).ToList();
                    if (incorrectNumbering.Any())
                    {
                        var details = string.Join(", ", incorrectNumbering.Select(n => $"{n.Position} ({n.Alignment})"));
                        updateUI?.Invoke($"✓ Основная нумерация правильная, но есть ошибки: {details}", Brushes.Orange);
                    }
                    else
                    {
                        updateUI?.Invoke($"✓ Нумерация правильная: {correctNumbering.First().Position} {correctNumbering.First().Alignment}", Brushes.Green);
                    }
                }
                else
                {
                    var foundDetails = string.Join(", ", foundNumberings.Select(n => $"{n.Position} ({n.Alignment})"));
                    updateUI?.Invoke($"❌ Несоответствие: требуется {requiredPosition} {requiredAlignment}, найдено {foundDetails}", Brushes.Red);
                    tempErrors.Add($"Нумерация не соответствует: требуется {requiredPosition} {requiredAlignment}");
                    isValid = false;
                }

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Метод для получения выравнивания
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="paragraphStyle"></param>
        /// <returns></returns>
        private (string Alignment, bool IsDefined) GetActualAlignment(Paragraph paragraph, Style paragraphStyle)
        {
            // 1. Проверяем явное выравнивание в параграфе
            if (paragraph.ParagraphProperties?.Justification?.Val?.Value != null)
            {
                return (GetAlignmentString(paragraph.ParagraphProperties.Justification), true);
            }

            // 2. Проверяем стиль параграфа
            var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (paraStyleId != null)
            {
                var style = GetStyleById(paraStyleId);
                while (style != null)
                {
                    if (style.StyleParagraphProperties?.Justification?.Val?.Value != null)
                    {
                        return (GetAlignmentString(style.StyleParagraphProperties.Justification), true);
                    }
                    style = style.BasedOn?.Val?.Value != null ? GetStyleById(style.BasedOn.Val.Value) : null;
                }
            }

            // 3. Проверяем стиль Normal
            var normalStyle = GetStyleById("Normal");
            if (normalStyle?.StyleParagraphProperties?.Justification?.Val?.Value != null)
            {
                return (GetAlignmentString(normalStyle.StyleParagraphProperties.Justification), true);
            }

            return ("Left", true);
        }

        /// <summary>
        /// Вспомогательный метод для получения стиля по ID
        /// </summary>
        /// <param name="styleId"></param>
        /// <returns></returns>
        private Style GetStyleById(string styleId)
        {
            var stylesPart = _wordDoc.MainDocumentPart.StyleDefinitionsPart;
            return stylesPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
        }

        /// <summary>
        /// Проверка размера бумаги на соответствие ГОСТу
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<string> Errors)> CheckPaperSizeAsync(WordprocessingDocument doc, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                var sectPr = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
                if (sectPr == null)
                {
                    updateUI?.Invoke("Не удалось найти раздел документа для проверки", Brushes.Red);
                    tempErrors.Add("Не удалось найти раздел документа для проверки формата бумаги");
                    return (false, tempErrors);
                }

                var pgSz = sectPr.Elements<PageSize>().FirstOrDefault();
                if (pgSz == null)
                {
                    updateUI?.Invoke("Не найден элемент PageSize", Brushes.Red);
                    tempErrors.Add("Не найден элемент PageSize");
                    return (false, tempErrors);
                }

                double widthMm = pgSz.Width.Value / 1440.0 * 25.4;
                double heightMm = pgSz.Height.Value / 1440.0 * 25.4;

                double docWidth = Math.Min(widthMm, heightMm);
                double docHeight = Math.Max(widthMm, heightMm);

                double gostWidthMm = (gost.PaperWidthMm ?? 0) * 10;
                double gostHeightMm = (gost.PaperHeightMm ?? 0) * 10;

                double gostWidth = Math.Min(gostWidthMm, gostHeightMm);
                double gostHeight = Math.Max(gostWidthMm, gostHeightMm);

                bool isCorrectSize = Math.Abs(docWidth - gostWidth) <= 1 && Math.Abs(docHeight - gostHeight) <= 1;

                if (isCorrectSize)
                {
                    updateUI?.Invoke($"Формат бумаги: {gost.PaperSize} ({docWidth:F1}×{docHeight:F1} мм)", Brushes.Green);
                }
                else
                {
                    updateUI?.Invoke($"Требуется {gost.PaperSize} ({gostWidth:F1}×{gostHeight:F1} мм), текущий: {docWidth:F1}×{docHeight:F1} мм", Brushes.Red);
                    tempErrors.Add($"Размер бумаги не соответствует ГОСТу");
                }

                return (isCorrectSize, tempErrors);
            });
        }

        /// <summary>
        /// Проверка ориентации листа
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gost"></param>
        /// <param name="updateUI"></param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<string> Errors)> CheckPageOrientationAsync(WordprocessingDocument doc, Gost gost, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                var sectPr = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
                if (sectPr == null)
                {
                    updateUI?.Invoke("Не удалось найти раздел документа для проверки ориентации", Brushes.Red);
                    tempErrors.Add("Не удалось найти раздел документа для проверки ориентации");
                    return (false, tempErrors);
                }

                var pgSz = sectPr.Elements<PageSize>().FirstOrDefault();
                if (pgSz == null)
                {
                    updateUI?.Invoke("Не найден элемент PageSize", Brushes.Red);
                    tempErrors.Add("Не найден элемент PageSize");
                    return (false, tempErrors);
                }

                bool isPortrait = pgSz.Orient == null || pgSz.Orient.Value == PageOrientationValues.Portrait;
                bool shouldBePortrait = gost.PageOrientation == "Portrait";

                if (isPortrait == shouldBePortrait)
                {
                    updateUI?.Invoke($"Ориентация: {(shouldBePortrait ? "Книжная" : "Альбомная")} (соответствует)", Brushes.Green);
                }
                else
                {
                    updateUI?.Invoke($"Ориентация: {(isPortrait ? "Книжная" : "Альбомная")} (должна быть {(shouldBePortrait ? "Книжная" : "Альбомная")})", Brushes.Red);
                    tempErrors.Add($"Ориентация не соответствует ГОСТу");
                    isValid = false;
                }

                return (isValid, tempErrors);
            });
        }

        /// <summary>
        /// Проверка полей документа
        /// </summary>
        /// <param name="requiredMarginTop"></param>
        /// <param name="requiredMarginBottom"></param>
        /// <param name="requiredMarginLeft"></param>
        /// <param name="requiredMarginRight"></param>
        /// <param name="body"></param>
        /// <param name="updateUI">Функция для обновления UI</param>
        /// <returns></returns>
        public async Task<(bool IsValid, List<string> Errors)> CheckMarginsAsync(double? requiredMarginTop, double? requiredMarginBottom, double? requiredMarginLeft, double? requiredMarginRight, Body body, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                var pageMargin = body.Elements<SectionProperties>().FirstOrDefault()?.Elements<PageMargin>().FirstOrDefault();

                if (pageMargin == null)
                {
                    updateUI?.Invoke("Не удалось найти поля документа для проверки", Brushes.Red);
                    tempErrors.Add("Не удалось найти поля документа для проверки");
                    return (false, tempErrors);
                }

                // Преобразование в сантиметры (1 см = 567 twips)
                double marginTopInCm = pageMargin.Top.Value / 567.0;
                double marginBottomInCm = pageMargin.Bottom.Value / 567.0;
                double marginLeftInCm = pageMargin.Left.Value / 567.0;
                double marginRightInCm = pageMargin.Right.Value / 567.0;

                // Проверка с погрешностью 0.01 см
                if (requiredMarginTop.HasValue && Math.Abs(marginTopInCm - requiredMarginTop.Value) > 0.01)
                {
                    isValid = false;
                    tempErrors.Add($"Верхнее поле не соответствует ГОСТу. Требуется: {requiredMarginTop.Value} см, текущее: {marginTopInCm:F2} см.");
                }

                if (requiredMarginBottom.HasValue && Math.Abs(marginBottomInCm - requiredMarginBottom.Value) > 0.01)
                {
                    isValid = false;
                    tempErrors.Add($"Нижнее поле не соответствует ГОСТу. Требуется: {requiredMarginBottom.Value} см, текущее: {marginBottomInCm:F2} см.");
                }

                if (requiredMarginLeft.HasValue && Math.Abs(marginLeftInCm - requiredMarginLeft.Value) > 0.01)
                {
                    isValid = false;
                    tempErrors.Add($"Левое поле не соответствует ГОСТу. Требуется: {requiredMarginLeft.Value} см, текущее: {marginLeftInCm:F2} см.");
                }

                if (requiredMarginRight.HasValue && Math.Abs(marginRightInCm - requiredMarginRight.Value) > 0.01)
                {
                    isValid = false;
                    tempErrors.Add($"Правое поле не соответствует ГОСТу. Требуется: {requiredMarginRight.Value} см, текущее: {marginRightInCm:F2} см.");
                }

                // Обновление UI
                if (isValid)
                {
                    updateUI?.Invoke("Поля документа соответствуют ГОСТу.", Brushes.Green);
                }
                else
                {
                    updateUI?.Invoke("Поля документа не соответствуют ГОСТу.", Brushes.Red);
                }

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
    }
}
