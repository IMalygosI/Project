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
        /// <param name="wordDoc">Документ для проверки</param>
        /// <param name="requiredNumbering">Требуется ли нумерация</param>
        /// <param name="requiredAlignment">Требуемое выравнивание</param>
        /// <param name="requiredPosition">Требуемое положение</param>
        /// <param name="updateUI">Функция для обновления UI</param>
        /// <returns>Результат проверки</returns>
        public async Task<(bool IsValid, List<string> Errors)> CheckPageNumberingAsync(WordprocessingDocument wordDoc, bool requiredNumbering, string requiredAlignment, string requiredPosition, Action<string, IBrush> updateUI)
        {
            return await Task.Run(() =>
            {
                var tempErrors = new List<string>();
                bool isValid = true;

                bool hasCorrectNumbering = false;
                bool hasExtraNumbering = false;
                string actualCorrectPosition = "";
                string actualCorrectAlignment = "";
                List<string> extraNumberings = new List<string>();

                // Проверка верхних колонтитулов
                if (wordDoc.MainDocumentPart.HeaderParts != null)
                {
                    foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                    {
                        foreach (var paragraph in headerPart.Header.Elements<Paragraph>())
                        {
                            var pageField = paragraph.Descendants<SimpleField>().FirstOrDefault(f => f.Instruction?.Value?.Contains("PAGE") == true);

                            if (pageField != null)
                            {
                                var justification = paragraph.ParagraphProperties?.Justification;
                                string alignment = GetAlignmentString(justification);
                                string position = "Top";

                                // Проверка на соответствие требованиям
                                bool positionMatch = string.IsNullOrEmpty(requiredPosition) || position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase);
                                bool alignmentMatch = string.IsNullOrEmpty(requiredAlignment) || alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase);

                                if (positionMatch && alignmentMatch)
                                {
                                    hasCorrectNumbering = true;
                                    actualCorrectPosition = position;
                                    actualCorrectAlignment = alignment;
                                }
                                else
                                {
                                    hasExtraNumbering = true;
                                    extraNumberings.Add($"{position}, {alignment}");
                                }
                            }
                        }
                    }
                }

                // Проверка нижних колонтитулов
                if (wordDoc.MainDocumentPart.FooterParts != null)
                {
                    foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                    {
                        foreach (var paragraph in footerPart.Footer.Elements<Paragraph>())
                        {
                            var pageField = paragraph.Descendants<SimpleField>().FirstOrDefault(f => f.Instruction?.Value?.Contains("PAGE") == true);

                            if (pageField != null)
                            {
                                var justification = paragraph.ParagraphProperties?.Justification;
                                string alignment = GetAlignmentString(justification);
                                string position = "Bottom";

                                // Проверка на соответствие требованиям
                                bool positionMatch = string.IsNullOrEmpty(requiredPosition) || position.Equals(requiredPosition, StringComparison.OrdinalIgnoreCase);
                                bool alignmentMatch = string.IsNullOrEmpty(requiredAlignment) || alignment.Equals(requiredAlignment, StringComparison.OrdinalIgnoreCase);

                                if (positionMatch && alignmentMatch)
                                {
                                    hasCorrectNumbering = true;
                                    actualCorrectPosition = position;
                                    actualCorrectAlignment = alignment;
                                }
                                else
                                {
                                    hasExtraNumbering = true;
                                    extraNumberings.Add($"{position}, {alignment}");
                                }
                            }
                        }
                    }
                }

                // Формируем сообщение об ошибке
                if (!hasCorrectNumbering && !hasExtraNumbering)
                {
                    updateUI?.Invoke("Нумерация страниц отсутствует", Brushes.Red);
                    tempErrors.Add("Нумерация страниц отсутствует");
                    isValid = false;
                }
                else if (hasCorrectNumbering && !hasExtraNumbering)
                {
                    // Нумерация присутствует и соответствует ГОСТу
                    string message = string.IsNullOrEmpty(requiredAlignment) && string.IsNullOrEmpty(requiredPosition) ? "Нумерация страниц присутствует и соответствует ГОСТу." :
                                                                                             $"Нумерация соответствует ГОСТу ({actualCorrectPosition}, {actualCorrectAlignment})";
                    updateUI?.Invoke(message, Brushes.Green);
                }
                else
                {
                    // Есть ошибка, если есть лишняя нумерация
                    string message = $"Нумерация не соответствует ГОСТу. Требуется: " +
                                     $"Положение: {requiredPosition ?? "не указано"}, " +
                                     $"Выравнивание: {requiredAlignment ?? "не указано"}. \n" +
                                     $"Найдена неверная нумерация: {string.Join("; ", extraNumberings.Select(e => $"Положение: {e.Split(',')[0]}, Выравнивание: {e.Split(',')[1]}"))}";

                    updateUI?.Invoke(message, Brushes.Red);
                    tempErrors.Add("Лишняя нумерация страниц");
                    isValid = false;
                }

                return (isValid, tempErrors);
            });
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
