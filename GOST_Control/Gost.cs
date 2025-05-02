using System;
using System.Collections.Generic;

namespace GOST_Control
{
    /// <summary>
    /// Модель данных ГОСТа
    /// </summary>
    public partial class Gost
    {
        // ======================= ОСНОВНЫЕ ДАННЫЕ ГОСТА =======================
        public int GostId { get; set; } // ID ГОСТа
        public string Name { get; set; } = null!; // Наименование ГОСТа
        public string? Description { get; set; } // Описание ГОСТа


        // ======================= ПРОСТОЙ ТЕКСТ (ОСНОВНОЕ СОДЕРЖИМОЕ) =======================
        public string? FontName { get; set; } // Шрифт-тип простого текста ГОСТа
        public double? FontSize { get; set; } // Шрифт-размер простого текста ГОСТа
        public double? LineSpacing { get; set; } // Межстрочный интервал
        public string? TextIndentOrOutdent { get; set; } // Отступ или Выступ
        public double? FirstLineIndent { get; set; } // Отступ первой строки
        public double? IndentLeftText { get; set; } // Отступ слева
        public double? IndentRightText { get; set; } // Отступ Справа
        public string? TextAlignment { get; set; } // Выравнивание текста
        public string? LineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно"
        public double? LineSpacingValue { get; set; } // Значение интервала
        public double? LineSpacingBefore { get; set; } // Интервал перед
        public double? LineSpacingAfter { get; set; } // Интервал после


        // ======================= ОТСТУПЫ В ДОКУМЕНТЕ =======================
        public double? MarginTop { get; set; } // Отступ верхний в документе
        public double? MarginBottom { get; set; } // Отступ нижний в документе
        public double? MarginLeft { get; set; } // Отступ левый в документе
        public double? MarginRight { get; set; } // Отступ правый в документе


        // ======================= НУМЕРАЦИЯ СТРАНИЦ =======================
        public bool? PageNumbering { get; set; } // Нумерация страниц (включена/выключена)
        public string? PageNumberingAlignment { get; set; } // Выравнивание нумерации (Left/Center/Right)
        public string? PageNumberingPosition { get; set; } // Положение нумерации (Top/Bottom)


        // ======================= ЗАГОЛОВКИ И РАЗДЕЛЫ =======================
        public string? RequiredSections { get; set; } // Обязательные разделы (Введение, Заключение и т.д.)
        public string? HeaderFontName { get; set; } // Шрифт заголовков
        public double? HeaderFontSize { get; set; } // Размер шрифта заголовков
        public string? HeaderAlignment { get; set; } // Выравнивание заголовков (Center/Left/Right)
        public double? HeaderLineSpacing { get; set; } // Межстрочный интервал заголовков 
        public string? HeaderIndentOrOutdent { get; set; } // Отступ или Выступ
        public double? HeaderFirstLineIndent { get; set; } // Отступ первой строки заголовков 
        public double? HeaderIndentLeft { get; set; } // Отступ слева
        public double? HeaderIndentRight { get; set; } // Отступ справа
        public string? HeaderLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" 
        public double? HeaderLineSpacingValue { get; set; } // Значение интервала 
        public double? HeaderLineSpacingBefore { get; set; } // Интервал перед 
        public double? HeaderLineSpacingAfter { get; set; } // Интервал после 


        // ======================= ФОРМАТ И ОРИЕНТАЦИЯ СТРАНИЦЫ =======================
        public string? PaperSize { get; set; } // Формат бумаги (A4, A5 и т.д.)
        public string? PageOrientation { get; set; } // Ориентация (Portrait/Book - книжная, Landscape - альбомная)
        public double? PaperWidthMm { get; set; } // Ширина страницы в мм
        public double? PaperHeightMm { get; set; } // Высота страницы в мм


        // ======================= ОГЛАВЛЕНИЕ (TOC) =======================
        public bool? RequireTOC { get; set; } // Наличие оглавления 
        public string? TocFontName { get; set; } // Шрифт оглавления 
        public double? TocFontSize { get; set; } // Размер шрифта оглавления
        public string? TocAlignment { get; set; } // Выравнивание оглавления
        public double? TocLineSpacing { get; set; } // Межстрочный интервал оглавления
        public double? TocFirstLineIndent { get; set; } //
        public string? TocIndentOrOutdent { get; set; } //
        public double? TocIndentLeft { get; set; } //
        public double? TocIndentRight { get; set; } //
        public string? TocLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" 
        public double? TocLineSpacingValue { get; set; } // Значение интервала 
        public double? TocLineSpacingBefore { get; set; } // Интервал перед 
        public double? TocLineSpacingAfter { get; set; } // Интервал после 


        // ======================= СПИСКИ (МАРКИРОВАННЫЕ/НУМЕРОВАННЫЕ) =======================
        public bool? RequireBulletedLists { get; set; } // Наличие маркированных списков
        public double? ListHangingIndent { get; set; } // Отступ для списков
        public double? BulletLineSpacing { get; set; } // Межстрочный интервал для списков
        public string? BulletFontName { get; set; } // Шрифт маркированных списков
        public double? BulletFontSize { get; set; } // Размер шрифта маркированных списков
        public string? BulletIndentOrOutdent { get; set; } // 
        public double? BulletIndentLeft { get; set; } // 
        public double? BulletIndentRight { get; set; } // 
        public string? BulletLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" 
        public double? BulletLineSpacingValue { get; set; } // Значение интервала 
        public double? BulletLineSpacingBefore { get; set; } // Интервал перед 
        public double? BulletLineSpacingAfter { get; set; } // Интервал после


        // ===== ДЛЯ МНОГОУРОВНЕВЫХ СПИСКОВ =====
        public double? ListLevel1Indent { get; set; }  // Отступ 1-го уровня (в см)
        public double? ListLevel2Indent { get; set; }  // Отступ 2-го уровня
        public double? ListLevel3Indent { get; set; }  // Отступ 3-го уровня

        public string? ListLevel1NumberFormat { get; set; }  // Формат нумерации (например, "1.", "a)")
        public string? ListLevel2NumberFormat { get; set; }
        public string? ListLevel3NumberFormat { get; set; }

        public string? ListLevel1IndentOrOutdent { get; set; } // 
        public string? ListLevel2IndentOrOutdent { get; set; } // 
        public string? ListLevel3IndentOrOutdent { get; set; } // 
    }
}