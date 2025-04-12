using System;
using System.Collections.Generic;

namespace GOST_Control
{
    /// <summary>
    /// Модель данных ГОСТа
    /// </summary>
    public partial class Gost
    {
        // ID ГОСТа
        public int GostId { get; set; }
        // Наименование ГОСТа
        public string Name { get; set; } = null!;
        // Описание ГОСТа
        public string? Description { get; set; }
        // Шрифт ГОСТа
        public string? FontName { get; set; }
        // Размер шрифта ГОСТа
        public double? FontSize { get; set; }

        // Отступы ГОСТа
        public double? MarginTop { get; set; }
        public double? MarginBottom { get; set; }
        public double? MarginLeft { get; set; }
        public double? MarginRight { get; set; }

        // Нумерация ГОСТа
        public bool? PageNumbering { get; set; }
        // Межстрочный интервал
        public double? LineSpacing { get; set; }
        // Отступ первой строки
        public double? FirstLineIndent { get; set; }        
        // Выравнивание текста
        public string? TextAlignment { get; set; } // Justify - По ширине

        // Оглавления-Заголовки
        public string? RequiredSections { get; set; }        // Поле на проверку Оглавлений-Заголовков (Введение-Заключение-Список литературы)
        public double? HeaderFontSize { get; set; }// Размер шрифта Оглавлений-заголовков       
        public string? HeaderAlignment { get; set; } // Выравнивание заголовков ("По центру", "По левому краю")

    }
}