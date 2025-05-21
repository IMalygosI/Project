using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace GOST_Control
{
    /// <summary>
    /// Модель данных ГОСТа
    /// </summary>
    public partial class Gost : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // ======================= ОСНОВНЫЕ ДАННЫЕ ГОСТА =======================
        public int GostId { get; set; } // ID ГОСТа
        public string Name { get; set; } = null!; // Наименование ГОСТа
        public string? Description { get; set; } // Описание ГОСТа


        // ======================= ПРОСТОЙ ТЕКСТ (ОСНОВНОЕ СОДЕРЖИМОЕ) =======================

        public string? _fontName; // Шрифт-тип простого текста ГОСТа +
        public string? FontName
        {
            get => _fontName;
            set
            {
                if (_fontName != value)
                {
                    _fontName = value;
                    OnPropertyChanged();
                }
            }
        }

        public double? _fontSize;// Шрифт-размер простого текста ГОСТа +
        public double? FontSize
        {
            get => _fontSize;
            set
            {
                if (_fontSize != value)
                {
                    _fontSize = value;
                    OnPropertyChanged();
                }
            }
        }
        public string? TextAlignment { get; set; } // Выравнивание текста +
        public string? TextIndentOrOutdent { get; set; } // Отступ или Выступ +
        public double? FirstLineIndent { get; set; } // Отступ первой строки +
        public double? IndentLeftText { get; set; } // Отступ слева +
        public double? IndentRightText { get; set; } // Отступ Справа +
        public string? LineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" +
        public double? LineSpacingValue { get; set; } // Значение интервала +
        public double? LineSpacingBefore { get; set; } // Интервал перед +
        public double? LineSpacingAfter { get; set; } // Интервал после +

        // ======================= ОТСТУПЫ В ДОКУМЕНТЕ =======================
        public double? MarginTop { get; set; } // Отступ верхний в документе +
        public double? MarginBottom { get; set; } // Отступ нижний в документе +
        public double? MarginLeft { get; set; } // Отступ левый в документе +
        public double? MarginRight { get; set; } // Отступ правый в документе +


        // ======================= НУМЕРАЦИЯ СТРАНИЦ =======================
        public bool? PageNumbering { get; set; } // Нумерация страниц (включена/выключена) +
        public string? PageNumberingAlignment { get; set; } // Выравнивание нумерации (Left/Center/Right) +
        public string? PageNumberingPosition { get; set; } // Положение нумерации (Top/Bottom) +


        // ======================= ЗАГОЛОВКИ =======================
        public string? RequiredSections { get; set; } // Обязательные разделы (Введение, Заключение и т.д.) -
        public string? HeaderFontName { get; set; } // Шрифт заголовков +
        public double? HeaderFontSize { get; set; } // Размер шрифта заголовков +
        public string? HeaderAlignment { get; set; } // Выравнивание заголовков (Center/Left/Right) +
        public string? HeaderIndentOrOutdent { get; set; } // Отступ или Выступ +
        public double? HeaderFirstLineIndent { get; set; } // Отступ первой строки заголовков +
        public double? HeaderIndentLeft { get; set; } // Отступ слева +
        public double? HeaderIndentRight { get; set; } // Отступ справа +
        public string? HeaderLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" +
        public double? HeaderLineSpacingValue { get; set; } // Значение интервала +
        public double? HeaderLineSpacingBefore { get; set; } // Интервал перед +
        public double? HeaderLineSpacingAfter { get; set; } // Интервал после +


        // ======================= ДОП.ЗАГОЛОВКИ =======================
        public string? AdditionalHeaderFontName { get; set; } // Шрифт дополнительных заголовков  +
        public double? AdditionalHeaderFontSize { get; set; } // Размер шрифта дополнительных заголовков +
        public string? AdditionalHeaderAlignment { get; set; } // Выравнивание дополнительных заголовков (Center/Left/Right) +
        public string? AdditionalHeaderIndentOrOutdent { get; set; } // Отступ или Выступ для дополнительных заголовков +
        public double? AdditionalHeaderFirstLineIndent { get; set; } // Отступ первой строки дополнительных заголовков +
        public double? AdditionalHeaderIndentLeft { get; set; } // Отступ слева для дополнительных заголовков +
        public double? AdditionalHeaderIndentRight { get; set; } // Отступ справа для дополнительных заголовков +
        public string? AdditionalHeaderLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" для дополнительных заголовков +
        public double? AdditionalHeaderLineSpacingValue { get; set; } // Значение интервала для дополнительных заголовков +
        public double? AdditionalHeaderLineSpacingBefore { get; set; } // Интервал перед для дополнительных заголовков +
        public double? AdditionalHeaderLineSpacingAfter { get; set; } // Интервал после для дополнительных заголовков +


        // ======================= ФОРМАТ И ОРИЕНТАЦИЯ СТРАНИЦЫ =======================
        public string? PageOrientation { get; set; } // Ориентация (Portrait - книжная, Landscape - альбомная) +
        public string? PaperSize { get; set; } // Формат бумаги (A4, A5 и т.д.) +
        public double? PaperWidthMm { get; set; } // Ширина страницы в мм +
        public double? PaperHeightMm { get; set; } // Высота страницы в мм +


        // ======================= ОГЛАВЛЕНИЕ (TOC) =======================
        public bool? RequireTOC { get; set; } // Наличие оглавления +
        public string? TocFontName { get; set; } // Шрифт оглавления +
        public double? TocFontSize { get; set; } // Размер шрифта оглавления +
        public string? TocAlignment { get; set; } // Выравнивание оглавления +
        public string? TocIndentOrOutdent { get; set; } //  Отступ или Выступ +
        public double? TocFirstLineIndent { get; set; } // первая строка +
        public double? TocIndentLeft { get; set; } // левая стр +
        public double? TocIndentRight { get; set; } // правая стр +
        public string? TocLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" +
        public double? TocLineSpacing { get; set; } // Межстрочный интервал оглавления+
        public double? TocLineSpacingBefore { get; set; } // Интервал перед +
        public double? TocLineSpacingAfter { get; set; } // Интервал после +



        // ======================= СПИСКИ (МАРКИРОВАННЫЕ/НУМЕРОВАННЫЕ) =======================
        public bool? RequireBulletedLists { get; set; } // Наличие маркированных списков +
        public string? BulletFontName { get; set; } // Шрифт маркированных списков +
        public double? BulletFontSize { get; set; } // Размер шрифта маркированных списков +
        public string? BulletAlignment { get; set; } // Выравнивание Списков +
        public string? ListLevel1IndentOrOutdent { get; set; } //  "Отступ" +
        public double? ListLevel1Indent { get; set; } // Отступ для списков Отступ 1-го уровня +
        public double? ListLevel1BulletIndentLeft { get; set; } //  +
        public double? ListLevel1BulletIndentRight { get; set; } //  +
        public string? BulletLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно"  +
        public double? BulletLineSpacingValue { get; set; } // Значение интервала  +
        public double? BulletLineSpacingBefore { get; set; } // Интервал перед  +
        public double? BulletLineSpacingAfter { get; set; } // Интервал после +

        // ===== ДЛЯ МНОГОУРОВНЕВЫХ СПИСКОВ =====

        // отступ первой строки
        public double? ListLevel2Indent { get; set; } // Отступ 2-го уровня +
        public double? ListLevel3Indent { get; set; } // Отступ 3-го уровня +
        public double? ListLevel4Indent { get; set; } // Отступ 4-го уровня +
        public double? ListLevel5Indent { get; set; } // Отступ 5-го уровня +
        public double? ListLevel6Indent { get; set; } // Отступ 6-го уровня +
        public double? ListLevel7Indent { get; set; } // Отступ 7-го уровня +
        public double? ListLevel8Indent { get; set; } // Отступ 8-го уровня +
        public double? ListLevel9Indent { get; set; } // Отступ 9-го уровня +

        // левый отступ
        public double? ListLevel2BulletIndentLeft { get; set; }// +
        public double? ListLevel3BulletIndentLeft { get; set; }// +
        public double? ListLevel4BulletIndentLeft { get; set; }// +
        public double? ListLevel5BulletIndentLeft { get; set; }// +
        public double? ListLevel6BulletIndentLeft { get; set; }// +
        public double? ListLevel7BulletIndentLeft { get; set; }// +
        public double? ListLevel8BulletIndentLeft { get; set; }// +
        public double? ListLevel9BulletIndentLeft { get; set; }// +

        // правый отступ
        public double? ListLevel2BulletIndentRight { get; set; }// +
        public double? ListLevel3BulletIndentRight { get; set; }// +
        public double? ListLevel4BulletIndentRight { get; set; }// +
        public double? ListLevel5BulletIndentRight { get; set; }// +
        public double? ListLevel6BulletIndentRight { get; set; }// +
        public double? ListLevel7BulletIndentRight { get; set; }// +
        public double? ListLevel8BulletIndentRight { get; set; }// +
        public double? ListLevel9BulletIndentRight { get; set; }// +

        // формат нумерации
        public string? ListLevel1NumberFormat { get; set; }  // Формат нумерации (например, "1.", "a)")
        public string? ListLevel2NumberFormat { get; set; }
        public string? ListLevel3NumberFormat { get; set; }
        public string? ListLevel4NumberFormat { get; set; }
        public string? ListLevel5NumberFormat { get; set; }
        public string? ListLevel6NumberFormat { get; set; }
        public string? ListLevel7NumberFormat { get; set; }
        public string? ListLevel8NumberFormat { get; set; }
        public string? ListLevel9NumberFormat { get; set; }

        // тип первой строки
        public string? ListLevel2IndentOrOutdent { get; set; } // +
        public string? ListLevel3IndentOrOutdent { get; set; } //+
        public string? ListLevel4IndentOrOutdent { get; set; }//+
        public string? ListLevel5IndentOrOutdent { get; set; }//+
        public string? ListLevel6IndentOrOutdent { get; set; }//+
        public string? ListLevel7IndentOrOutdent { get; set; }//+
        public string? ListLevel8IndentOrOutdent { get; set; }//+
        public string? ListLevel9IndentOrOutdent { get; set; }//+


        // ======================= КАРТИНКИ =======================
        public string? ImageCaptionFontName { get; set; } // Шрифт подписей под изображениями +
        public double? ImageCaptionFontSize { get; set; } // Размер шрифта подписей +
        public string? ImageCaptionIndentOrOutdent { get; set; } // Отступ или Выступ +
        public double? ImageCaptionFirstLineIndent { get; set; } // Отступ первой строки подписи +
        public double? ImageCaptionIndentLeft { get; set; } // Отступ слева у подписи +
        public double? ImageCaptionIndentRight { get; set; } // Отступ справа у подписи +
        public string? ImageCaptionAlignment { get; set; } // Выравнивание подписи (Left/Center/Right) +
        public string? ImageCaptionLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" +
        public double? ImageCaptionLineSpacingValue { get; set; } // Значение интервала +
        public double? ImageCaptionLineSpacingBefore { get; set; } // Интервал перед подписью +
        public double? ImageCaptionLineSpacingAfter { get; set; } // Интервал после подписи +


        // ======================= ТАБЛИЦЫ И ПОДПИСИ =======================
        public string? TableCaptionFontName { get; set; } // Шрифт подписей над таблицами +
        public double? TableCaptionFontSize { get; set; } // Размер шрифта подписей +
        public string? TableCaptionIndentOrOutdent { get; set; } // Отступ или Выступ +
        public double? TableCaptionFirstLineIndent { get; set; } // Отступ первой строки подписи +
        public double? TableCaptionIndentLeft { get; set; } // Отступ слева у подписи +
        public double? TableCaptionIndentRight { get; set; } // Отступ справа у подписи +
        public string? TableCaptionAlignment { get; set; } // Выравнивание подписи (Left/Center/Right) +
        public string? TableCaptionLineSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" +
        public double? TableCaptionLineSpacingValue { get; set; } // Значение интервала +
        public double? TableCaptionLineSpacingBefore { get; set; } // Интервал перед подписью +
        public double? TableCaptionLineSpacingAfter { get; set; } // Интервал после подписи +

        // ======================= ТЕКСТ В ТАБЛИЦЕ =======================
        public double? TableFontSize { get; set; } // Шрифт подписей над таблицами +
        public string? TableAlignment { get; set; } // Выравнивание текста в таблицах (Left/Center/Right) +
        public string? TableIndentOrOutdent { get; set; } // Отступ или Выступ +
        public double? TableFirstLineIndent { get; set; } // Отступ первой строки подписи +
        public double? TableIndentLeft { get; set; } // Отступ слева у подписи +
        public double? TableIndentRight { get; set; } // Отступ справа у подписи +
        public string? TableSpacingType { get; set; } // "Множитель"/"Минимум"/"Точно" +
        public double? TableLineSpacingValue { get; set; } // Значение интервала +
        public double? TableLineSpacingBefore { get; set; } // Интервал перед подписью +
        public double? TableLineSpacingAfter { get; set; } // Интервал после подписи +
    }
}