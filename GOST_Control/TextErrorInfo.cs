using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace GOST_Control
{
    /// <summary>
    /// Класс сбора и передачи ошибок
    /// </summary>
    public class TextErrorInfo
    {
        public string ErrorMessage { get; set; }  // Текст ошибки
        public Run ProblemRun { get; set; }       // Проблемный Run
        public Paragraph ProblemParagraph { get; set; } // Проблемный Paragraph (если ошибка в paragraph)
    }
}