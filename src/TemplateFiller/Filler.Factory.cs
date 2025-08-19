using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateFiller
{
    partial class Filler
    {
        public static Filler Word => DefaultWordFiller;
        public static Filler Excel => DefaultExcelFiller;

        private static readonly Filler DefaultWordFiller = new WordFiller();
        private static readonly Filler DefaultExcelFiller = new ExcelFiller();
    }
}
