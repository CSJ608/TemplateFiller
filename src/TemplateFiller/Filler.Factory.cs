namespace TemplateFiller
{
    partial class Filler
    {
        /// <summary>
        /// Word填充器
        /// </summary>
        public static Filler Word => DefaultWordFiller;
        /// <summary>
        /// Excel填充器
        /// </summary>
        public static Filler Excel => DefaultExcelFiller;

        private static readonly Filler DefaultWordFiller = new WordFiller();
        private static readonly Filler DefaultExcelFiller = new ExcelFiller();
    }
}
