using TemplateFiller;
using TemplateFiller.Abstractions;
using TemplateFiller.Extensions;

namespace ConsoleApp
{
    public static partial class Examples
    {
        public static class BatchPrint
        {
            public static void Run()
            {
                var data = UsingClass.TestData;

                var excelNames = BuildNames("xlsx");
                Filler.Excel.FillTemplate("Templates\\Test1.xlsx", excelNames.Select(t => new FileBag(t, data)));

                var wordNames = BuildNames("docx");
                Filler.Word.FillTemplate("Templates\\Test2.docx", wordNames.Select(t => new FileBag(t, data)));
            }

            private static List<string> BuildNames(string extension)
            {
                var names = new List<string>();
                for (int i = 0; i < 10; i++)
                {
                    names.Add($"temp//{nameof(BatchPrint)}//output_{i}.{extension}");
                }

                return names;
            }
        }
    }

}
