using TemplateFiller;
using TemplateFiller.Extensions;

namespace ConsoleApp
{
    public static partial class Examples
    {
        public static class UsingAsync
        {
            public static async Task RunAsync()
            {
                var data = UsingClass.TestData;

                await Filler.Excel.FillTemplateAsync("Templates\\Test1.xlsx", $"temp//{nameof(UsingAsync)}//output.xlsx", data);
                await Filler.Word.FillTemplateAsync("Templates\\Test2.docx", $"temp//{nameof(UsingAsync)}//output.docx", data);
            }
        }
    }
}
