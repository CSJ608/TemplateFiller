using TemplateFiller;
using TemplateFiller.Extensions;

namespace ConsoleApp
{
    public static partial class Examples
    {
        public static class UsingClass
        {
            public static void Run()
            {
                Filler.Excel.FillTemplate("Templates\\Test1.xlsx", $"temp//{nameof(UsingClass)}//output.xlsx", TestData);
                Filler.Word.FillTemplate("Templates\\Test2.docx", $"temp//{nameof(UsingClass)}//output.docx", TestData);
            }

            public static DataSource TestData => new()
            {
                StartTime = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"),
                EndTime = DateTime.Now.ToString("yyyy-MM-dd"),
                PrintTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                Persons = [
                        new (){ Info = new() { Name = "张0", Age = 20, Sex = "男" }, WorkNo = "20001", Description = "456" },
                        new (){ Info = new() { Name = "张1", Age = 30, Sex = "女" }, WorkNo = "30001", Description = "" },
                        new (){ Info = new() { Name = "张2", Age = 40, Sex = "男" }, WorkNo = "40001", Description = "789" },
                        new (){ Info = new() { Name = "张3", Age = 50, Sex = "女" }, WorkNo = "50001", Description = "" },
                        new (){ Info = new() { Name = "张4", Age = 60, Sex = "男" }, WorkNo = "60001", Description = "123" }
                    ],
                Counts = [1, 2, 3]
            };
        }
    }

}
