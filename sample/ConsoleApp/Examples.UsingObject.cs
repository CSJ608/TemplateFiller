using TemplateFiller;
using TemplateFiller.Extensions;

namespace ConsoleApp
{
    public static partial class Examples
    {
        public static class UsingObject
        {
            public static void Run()
            {
                var data = new Dictionary<string, object>()
                {
                    ["StartTime"] = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"),
                    ["EndTime"] = DateTime.Now.ToString("yyyy-MM-dd"),
                    ["PrintTime"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    ["Persons"] = new[]{
                        new { Info = new { Name = "张三", Age = 18, Sex = "男" }, WorkNo = "20001",  Description = "123"},
                        new { Info = new { Name = "张1", Age = 20, Sex = "女" }, WorkNo = "30001",  Description = ""},
                        new { Info = new { Name = "张2", Age = 30, Sex = "男" }, WorkNo = "40001",  Description = ""},
                        new { Info = new { Name = "张3", Age = 40, Sex = "女" }, WorkNo = "50001",  Description = ""},
                        new { Info = new { Name = "张4", Age = 50, Sex = "男" }, WorkNo = "60001",  Description = ""},
                    },
                    ["Counts"] = new int[] { 1, 2, 3 }
                };

                Filler.Excel.FillTemplate("Templates\\Test1.xlsx", $"temp//{nameof(UsingObject)}//output.xlsx", data);
                Filler.Word.FillTemplate("Templates\\Test2.docx", $"temp//{nameof(UsingObject)}//output.docx", data);
            }
        }
    }

}
