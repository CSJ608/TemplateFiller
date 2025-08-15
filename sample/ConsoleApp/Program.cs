// See https://aka.ms/new-console-template for more information
using ConsoleApp;
using TemplateFiller;

Console.WriteLine("Hello, World!");

var data = new Dictionary<string, object>()
{
    ["StartTime"] = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"),
    ["EndTime"] = DateTime.Now.ToString("yyyy-MM-dd"),
    ["PrintTime"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
    ["Persons"] = new[]
    {
        new { Info = new { Name = "张三", Age = 18, Sex = "男" }, WorkNo = "20001",  Description = "123"},
        new { Info = new { Name = "张1", Age = 20, Sex = "女" }, WorkNo = "30001",  Description = ""},
        new { Info = new { Name = "张2", Age = 30, Sex = "男" }, WorkNo = "40001",  Description = ""},
        new { Info = new { Name = "张3", Age = 40, Sex = "女" }, WorkNo = "50001",  Description = ""},
        new { Info = new { Name = "张4", Age = 50, Sex = "男" }, WorkNo = "60001",  Description = ""},
    },
    ["Counts"] = new int[] {1, 2, 3}
};

var filler = new NpoiExcelTemplateFiller();
filler.FillTemplate("Templates\\Test1.xlsx", "output.xlsx", data);

var data2 = new DataSource()
{
    StartTime = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"),
    EndTime = DateTime.Now.ToString("yyyy-MM-dd"),
    PrintTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
    Persons = new Person[]
    {
        new (){ Info = new() { Name = "张0", Age = 20, Sex = "男" }, WorkNo = "20001", Description = "456" },
        new (){ Info = new() { Name = "张1", Age = 30, Sex = "女" }, WorkNo = "30001", Description = "" },
        new (){ Info = new() { Name = "张2", Age = 40, Sex = "男" }, WorkNo = "40001", Description = "789" },
        new (){ Info = new() { Name = "张3", Age = 50, Sex = "女" }, WorkNo = "50001", Description = "" },
        new (){ Info = new() { Name = "张4", Age = 60, Sex = "男" }, WorkNo = "60001", Description = "123" }
    },
    Counts = new int[]{1, 2, 3}
};

filler.FillTemplate("Templates\\Test1.xlsx", "output2.xlsx", data2);

var wordFiller = new NpoiWordTemplateFiller();
wordFiller.FillTemplate("Templates\\Test2.docx", "output.docx", data);
wordFiller.FillTemplate("Templates\\Test2.docx", "output2.docx", data);