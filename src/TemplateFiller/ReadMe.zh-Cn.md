# 什么是模板填充器
模板填充器的功能是在模板中查找占位符，识别占位符的类别，然后根据占位符的路径，从数据源读取数据，并替换占位符。

支持的模板：
- .docx
- .xlsx
- .xls

支持的数据格式
- 文本
- 图像

## 占位符说明

### 简单值占位符

使用":"分隔不同层级。在替换占位符时，仅使用匹配的值替换占位符。

```csharp
// 在占位符处填充一个值，值来源为 Source.Name
var p1 = "{Name}";

// 在占位符处填充一个值，值来源为 Source.Customer.Name
var p2 = "{Customer:Name}";

// 在占位符处填充一个值，值来源为 Source.Project.Customer.Name
var p3 = "{Project:Customer:Name}";
```

#### Word 示例

模板：

![WordFillingTextTemplate](https://raw.githubusercontent.com/CSJ608/TemplateFiller/main/raw/word_filling_text_template.png)

示例程序：

```csharp
using FillingTextExample;
using TemplateFiller.Extensions;

var testData = new DataSource()
{
    StartTime = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"),
    EndTime = DateTime.Now.ToString("yyyy-MM-dd"),
    UserName = "Admin",
    PrintTime = DateTime.Now,
    V1 = "Value1",
    V2 = 12345
};

TemplateFiller.Filler.Word.FillTemplate("Template.docx", "temp\\output.docx", testData);
```

结果：

![WordFillingTextResult](https://raw.githubusercontent.com/CSJ608/TemplateFiller/main/raw/word_filling_text_result.png)

#### Excel 模板示例

### 数组占位符
使用":"分隔不同层级，并且允许使用一次"."，声明可枚举集合每个元素应取哪个字段

#### 示例
```csharp
// 在占位符及其下方，填充多个值，值来源为 Source.Names
var p1 = "[Names]";

// 在占位符及其下方，填充多个值，值来源为 Source.Student.Names
var p2 = "[Student:Names]";

// 在占位符及其下方，填充多个值，值来源为 Source.Students.Select(item => item.Name)
var p3 = "[Students.Name]";

// 在占位符及其下方，填充多个值，值来源为 Source.Students.Select(item => item.Info.Name)
var p4 = "[Students.Info:Name]";

// 在占位符及其下方，填充多个值，值来源为 Source.Project.Students.Select(item => item.Name)
var p5 = "[Project:Students.Name]";

// 在占位符及其下方，填充多个值，值来源为 Source.Project.Students.Select(item => item.Info.Name)
var p6 = "[Project:Students.Info:Name]";
```

#### 注意

在填充数组占位符时，会向数组占位符下方依次填充数据源中的值。填充时若发现，待填充的行存在合并单元格，或者存在非空文本，则会在该行上方插入新的一行，在新增的行中填充数据。

### 图片占位符

图片占位符


## 数据源说明




# 开始使用

假设我按照入职时间，筛选了若干员工数据，然后需要导出一个表，体现出筛选的时间范围、导出日期和员工数据。

数据格式为：
```csharp
public class DataSource
{
    public string StartTime { get; set; } = string.Empty;
    public string EndTime { get; set; } = string.Empty;
    public string PrintTime { get; set; } = string.Empty;
    public Person[] Persons { get; set; } = [];
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Sex { get; set; } = string.Empty;
    public string WorkNo { get; set; } = string.Empty;
    public string Description {  get; set; } = string.Empty;
}
```

## Excel

定义模板：
![Excel Template Input](https://raw.githubusercontent.com/CSJ608/TemplateFiller/main/raw/excel_template.png)

```csharp
var filler = new NpoiExcelTemplateFiller();
var data = new DataSource()
{
    StartTime = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"),
    EndTime = DateTime.Now.ToString("yyyy-MM-dd"),
    PrintTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
    Persons = new Person[]
    {
        new (){ Name = "张0", Age = 20, Sex = "男", WorkNo = "20001", Description = "456" },
        new (){ Name = "张1", Age = 30, Sex = "女", WorkNo = "30001", Description = "" },
        new (){ Name = "张2", Age = 40, Sex = "男", WorkNo = "40001", Description = "789" },
        new (){ Name = "张3", Age = 50, Sex = "女", WorkNo = "50001", Description = "" },
        new (){ Name = "张4", Age = 60, Sex = "男", WorkNo = "60001", Description = "123" }
    }
};

// 填充模板
Filler.Excel.FillTemplate("Template.xlsx", "output.xlsx", data);
```

填充后的效果：
![Excel Template Output](https://raw.githubusercontent.com/CSJ608/TemplateFiller/main/raw/excel_template_filled_result.png)