# 什么是模板填充器
模板填充器的功能是在模板中查找占位符，识别占位符的类别，然后根据占位符的路径，从数据源读取数据，并替换占位符。

## 模板支持

- [x] Excel (.xls, .xlsx)
- [x] Word (.docx)

## 填充的数据

- [x] 文本
- [x] 表格
- [ ] 图像 

## 占位符说明

### 简单值占位符
使用":"分隔不同层级。在替换占位符时，仅使用匹配的值替换占位符。

#### 示例

```csharp
// 在占位符处填充一个值，值来源为 Source.Name
var p1 = "{Name}";

// 在占位符处填充一个值，值来源为 Source.Customer.Name
var p2 = "{Customer:Name}";

// 在占位符处填充一个值，值来源为 Source.Project.Customer.Name
var p3 = "{Project:Customer:Name}";
```

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

#### 补充说明

在占位符及其下方填充多个值的过程中，若发现待填充的行存在合并单元格或者任意文本，则会插入一行，避免修改原本的数据。

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