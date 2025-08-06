# 什么是模板填充
该工具的主要功能是在模板中查找占位符，根据占位符的路径，从数据源读取数据，并替换占位符。

## 模板支持情况

- [x] Excel (.xls, .xlsx)
- [ ] Word (.doc, .docx)

## 可填充的数据

- [x] 文本
- [x] 表格
- [ ] 图像 

## 占位符说明

### 直接替换的文本
用 <kbd>{</kbd>Path<kbd>}</kbd> 标识。例如：

- {StartTime}：将数据源中路径为"StartTime"的数据填充至占位符处
- {Person:Name}：将数据源中路径为"Person:Name"的数据填充至占位符处

### 填充表格
用 <kbd>[</kbd>Path<kbd>]</kbd> 标识。例如：

- [Persons:Name]：遍历数据源中路径为"Persons"的集合，并尝试获取每个元素的"Name"属性的值，然后从占位符所处位置开始，向下依次填充属性值。如果向下填充的过程中发现任意文本，则插入新的一行，避免覆盖已有文本。

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
![Excel Template Input](https://github.com/CSJ608/TemplateFiller/main/src/TemplateFiller/image.png?raw=true)

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
filler.FillTemplate("Template.xlsx", "output.xlsx", data);
```

填充后的效果：
![Excel Template Output](https://github.com/CSJ608/TemplateFiller/main/src/TemplateFiller/image-1.png?raw=true)