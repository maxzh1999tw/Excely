# Excely

一個基於 .NET 6 的套件，用於簡化從各種常用資料結構到 Excel 或 Csv 的資料匯出，以及從 Excel、Csv 到常用資料結構的資料匯入。

## 功能

- 自訂資料匯出規則
- 提供多種 Shader，用於進一步美化或調整 Excel 檔案
- 簡單的 API，方便快速上手

## 使用範例

以下是一個簡單的使用範例，展示如何將一個 `Student` 的 List 匯出為 Excel 檔案。

```csharp
var students = new List<Student>()
{
    new Student(0, "Test1", DateTime.Parse("2020/04/17")),
    new Student(1, "Test2", DateTime.Parse("1964/10/11")),
};

var exporter = new ClassListExporter<Student>
{
    CustomValuePolicy = (student, property) => property.Name switch
    {
        nameof(Student.Birthday) => student.Birthday.ToString("yyyy/MM/dd"),
        _ => property.GetValue(student),
    },
    Shaders = new IShader[]
    {
        new CellFittingShader()
    }
};

using var excel = exporter.ToExcel(students);
excel.SaveAs(new FileInfo("TestFile.xlsx"));
```

## 貢獻

Issue 頁面可能有些懸賞任務需要您的協助，
如果您有任何建議或發現任何問題，也歡迎開啟 issue 或提交 pull request。

## 授權

本套件使用 Apache License 授權，詳情請參見 [LICENSE](LICENSE)。
