# Excely
一個基於 .NET 6 的套件，用於簡化各種常用資料結構與 Excel、Csv 之間的資料匯出及匯入。

## 功能
- 簡單的 API，用最快的速度上手並使用！
- 完整的參數設定與自定義功能，每個環節都可以客製化。
- 提供多種 Shader，用於進一步美化或調整匯出結果。

## 使用範例

以下是一個簡單的使用範例，展示如何將一個 `Student` 的 List 匯出為 Excel 檔案，再匯入回 List。

```csharp
static void Main(string[] args)
{
    var students = new List<Student>()
    {
        new Student(0, "Test1", DateTime.Now),
        new Student(1, "Test2", DateTime.Now),
    };

    // 匯出為 Excel
    var exporter = ExcelyExporter.FromClassList<Student>();
    using var excel = exporter.ToExcel(students);

    // 匯入為 List<Student>
    var importer = new XlsxImporterBuilder().BuildForClassList<Student>();
    IEnumerable<Student> importResult = importer.Import(excel.Workbook.Worksheets.First());
}
```

您可以瀏覽 [開始使用 Excely](https://github.com/maxzh1999tw/Excely/wiki/%E9%96%8B%E5%A7%8B%E4%BD%BF%E7%94%A8-Excely) 來了解更多。

## 貢獻
可以先到 [Wiki](https://github.com/maxzh1999tw/Excely/wiki) 大致了解一下專案資訊。  
[Issue](https://github.com/maxzh1999tw/Excely/issues) 頁面可能有些懸賞任務需要您的協助。  
如果您有任何建議或發現任何問題，也歡迎開啟 issue 或提交 pull request。

## 授權

本套件使用 Apache-2.0 License 授權，詳情請參見 [LICENSE](LICENSE)。
