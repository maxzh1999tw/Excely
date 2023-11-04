using Excely.Debug.Models;
using Excely.EPPlus.LGPL.Exporters;
using Excely.EPPlus.LGPL.Shaders.Xlsx;
using Excely.EPPlus.LGPL.TableFactories;
using Excely.EPPlus.LGPL.TableWriters;
using Excely.Shaders;
using Excely.TableImporter;
using System.Reflection;
using OfficeOpenXml;
using Excely.Exporters;
using Excely.TableFactories;

namespace Excely.Debug
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var students = new List<Student>()
            {
                new Student(0, "Test1", DateTime.Parse("2020/04/17")),
                new Student(1, "Test2", null),
            };

            var exporter = ExcelyExporter.FromClassList<Student>();
            using var excel = exporter.ToExcel(students);

            var reader = new XlsxTableFactory();
            var table = reader.GetTable(excel.Workbook.Worksheets.First());
            var importResult = new ClassListTableImporter<Student>()
            {
                StopWhenError = false
            }.Import(table);

            if(importResult.CellErrors.Any())
            {
                var errorRows = table.Data.Where(x => importResult.CellErrors.Any(c => c.Key.Row == table.Data.IndexOf(x))).ToList();
                errorRows.Insert(0, table.Data[0]);
                var errTable = new ExcelyTable(errorRows);
                using var errExcel = new ExcelPackage();
                var sheet = errExcel.Workbook.Worksheets.Add("sheet1");
                new XlsxTableWriter(sheet).Write(errTable);
                new ErrorMarkShader(importResult.CellErrors.ToDictionary(x => x.Key, x => x.Value.Message), "Max").Excute(sheet);
                errExcel.SaveAs(new FileInfo("err.xlsx"));
            }
        }

        static object? MyCustomValuePolicy(Student student, PropertyInfo property) =>
            property.Name switch
            {
                nameof(Student.Birthday) => student.Birthday?.ToString("yyyy/MM/dd"),
                _ => property.GetValue(student),
            };
    }
}