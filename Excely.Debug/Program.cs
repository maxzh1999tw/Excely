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

            var exporter = ExcelyExporter.FromClassList<Student>(
                customValuePolicy: (obj, p) => p.Name switch
                {
                    nameof(Student.Birthday) => obj.Birthday?.ToString("yyyy/MM/dd"),
                    _ => p.GetValue(obj),
                });
            using var excel = exporter.ToExcel(students);

            var reader = new XlsxTableFactory();
            var table = reader.GetTable(excel.Workbook.Worksheets.First());

            var cellErrors = new Dictionary<CellLocation, string>();
            var importResult = new ClassListTableImporter<Student>()
            {
                StopWhenError = false,
                ErrorHandlingPolicy = (cell, student, prop, value, ex) =>
                {
                    cellErrors.Add(cell, ex.ToString());
                    return true;
                }
            }.Import(table);

            if(cellErrors.Any())
            {
                using var errExcel = new ExcelPackage();
                var sheet = errExcel.Workbook.Worksheets.Add("sheet1");
                new XlsxTableWriter(sheet).Write(table);
                new ErrorMarkShader(cellErrors, "Max").Excute(sheet);
                errExcel.SaveAs(new FileInfo("err.xlsx"));
            }
        }
    }
}