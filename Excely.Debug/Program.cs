using Excely.Debug.Models;
using Excely.EPPlus.LGPL.Exporters;
using Excely.EPPlus.LGPL.Shaders.Xlsx;
using Excely.EPPlus.LGPL.TableFactories;
using Excely.Shaders;
using Excely.TableWriters;

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

            var exporter = new ClassListExporter<Student>
            {
                CustomValuePolicy = (student, property) => property.Name switch
                {
                    nameof(Student.Birthday) => student.Birthday?.ToString("yyyy/MM/dd"),
                    _ => property.GetValue(student),
                },
                Shaders = new IShader[]
                {
                    new CellFittingShader()
                }
            };

            using var excel = exporter.ToExcel(students);

            var reader = new XlsxTableFactory();
            var table = reader.GetTable(excel.Workbook.Worksheets.First());
            var list = new ClassListTableImporter<Student>()
            {
                StopWhenError = false
            }.Import(table);
        }
    }
}