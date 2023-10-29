using Excely.Debug.Models;
using Excely.Exporters;
using Excely.Shaders;
using Excely.Shaders.Xlsx;

namespace Excely.Debug
{
    internal class Program
    {
        static void Main(string[] args)
        {
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
        }
    }
}