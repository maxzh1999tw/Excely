using Excely.Debug.Models;
using Excely.Workflows;

namespace Excely.Debug
{
    internal class Program
    {
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
            var importResult = importer.Import(excel.Workbook.Worksheets.First());
        }
    }
}