using ClosedXML.Excel;
using Excely.Workflows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Excely.ClosedXML.UnitTests
{
    [TestClass]
    public class XlsxImporterTests
    {
        [TestMethod]
        public void Importer_ShouldReadFileMultipleTimesWithoutLocking()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.AddWorksheet("sheet1");
                ws.Cell(1, 1).Value = "Name";
                ws.Cell(1, 2).Value = "Value";
                ws.Cell(2, 1).Value = "A";
                ws.Cell(2, 2).Value = 1;
                workbook.SaveAs(tempFile);
            }

            var importer = new XlsxImporter();

            // Act & Assert - should not throw because of file locking
            importer.ToDictionaryList(tempFile).ToList();
            importer.ToDictionaryList(tempFile).ToList();
        }
    }
}
