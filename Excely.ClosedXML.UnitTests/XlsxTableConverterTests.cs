using ClosedXML.Excel;
using Excely.ClosedXML.TableConverters;

namespace Excely.ClosedXML.UnitTests
{
    [TestClass]
    public class XlsxTableConverterTests
    {
        /// <summary>
        /// 確保 ClosedXML 在處理字串時不會自動轉型為數字
        /// </summary>
        [TestMethod]
        public void ConvertPaddingStartWith0Data_ShouldReturnExcelWithCurrectData()
        {
            // Arrange
            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("sheet1");

            var converter = new XlsxTableConverter(worksheet);
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Name", "No" },
                new List<object?> { "John", "01" },
            });

            // Act
            converter.ConvertFrom(table);

            // Assert
            Assert.AreEqual("01", worksheet.Cell(2, 2).Value.ToString());
        }
    }
}