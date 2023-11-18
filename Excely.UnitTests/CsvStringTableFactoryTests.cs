using Excely.TableFactories;

namespace Excely.UnitTests
{
    [TestClass]
    public class CsvStringTableFactoryTests
    {
        [TestMethod]
        public void GetTable_WithSingleRow_ShouldReturnCorrectTableStructure()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "Id,Name,Email";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(1, table.MaxRowCount);
            Assert.AreEqual(3, table.MaxColCount);
            Assert.AreEqual("Id", table.Data[0][0]);
            Assert.AreEqual("Name", table.Data[0][1]);
            Assert.AreEqual("Email", table.Data[0][2]);
        }

        [TestMethod]
        public void GetTable_WithSingleColumn_ShouldReturnCorrectTableStructure()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "Id\n1\n2";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(3, table.MaxRowCount);
            Assert.AreEqual(1, table.MaxColCount);
            Assert.AreEqual("Id", table.Data[0][0]);
            Assert.AreEqual("1", table.Data[1][0]);
            Assert.AreEqual("2", table.Data[2][0]);
        }

        [TestMethod]
        public void GetTable_SingleColumnWithEmptyColumn_ShouldKeepEmptyColumns()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "Id\n\n2";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(3, table.MaxRowCount);
            Assert.AreEqual(1, table.MaxColCount);
            Assert.AreEqual("Id", table.Data[0][0]);
            Assert.AreEqual("", table.Data[1][0]);
            Assert.AreEqual("2", table.Data[2][0]);
        }

        [TestMethod]
        public void GetTable_WithMultipleRows_ShouldReturnCorrectTableStructure()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "Id,Name,Email\n1,Alice,alice@example.com\n2,Bob,bob@example.com";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(3, table.MaxRowCount); // Including all rows
            Assert.AreEqual(3, table.MaxColCount); // Three columns
            Assert.AreEqual("alice@example.com", table.Data[1][2]);
            Assert.AreEqual("bob@example.com", table.Data[2][2]);
        }

        [TestMethod]
        public void GetTable_WithEmptyString_ShouldReturnEmptyTable()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(0, table.MaxRowCount);
            Assert.AreEqual(0, table.MaxColCount);
        }

        [TestMethod]
        public void GetTable_WithSpecialCharacters_ShouldHandleCommasAndNewLinesCorrectly()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "\"Id, \"\"Identifier\"\"\",\"Name\nFull\",Email";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(1, table.MaxRowCount);
            Assert.AreEqual(3, table.MaxColCount);
            Assert.AreEqual("Id, \"Identifier\"", table.Data[0][0]);
            Assert.AreEqual("Name\nFull", table.Data[0][1]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void GetTable_WithUnclosedQuotes_ShouldThrowException()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "\"Id,Name,Email";

            // Act
            var table = factory.GetTable(csvData);

            // Assert is handled by ExpectedException
        }

        [TestMethod]
        public void GetTable_WithBlankLines_ShouldIgnoreBlankLines()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "\nId,Name,Email\n\n1,John,john@example.com\n";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(2, table.MaxRowCount);
            Assert.AreEqual("John", table.Data[1][1]);
        }

        [TestMethod]
        public void GetTable_ShouldHandleBothNewLineFormatsCorrectly()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            // 使用 \n 和 \r\n 作為換行符
            var csvData = "Id,Name,Email\n1,Alice,alice@example.com\r\n2,Bob,bob@example.com";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(3, table.MaxRowCount); // 包括所有行（標題行和兩個數據行）
            Assert.AreEqual(3, table.MaxColCount); // 三列
            Assert.AreEqual("1", table.Data[1][0]); // 檢查第二行的第一列
            Assert.AreEqual("2", table.Data[2][0]); // 檢查第三行的第一列
        }

        [TestMethod]
        public void GetTable_ShouldHandleBothNewLineFormatsWithQuotesCorrectly()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            // 使用 \n 和 \r\n 作為換行符
            var csvData = "Id,Name,\"Email\"\n1,Alice,\"alice@example.com\"\r\n2,Bob,\"bob@example.com\"";

            // Act
            var table = factory.GetTable(csvData);

            // Assert
            Assert.AreEqual(3, table.MaxRowCount); // 包括所有行（標題行和兩個數據行）
            Assert.AreEqual(3, table.MaxColCount); // 三列
            Assert.AreEqual("1", table.Data[1][0]); // 檢查第二行的第一列
            Assert.AreEqual("2", table.Data[2][0]); // 檢查第三行的第一列
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void GetTable_WithCharactersAfterQuotes_ShouldThrowException()
        {
            // Arrange
            var factory = new CsvStringTableFactory();
            var csvData = "field1,\"field\"2";

            // Act
            var table = factory.GetTable(csvData);

            // Assert is handled by ExpectedException
        }

        // 可以根據需要添加更多的測試案例...
    }
}
