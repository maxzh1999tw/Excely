using Excely.TableConverters;
using System.Text;

namespace Excely.UnitTests
{
    [TestClass]
    public class CsvStringTableConverterTests
    {
        [TestMethod]
        public void ConvertFrom_ConvertToStringWithSimpleData_ShouldReturnCorrectCsvString()
        {
            // Arrange
            var converter = new CsvStringTableConverter<string>();
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Name", "Surname", "Age" },
                new List<object?> { "John", "Doe", 30 },
                new List<object?> { "Jane", "Doe", null }
            });

            // Act
            var csvString = converter.ConvertFrom(table);

            // Assert
            Assert.AreEqual(
                $"Name,Surname,Age{Environment.NewLine}John,Doe,30{Environment.NewLine}Jane,Doe,",
                csvString);
        }

        [TestMethod]
        public void ConvertFrom_ConvertToMemoryStreamWithSimpleData_ShouldReturnCorrectMemoryStream()
        {
            // Arrange
            var converter = new CsvStringTableConverter<MemoryStream>();
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Name", "Surname", "Age" },
                new List<object?> { "John", "Doe", 30 },
                new List<object?> { "Jane", "Doe", null }
            });

            // Act
            var csvStream = converter.ConvertFrom(table);

            // Assert
            Assert.AreEqual(
                $"Name,Surname,Age{Environment.NewLine}John,Doe,30{Environment.NewLine}Jane,Doe,",
                Encoding.Default.GetString(csvStream.ToArray()));
        }

        [TestMethod]
        public void ConvertFrom_ConvertToStreamWithSimpleData_ShouldReturnCorrectStream()
        {
            // Arrange
            var converter = new CsvStringTableConverter<Stream>();
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Name", "Surname", "Age" },
                new List<object?> { "John", "Doe", 30 },
                new List<object?> { "Jane", "Doe", null }
            });

            // Act
            var csvStream = converter.ConvertFrom(table);
            using StreamReader reader = new(csvStream, Encoding.Default);
            var csvString = reader.ReadToEnd();

            // Assert
            Assert.AreEqual(
                $"Name,Surname,Age{Environment.NewLine}John,Doe,30{Environment.NewLine}Jane,Doe,",
                csvString);
        }

        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void ConvertFrom_ConvertToNotSupportedType_ShouldThrowException()
        {
            // Arrange
            var converter = new CsvStringTableConverter<StringBuilder>();
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Name", "Surname", "Age" },
                new List<object?> { "John", "Doe", 30 },
                new List<object?> { "Jane", "Doe", null }
            });

            // Act
            var csvStream = converter.ConvertFrom(table);

            // Assert is handled by ExpectedException
        }
    }
}
