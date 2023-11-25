using Excely.TableConverters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excely.UnitTests
{
    [TestClass]
    public class ClassListTableConverterTests
    {
        public class SampleClass
        {
            public int Id { get; set; }
            public string? Name { get; set; }
            public double? Value { get; set; }
        }

        // 基本轉換測試
        [TestMethod]
        public void Convert_ShouldReturnCorrectListOfSampleClass()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Id", "Name", "Value" },
                new List<object?> { 1, "Item 1", 10.0 },
                new List<object?> { 2, "Item 2", 20.0 }
            });
            var converter = new ClassListTableConverter<SampleClass>();

            // Act
            var result = converter.Convert(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("Item 1", result[0].Name);
            Assert.AreEqual(1, result[0].Id);
            Assert.AreEqual(10.0, result[0].Value);
        }

        // 空值轉換測試
        [TestMethod]
        public void Convert_ShouldHandleNullValues()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Id", "Name", "Value" },
                new List<object?> { null, null, null }
            });
            var converter = new ClassListTableConverter<SampleClass>();

            // Act
            var result = converter.Convert(table).ToList();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.IsNull(result[0].Name);
            Assert.AreEqual(0, result[0].Id);
            Assert.IsNull(result[0].Value);
        }

        // 錯誤處理測試
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Convert_ShouldThrowExceptionOnError()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Id", "Name", "Value" },
                new List<object?> { "Invalid", "Item", 10.0 }
            });
            var converter = new ClassListTableConverter<SampleClass>();

            // Act
            converter.Convert(table).ToList();
        }

        [TestMethod]
        public void Convert_WithCustomPropertyNamePolicy_ShouldMatchCorrectly()
        {
            var options = new ClassListTableConverterOptions<SampleClass>
            {
                PropertyNamePolicy = prop => prop.Name.ToUpper()
            };
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "ID", "NAME", "VALUE" },
                new List<object?> { 1, "Item 1", 10.0 }
            });
            var converter = new ClassListTableConverter<SampleClass>(options);

            var result = converter.Convert(table).FirstOrDefault();

            Assert.IsNotNull(result);
            Assert.AreEqual("Item 1", result.Name);
        }

        [TestMethod]
        public void Convert_WithCustomPropertyIndexPolicy_ShouldMatchCorrectly()
        {
            var options = new ClassListTableConverterOptions<SampleClass>
            {
                HasSchema = false,
                PropertyIndexPolicy = (props, prop) => prop.Name == "Name" ? 1 : Array.IndexOf(props, prop)
            };
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { 1, "Item 1", 10.0 }
            });
            var converter = new ClassListTableConverter<SampleClass>(options);

            var result = converter.Convert(table).FirstOrDefault();

            Assert.IsNotNull(result);
            Assert.AreEqual("Item 1", result.Name);
        }

        [TestMethod]
        public void Convert_WithCustomValuePolicy_ShouldTransformValues()
        {
            var options = new ClassListTableConverterOptions<SampleClass>
            {
                CustomValuePolicy = (prop, obj) => prop.Name == "Name" ? obj?.ToString()?.ToUpper() : obj
            };
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Id", "Name", "Value" },
                new List<object?> { 1, "item", 10.0 }
            });
            var converter = new ClassListTableConverter<SampleClass>(options);

            var result = converter.Convert(table).FirstOrDefault();

            Assert.IsNotNull(result);
            Assert.AreEqual("ITEM", result.Name);
        }

        [TestMethod]
        public void Convert_WithCustomErrorHandlingPolicy_ShouldHandleErrors()
        {
            var options = new ClassListTableConverterOptions<SampleClass>
            {
                StopWhenError = false,
            };
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Id", "Name", "Value" },
                new List<object?> { "Invalid", "Item", 10.0 }
            });
            var converter = new ClassListTableConverter<SampleClass>(options);

            var result = converter.Convert(table).ToList();

            Assert.AreEqual(0, result.Count);  // 預期忽略錯誤的行不會被添加
        }

        [TestMethod]
        public void Convert_WithHasSchemaFalse_ShouldIgnoreHeader()
        {
            var options = new ClassListTableConverterOptions<SampleClass>
            {
                HasSchema = false
            };
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { 1, "Item 1", 10.0 },
                new List<object?> { 1, "Item 2", 10.0 },
            });
            var converter = new ClassListTableConverter<SampleClass>(options);

            var result = converter.Convert(table).ToList();

            Assert.AreEqual(2, result.Count);  // 包括標題行和數據行
        }

    }
}
