using Excely.TableConverters;
using Excely.UnitTests.Models;

namespace Excely.UnitTests
{
    [TestClass]
    public class ClassListTableConverterTests
    {
        /// <summary>
        /// 基礎轉換測試，最簡單的資料，使用預設邏輯。
        /// </summary>
        [TestMethod]
        public void ConvertFromData_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
            });
            var converter = new ClassListTableConverter<SimpleClass>();

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 基礎轉換測試，有可能為 Null 的資料，使用預設邏輯。
        /// </summary>
        [TestMethod]
        public void ConvertFromNullableData_ShouldReturnCorrectListOfSimpleNullableClass()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { null, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 1, null, DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
                new List<object?> { 2, "Text2", null, SampleEnum.Enum3 },
                new List<object?> { 3, "Text3", DateTime.Parse("2023/1/3"), null },
            });
            var converter = new ClassListTableConverter<SimpleNullableClass>();

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(4, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 基礎轉換測試，包含無須轉換的欄位，使用預設邏輯。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithUselessField_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue", "UselessField" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1, "Useless" },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2, "Useless" },
            });
            var converter = new ClassListTableConverter<SimpleClass>();

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 自訂表頭測試。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithCustomSchema_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "數字", "文字", "日期", "Enum" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
            });
            var options = new ClassListTableConverterOptions<SimpleClass>
            {
                HasSchema = true,
                PropertyNamePolicy = p => p.Name switch
                {
                    nameof(SimpleClass.IntValue) => "數字",
                    nameof(SimpleClass.StrigValue) => "文字",
                    nameof(SimpleClass.DateTimeValue) => "日期",
                    nameof(SimpleClass.EnumValue) => "Enum",
                    _ => throw new NotImplementedException()
                }
            };
            var converter = new ClassListTableConverter<SimpleClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 不含表頭測試，使用預設邏輯。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithoutSchema_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
            });
            var options = new ClassListTableConverterOptions<SimpleNullableClass>
            {
                HasSchema = false
            };
            var converter = new ClassListTableConverter<SimpleNullableClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 不含表頭且順序調換測試。
        /// </summary>
        [TestMethod]
        public void ConvertFromColumnSwappedDataWithoutSchema_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { SampleEnum.Enum1, DateTime.Parse("2023/1/1"), "Text1", 1 },
                new List<object?> { SampleEnum.Enum2, DateTime.Parse("2023/1/2"), "Text2", 2 },
            });
            var options = new ClassListTableConverterOptions<SimpleNullableClass>
            {
                HasSchema = false,
                PropertyIndexPolicy = (propties, p) => p.Name switch
                {
                    nameof(SimpleClass.IntValue) => 3,
                    nameof(SimpleClass.StrigValue) => 2,
                    nameof(SimpleClass.DateTimeValue) => 1,
                    nameof(SimpleClass.EnumValue) => 0,
                    _ => throw new NotImplementedException()
                }
            };
            var converter = new ClassListTableConverter<SimpleNullableClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i][3], result[i].IntValue);
                Assert.AreEqual(table.Data[i][2], result[i].StrigValue);
                Assert.AreEqual(table.Data[i][1], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i][0], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 不含表頭且包含無須轉換的欄位測試。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithMissingFieldWithoutSchema_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1},
                new List<object?> { "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2},
            });
            var options = new ClassListTableConverterOptions<SimpleNullableClass>
            {
                HasSchema = false,
                PropertyIndexPolicy = (props, p) => p.Name switch
                {
                    nameof(SimpleNullableClass.IntValue) => null,
                    _ => Array.IndexOf(props, p) - 1
                }
            };
            var converter = new ClassListTableConverter<SimpleNullableClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(null, result[i].IntValue);
                Assert.AreEqual(table.Data[i][0], result[i].StrigValue);
                Assert.AreEqual(table.Data[i][1], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i][2], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 自訂寫入值測試。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithPropertyValueSettingPolicy_ShouldReturnCorrectList()
        {
#pragma warning disable CS8605 // Unboxing 可能 null 值。
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
            });
            var options = new ClassListTableConverterOptions<SimpleClass>
            {
                PropertyValueSettingPolicy = (p, value) => p.Name switch
                {
                    nameof(SimpleClass.IntValue) => (int)value * 2,
                    _ => value,
                }
            };
            var converter = new ClassListTableConverter<SimpleClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual((int)table.Data[i + 1][0] * 2, result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
#pragma warning restore CS8605 // Unboxing 可能 null 值。
        }

        /// <summary>
        /// 錯誤資料測試，錯誤時擲回錯誤。
        /// </summary>
        [TestMethod]
        public void ConvertFromErrorData_ShouldThrowExceptions()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), true },
            });
            var options = new ClassListTableConverterOptions<SimpleClass>
            {
                StopWhenError = false,
            };
            var converter = new ClassListTableConverter<SimpleClass>();

            // Assert
            Assert.ThrowsException<ArgumentException>(() => converter.ConvertFrom(table));
        }

        /// <summary>
        /// 錯誤資料測試，使用客製化錯誤處理忽略錯誤資料並記錄。
        /// </summary>
        [TestMethod]
        public void ConvertFromErrorData_ShouldRecordErrorAndReturnListWithoutErrorRecords()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), true },
            });
            var errors = new Dictionary<CellLocation, Exception>();
            var options = new ClassListTableConverterOptions<SimpleClass>
            {
                StopWhenError = false,
                ErrorHandlingPolicy = (location, obj, p, value, err) =>
                {
                    errors.Add(location, err);
                    return false;
                }
            };
            var converter = new ClassListTableConverter<SimpleClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(1, errors.Count);
            Assert.AreEqual(new CellLocation(2, 3), errors.First().Key);
            Assert.AreEqual(1, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 錯誤資料測試，錯誤時不停止。
        /// </summary>
        [TestMethod]
        public void ConvertFromErrorData_ShouldReturnCorrectListWithoutErrorRecords()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), true },
            });
            var options = new ClassListTableConverterOptions<SimpleClass>
            {
                StopWhenError = false,
            };
            var converter = new ClassListTableConverter<SimpleClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(1, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
                Assert.AreEqual(table.Data[i + 1][2], result[i].DateTimeValue);
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 錯誤資料測試，使用客製化錯誤處理修復錯誤資料。
        /// </summary>
        [TestMethod]
        public void ConvertFromErrorData_ShouldReturnListWithFixedData()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), true },
            });
            var errors = new Dictionary<CellLocation, Exception>();
            var options = new ClassListTableConverterOptions<SimpleNullableClass>
            {
                StopWhenError = false,
                ErrorHandlingPolicy = (location, obj, p, value, err) =>
                {
                    p.SetValue(obj, null);
                    return true;
                }
            };
            var converter = new ClassListTableConverter<SimpleNullableClass>(options);

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual(table.Data[1][0], result[0].IntValue);
            Assert.AreEqual(table.Data[1][1], result[0].StrigValue);
            Assert.AreEqual(table.Data[1][2], result[0].DateTimeValue);
            Assert.AreEqual(table.Data[1][3], result[0].EnumValue);
            Assert.AreEqual(table.Data[2][0], result[1].IntValue);
            Assert.AreEqual(table.Data[2][1], result[1].StrigValue);
            Assert.AreEqual(table.Data[2][2], result[1].DateTimeValue);
            Assert.AreEqual(null, result[1].EnumValue);
        }

        /// <summary>
        /// 自動轉型測試。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithWrongType_ShouldReturnCorrectList()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", "2023/1/1", SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", "2023/1/2", SampleEnum.Enum2 },
            });
            var converter = new ClassListTableConverter<SimpleClass>();

            // Act
            var result = converter.ConvertFrom(table).ToList();

            // Assert
            Assert.AreEqual(2, result.Count);
            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(table.Data[i + 1][0], result[i].IntValue);
                Assert.AreEqual(table.Data[i + 1][1], result[i].StrigValue);
#pragma warning disable CS8600 // 正在將 Null 常值或可能的 Null 值轉換為不可為 Null 的型別。
#pragma warning disable CS8604 // 可能有 Null 參考引數。
                Assert.AreEqual(DateTime.Parse((string)table.Data[i + 1][2]), result[i].DateTimeValue);
#pragma warning restore CS8604 // 可能有 Null 參考引數。
#pragma warning restore CS8600 // 正在將 Null 常值或可能的 Null 值轉換為不可為 Null 的型別。
                Assert.AreEqual(table.Data[i + 1][3], result[i].EnumValue);
            }
        }

        /// <summary>
        /// 關閉自動轉型測試。
        /// </summary>
        [TestMethod]
        public void ConvertFromDataWithWrongType_ShouldThrowException()
        {
            // Arrange
            var table = new ExcelyTable(new List<IList<object?>>
            {
                new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "EnumValue" },
                new List<object?> { 1, "Text1", "2023/1/1", SampleEnum.Enum1 },
                new List<object?> { 2, "Text2", "2023/1/2", SampleEnum.Enum2 },
            });
            var options = new ClassListTableConverterOptions<SimpleClass>
            {
                EnableAutoTypeConversion = false
            };
            var converter = new ClassListTableConverter<SimpleClass>(options);

            // Assert
            Assert.ThrowsException<ArgumentException>(() => converter.ConvertFrom(table));
        }
    }
}
