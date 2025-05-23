using Excely.TableConverters;
using Excely.UnitTests.Models;

namespace Excely.UnitTests
{
	[TestClass]
	public class DictionaryListTableConverterTests
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
			var converter = new DictionaryListTableConverter();

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(2, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual(table.Data[i + 1][0], result[i]["IntValue"]);
				Assert.AreEqual(table.Data[i + 1][1], result[i]["StrigValue"]);
				Assert.AreEqual(table.Data[i + 1][2], result[i]["DateTimeValue"]);
				Assert.AreEqual(table.Data[i + 1][3], result[i]["EnumValue"]);
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
			var converter = new DictionaryListTableConverter();

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(4, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual(table.Data[i + 1][0], result[i]["IntValue"]);
				Assert.AreEqual(table.Data[i + 1][1], result[i]["StrigValue"]);
				Assert.AreEqual(table.Data[i + 1][2], result[i]["DateTimeValue"]);
				Assert.AreEqual(table.Data[i + 1][3], result[i]["EnumValue"]);
			}
		}

		/// <summary>
		/// 有 null 表頭不匯入。
		/// </summary>
		[TestMethod]
		public void ConvertFromDataWithNullSchema_ShouldReturnCorrectList()
		{
			// Arrange
			var table = new ExcelyTable(new List<IList<object?>>
			{
				new List<object?> { "IntValue", "StrigValue", "DateTimeValue", null },
				new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
				new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
			});
			var converter = new DictionaryListTableConverter();

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(2, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual(table.Data[i + 1][0], result[i]["IntValue"]);
				Assert.AreEqual(table.Data[i + 1][1], result[i]["StrigValue"]);
				Assert.AreEqual(table.Data[i + 1][2], result[i]["DateTimeValue"]);
				Assert.AreEqual(table.Data[i + 1][3], result[i]["3"]);
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
				new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "Enum" },
				new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
				new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
			});
			var options = new DictionaryListTableConverterOptions
			{
				HasSchema = true,
				CustomKeyNamePolicy = param => param.FieldName + param.FieldIndex,
			};
			var converter = new DictionaryListTableConverter(options);

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(2, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual(table.Data[i + 1][0], result[i]["IntValue0"]);
				Assert.AreEqual(table.Data[i + 1][1], result[i]["StrigValue1"]);
				Assert.AreEqual(table.Data[i + 1][2], result[i]["DateTimeValue2"]);
				Assert.AreEqual(table.Data[i + 1][3], result[i]["Enum3"]);
			}
		}

		/// <summary>
		/// 自訂表頭測試，忽略某些欄位。
		/// </summary>
		[TestMethod]
		public void ConvertFromDataWithCustomSchema_IgnoreSomeField_ShouldReturnCorrectList()
		{
			// Arrange
			var table = new ExcelyTable(new List<IList<object?>>
			{
				new List<object?> { "IntValue", "StrigValue", "DateTimeValue", "Enum" },
				new List<object?> { 1, "Text1", DateTime.Parse("2023/1/1"), SampleEnum.Enum1 },
				new List<object?> { 2, "Text2", DateTime.Parse("2023/1/2"), SampleEnum.Enum2 },
			});
			var options = new DictionaryListTableConverterOptions
			{
				HasSchema = true,
				CustomKeyNamePolicy = param => param.FieldName == "Enum" ? null : param.FieldName,
			};
			var converter = new DictionaryListTableConverter(options);

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(2, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual(table.Data[i + 1][0], result[i]["IntValue"]);
				Assert.AreEqual(table.Data[i + 1][1], result[i]["StrigValue"]);
				Assert.AreEqual(table.Data[i + 1][2], result[i]["DateTimeValue"]);
				Assert.IsFalse(result[i].ContainsKey("Enum"));
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
			var options = new DictionaryListTableConverterOptions
			{
				HasSchema = false
			};
			var converter = new DictionaryListTableConverter(options);

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(2, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual(table.Data[i][0], result[i]["0"]);
				Assert.AreEqual(table.Data[i][1], result[i]["1"]);
				Assert.AreEqual(table.Data[i][2], result[i]["2"]);
				Assert.AreEqual(table.Data[i][3], result[i]["3"]);
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
			var options = new DictionaryListTableConverterOptions
			{
				CustomValuePolicy = param => param.Key switch
				{
					"IntValue" => (int)param.OriginalValue * 2,
					_ => param.OriginalValue,
				}
			};
			var converter = new DictionaryListTableConverter(options);

			// Act
			var result = converter.ConvertFrom(table).ToList();

			// Assert
			Assert.AreEqual(2, result.Count);
			for (int i = 0; i < result.Count; i++)
			{
				Assert.AreEqual((int)table.Data[i + 1][0] * 2, result[i]["IntValue"]);
				Assert.AreEqual(table.Data[i + 1][1], result[i]["StrigValue"]);
				Assert.AreEqual(table.Data[i + 1][2], result[i]["DateTimeValue"]);
				Assert.AreEqual(table.Data[i + 1][3], result[i]["EnumValue"]);
			}
#pragma warning restore CS8605 // Unboxing 可能 null 值。
		}
	}
}
