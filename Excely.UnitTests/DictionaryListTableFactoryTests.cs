using Excely.TableFactories;

namespace Excely.UnitTests
{
	[TestClass]
	public class DictionaryListTableFactoryTests
	{
		[TestMethod]
		public void GetTable_WithDefaultOptions_ShouldIncludeAllPropertiesWithSchema()
		{
			// Arrange
			var factory = new DictionaryListTableFactory();
			var testData = new List<Dictionary<string, object?>>
			{
				new Dictionary<string, object?>
				{
					{ "Id", 1 },
					{ "Name", "Alice" }
				}
			};

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(2, table.MaxRowCount); // Including schema
			Assert.AreEqual(2, table.MaxColumnCount); // Id and Name
			Assert.AreEqual("Id", table.Data[0][0]);
			Assert.AreEqual("Name", table.Data[0][1]);
			Assert.AreEqual(1, table.Data[1][0]);
			Assert.AreEqual("Alice", table.Data[1][1]);
		}

		[TestMethod]
		public void GetTable_WithOptionsNoSchema_ShouldExcludeSchema()
		{
			// Arrange
			var options = new DictionaryListTableFactoryOptions { WithSchema = false };
			var factory = new DictionaryListTableFactory(options);
			var testData = new List<Dictionary<string, object?>>
			{
				new Dictionary<string, object?>
				{
					{ "Id", 1 },
					{ "Name", "Alice" }
				}
			};

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(1, table.MaxRowCount); // Excluding schema
			Assert.AreEqual(1, table.Data[0][0]);
			Assert.AreEqual("Alice", table.Data[0][1]);
		}

		[TestMethod]
		public void GetTable_WithCustomPropertyShowPolicy_ShouldFilterProperties()
		{
			// Arrange
			var options = new DictionaryListTableFactoryOptions
			{
				KeyShowPolicy = param => param.Key == "Name"
			};
			var factory = new DictionaryListTableFactory(options);
			var testData = new List<Dictionary<string, object?>>
			{
				new Dictionary<string, object?>
				{
					{ "Id", 1 },
					{ "Name", "Alice" }
				}
			};

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(2, table.MaxRowCount); // Including schema
			Assert.AreEqual(1, table.MaxColumnCount); // Only Name
			Assert.AreEqual("Alice", table.Data[1][0]);
		}

		[TestMethod]
		public void GetTable_WithNullList_ShouldThrowArgumentNullException()
		{
			// Arrange
			var factory = new DictionaryListTableFactory();

			// Act & Assert
			Assert.ThrowsException<System.ArgumentNullException>(() => factory.GetTable(null));
		}

		[TestMethod]
		public void GetTable_WithCustomOrderPolicy_ShouldOrderColumnsAccordingly()
		{
			// Arrange
			var options = new DictionaryListTableFactoryOptions
			{
				CustomValuePolicy = param =>
					param.Key == "Name" ? 1 : 2
			};
			var factory = new DictionaryListTableFactory(options);
			var testData = new List<Dictionary<string, object?>>
			{
				new Dictionary<string, object?>
				{
					{ "Id", 1 },
					{ "Name", "Alice" }
				}
			};

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual("Id", table.Data[0][0]);
			Assert.AreEqual("Name", table.Data[0][1]);
			Assert.AreEqual(2, table.Data[1][0]);
			Assert.AreEqual(1, table.Data[1][1]);
		}

	}
}
