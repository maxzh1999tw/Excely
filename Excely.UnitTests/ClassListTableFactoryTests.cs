using Excely.TableFactories;

namespace Excely.UnitTests
{
	[TestClass]
	public class ClassListTableFactoryTests
	{
		private class TestClass
		{
			public int Id { get; set; }
			public string Name { get; set; }
		}

		[TestMethod]
		public void GetTable_WithDefaultOptions_ShouldIncludeAllPropertiesWithSchema()
		{
			// Arrange
			var factory = new ClassListTableFactory<TestClass>();
			var testData = new List<TestClass>
			{
				new TestClass { Id = 1, Name = "Alice" }
			};

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(2, table.MaxRowCount); // Including schema
			Assert.AreEqual(2, table.MaxColumnCount); // Id and Name
			Assert.AreEqual("Id", table.Data[0][0]);
			Assert.AreEqual("Name", table.Data[0][1]);
		}

		[TestMethod]
		public void GetTable_WithOptionsNoSchema_ShouldExcludeSchema()
		{
			// Arrange
			var options = new ClassListTableFactoryOptions<TestClass> { WithSchema = false };
			var factory = new ClassListTableFactory<TestClass>(options);
			var testData = new List<TestClass>
			{
				new TestClass { Id = 1, Name = "Alice" }
			};

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(1, table.MaxRowCount); // Excluding schema
		}

		[TestMethod]
		public void GetTable_WithCustomPropertyShowPolicy_ShouldFilterProperties()
		{
			// Arrange
			var options = new ClassListTableFactoryOptions<TestClass>
			{
				PropertyShowPolicy = param => param.Property.Name == "Name"
			};
			var factory = new ClassListTableFactory<TestClass>(options);
			var testData = new List<TestClass> { new TestClass { Id = 1, Name = "Alice" } };

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(2, table.MaxRowCount); // Including schema
			Assert.AreEqual(1, table.MaxColumnCount); // Only Name
			Assert.AreEqual("Alice", table.Data[1][0]);
		}

		[TestMethod]
		public void GetTable_WithEmptyList_ShouldReturnOnlySchema()
		{
			// Arrange
			var factory = new ClassListTableFactory<TestClass>();
			var testData = new List<TestClass>();

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual(1, table.MaxRowCount); // Only schema
		}

		[TestMethod]
		public void GetTable_WithNullList_ShouldThrowArgumentNullException()
		{
			// Arrange
			var factory = new ClassListTableFactory<TestClass>();

			// Act & Assert
			Assert.ThrowsException<System.ArgumentNullException>(() => factory.GetTable(null));
		}

		[TestMethod]
		public void GetTable_WithCustomOrderPolicy_ShouldOrderColumnsAccordingly()
		{
			// Arrange
			var options = new ClassListTableFactoryOptions<TestClass>
			{
				PropertyOrderPolicy = param =>
					param.Property.Name == "Name" ? 1 : 2
			};
			var factory = new ClassListTableFactory<TestClass>(options);
			var testData = new List<TestClass> { new TestClass { Id = 1, Name = "Alice" } };

			// Act
			var table = factory.GetTable(testData);

			// Assert
			Assert.AreEqual("Name", table.Data[0][0]);
			Assert.AreEqual("Id", table.Data[0][1]);
		}

	}
}
