namespace Excely.UnitTests
{
    [TestClass]
    public class CellLocationTests
    {
        [TestMethod]
        public void Constructor_SetsRowAndColumn()
        {
            // Arrange
            int testRow = 5;
            int testColumn = 10;

            // Act
            CellLocation cellLocation = new CellLocation(testRow, testColumn);

            // Assert
            Assert.AreEqual(testRow, cellLocation.Row);
            Assert.AreEqual(testColumn, cellLocation.Column);
        }

        [TestMethod]
        public void Equals_ReturnsTrueForEqualObjects()
        {
            // Arrange
            CellLocation a = new CellLocation(5, 10);
            CellLocation b = new CellLocation(5, 10);

            // Act & Assert
            Assert.IsTrue(a.Equals(b));
        }

        [TestMethod]
        public void Equals_ReturnsFalseForDifferentObjects()
        {
            // Arrange
            CellLocation a = new CellLocation(5, 10);
            CellLocation b = new CellLocation(6, 11);

            // Act & Assert
            Assert.IsFalse(a.Equals(b));
        }

        [TestMethod]
        public void EqualityOperator_ReturnsTrueForEqualObjects()
        {
            // Arrange
            CellLocation a = new CellLocation(5, 10);
            CellLocation b = new CellLocation(5, 10);

            // Act & Assert
            Assert.IsTrue(a == b);
        }

        [TestMethod]
        public void InequalityOperator_ReturnsFalseForEqualObjects()
        {
            // Arrange
            CellLocation a = new CellLocation(5, 10);
            CellLocation b = new CellLocation(5, 10);

            // Act & Assert
            Assert.IsFalse(a != b);
        }

        [TestMethod]
        public void GetHashCode_ReturnsDifferentValuesForDifferentObjects()
        {
            // Arrange
            CellLocation a = new CellLocation(5, 10);
            CellLocation b = new CellLocation(6, 11);

            // Act & Assert
            Assert.AreNotEqual(a.GetHashCode(), b.GetHashCode());
        }

        [TestMethod]
        public void Equals_ReturnsFalseWhenComparedWithNonCellLocationObject()
        {
            // Arrange
            CellLocation cellLocation = new CellLocation(5, 10);
            object nonCellLocationObject = new { Row = 5, Column = 10 };

            // Act & Assert
            Assert.IsFalse(cellLocation.Equals(nonCellLocationObject));
        }

    }
}
