using NUnit.Framework;
using simte;
using simte.SeedWork;

namespace Tests
{
    [TestFixture]
    public class TablePositionFinder_Test
    {
        [Test]
        [TestCase(1)]
        [TestCase(7)]
        public void Initial_column_should_equal_left_table_position(int column)
        {
            var pos = new Position(row: 1, col: column);
            var tablePositionFinder = new TablePositionFinder(pos);

            var col = tablePositionFinder.GetColumnForNewRow();

            Assert.AreEqual(col, pos.Col);
        }

        [Test]
        [TestCase(1, 1)]
        [TestCase(2, 1)]
        [TestCase(2, 3)]
        public void TablePositionFinder_GetNewPosition(int row, int column)
        {
            var tablePositionFinder = new TablePositionFinder(new Position(row, column));

            var pos = tablePositionFinder.GetNewPosition(column);

            Assert.AreEqual(pos.Value.Row, row);
            Assert.AreEqual(pos.Value.Col, column);
        }

        [Test]
        [TestCase(1, 1)]
        [TestCase(2, 3)]
        public void TablePositionFinder_GetLastPosition(int row, int column)
        {
            var tablePositionFinder = new TablePositionFinder(new Position(row, column));

            var pos = tablePositionFinder.GetLastPosition();

            Assert.AreEqual(pos.Row, row);
            Assert.AreEqual(pos.Col, column);
        }
    }
}