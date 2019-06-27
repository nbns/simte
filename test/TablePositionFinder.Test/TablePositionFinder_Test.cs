using System;
using NUnit.Framework;
using simte;

namespace TablePositionFinder.Test
{
    [TestFixture]
    public class TablePositionFinderTest
    {
        [Test]
        [TestCase(1)]
        [TestCase(7)]
        public void Initial_column_should_equal_left_table_position(int column)
        {
            var pos = new Position(row: 1, col: column);
            var tablePositionFinder = new simte.SeedWork.TablePositionFinder(pos);

            var col = tablePositionFinder.GetColumnForNewRow();

            Assert.AreEqual(col, pos.Col);
        }

        [Test]
        [TestCase(1, 1)]
        [TestCase(2, 1)]
        [TestCase(2, 3)]
        public void TablePositionFinder_GetNewPosition(int row, int column)
        {
            var tablePositionFinder = new simte.SeedWork.TablePositionFinder(new Position(row, column));

            var pos = tablePositionFinder.GetNewPosition(column);

            Assert.AreEqual(pos.Row, row);
            Assert.AreEqual(pos.Col, column);
        }

        [Test]
        public void TablePositionFinder_GetNewPosition_InterceptCellsException()
        {
            var tablePositionFinder = new simte.SeedWork.TablePositionFinder((1, 1));

            // for new row get cell[1,1]
            var col = tablePositionFinder.GetColumnForNewRow();
            tablePositionFinder.GetNewPosition(col);
            
            // cell[1, 2, rowspan=2, colspan = 1] 
            tablePositionFinder.GetNewPosition(col + 1, rowspan: 2);

            // intersect cell with previous and cell[2, 1, rowspan=1, colspan=2]
            Assert.Throws<Exception>(() =>
            {
                col = tablePositionFinder.GetColumnForNewRow(); // next row
                tablePositionFinder.GetNewPosition(col, colspan: 2);
            });
        }

        [Test]
        [TestCase(1, 1)]
        [TestCase(2, 3)]
        public void TablePositionFinder_GetLastPosition(int row, int column)
        {
            var tablePositionFinder = new simte.SeedWork.TablePositionFinder(new Position(row, column));
            var pos = tablePositionFinder.GetLastPosition();
            Assert.AreEqual(pos.Row, row);
            Assert.AreEqual(pos.Col, column);
        }
    }
}