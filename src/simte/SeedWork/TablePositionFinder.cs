using System;
using System.Linq;

namespace simte.SeedWork
{
    public class TablePositionFinder
    {
        private int[] _indexOfRows = new int[100];
        private Position _topLeft;
        private int _lastIndex;

        // ctor
        public TablePositionFinder(Position topLeft)
        {
            _topLeft = topLeft;
            _lastIndex = topLeft.Col - 1;
        }

        public int GetColumnForNewRow()
        {
            var rows = _indexOfRows.Where(x => x > 0);
            return rows.Any()
                ? Array.IndexOf(_indexOfRows, rows.Min()) + 1
                : _topLeft.Col;
        }

        public Position GetNewPosition(int column, int colspan = 1, int rowspan = 1)
        {
            if (column < 1) throw new ArgumentNullException("Column is less than 1");

            column = findColumnForInsert(column);
            // allocate
            var maxLength = column + colspan - 1;
            if (maxLength > _indexOfRows.Length) _indexOfRows = reallocateBuffer(_indexOfRows, maxLength);

            // lastindex
            if (_lastIndex < (column - 1)) _lastIndex = column - 1;

            // fill 
            var row = _indexOfRows[column - 1]; // array [0..]
            for (int index = column - 1; index < maxLength; ++index)
            {
                if (row != _indexOfRows[index])
                {
                    // not appropriate row, column
                    throw new Exception($"The intersection of the cells in row={row}, column={column} (rowspan={rowspan}, colspan={colspan}");
                }

                _indexOfRows[index] += rowspan;
            }

            return new Position(_topLeft.Row + row, column);
        }

        public Position GetLastPosition()
            => new Position(_topLeft.Row + _indexOfRows[_lastIndex], _lastIndex + 1);

        private int[] reallocateBuffer(int[] indexOfRows, int maxLength)
        {
            var newIndexOfRows = new int[_indexOfRows.Length * 2 + maxLength];
            // copy 
            Buffer.BlockCopy(_indexOfRows, 0, newIndexOfRows, 0, _indexOfRows.Length);
            return newIndexOfRows;
        }

        private int findColumnForInsert(int column)
        {
            var min = _indexOfRows.Skip(column - 1).Where(x => x > 0);
            if (min.Any())
            {
                var idx = Array.IndexOf(_indexOfRows, min.Min());
                column = idx + 1;
            }

            return column;
        }
    }
}