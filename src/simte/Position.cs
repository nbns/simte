using System;

namespace simte
{
    public struct Position
    {
        public int Row { get; set; }
        public int Col { get; set; }

        // ctor
        public Position(int row, int col)
        {
            if (row < 1) throw new ArgumentException("row must be more 0", nameof(row));
            if (col < 1) throw new ArgumentException("col must me more 0", nameof(col));
            Row = row;
            Col = col;
        }

        public void Desctructor(out int row, out int col)
        {
            row = Row;
            col = Col;
        }

        public static implicit operator Position ((int, int) tuple2)
            => new Position(tuple2.Item1, tuple2.Item2);
    }
}
