using System;
using System.Collections.Generic;
using System.Drawing;

namespace simte.Table
{
    public interface ITableColumnBuilder : ITableRowBuilder
    {
        ITableColumnBuilder Column<T>(T? value, Action<ColumnOptionsBuilder> action) where T : struct;
        ITableColumnBuilder Column(Action<ColumnOptionsBuilder> action = null);
        ITableColumnBuilder Column<T>(T value, Action<ColumnOptionsBuilder> action = null);
        ITableColumnBuilder Column(Image value, string name, Action<ColumnOptionsBuilder> action = null);
        ITableColumnBuilder ColumnRange<T>(IEnumerable<T> values, Action<ColumnOptionsBuilder> action = null);
    }
}