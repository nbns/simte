using OfficeOpenXml;
using simte.Table;
using System;
using System.Collections.Generic;

namespace simte.EPPlus.Table
{
    internal sealed class TableColumnBuilder : BaseColumnBuilder, ITableColumnBuilder
    {
        private readonly RowBuilder _rowBuilder;

        // ctor
        public TableColumnBuilder(RowBuilder rowBuilder, ExcelWorksheet ws) 
            : base(ws)
        {
            _rowBuilder = rowBuilder ?? throw new ArgumentNullException(nameof(rowBuilder));
        }

        public ITableColumnBuilder Row => _rowBuilder.Row;

        public ITableRowBuilder RowSettings(Action<RowOptions> options)
            => _rowBuilder.RowSettings(options);

        public ITableColumnBuilder Column(Action<ColumnOptionsBuilder> action)
            => Column("", action);

        public ITableColumnBuilder Column<T>(T? value, Action<ColumnOptionsBuilder> action) where T : struct
            => value.HasValue ? Column(value.Value, action) : Column("", action);

        public ITableColumnBuilder Column<T>(T value, Action<ColumnOptionsBuilder> action)
        {
            var columnOptionsBuilder = new ColumnOptionsBuilder();
            action?.Invoke(columnOptionsBuilder);
            ColumnOptions options = columnOptionsBuilder;

            var pos = _rowBuilder.GetPositionForCurrentColumn(options.Colspan, options.Rowspan);

            setExcelRange(value, pos, options);
            _rowBuilder.NextColumn(options.Colspan); // next column in the current row
            return this;
        }

        public ITableColumnBuilder ColumnRange<T>(IEnumerable<T> values, Action<ColumnOptionsBuilder> action = null)
        {
            var columnOptionsBuilder = new ColumnOptionsBuilder();
            action?.Invoke(columnOptionsBuilder);
            ColumnOptions options = columnOptionsBuilder;

            foreach (var val in values)
            {
                var pos = _rowBuilder.GetPositionForCurrentColumn(options.Colspan, options.Rowspan);

                setExcelRange(val, pos, options);
                _rowBuilder.NextColumn(options.Colspan); // next column in the current row
            }
            return this;
        }

    }
}

