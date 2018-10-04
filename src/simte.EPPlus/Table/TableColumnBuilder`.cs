using System;
using simte.Table;

namespace simte.EPPlus.Table
{
    internal sealed class TableColumnBuilder<TSource> :  ITableColumnBuilder<TSource>
    {
        private readonly RowBuilder<TSource> _rowBuilder;

        // ctor
        public TableColumnBuilder(RowBuilder<TSource> rowBuilder)
        {
            _rowBuilder = rowBuilder ?? throw new ArgumentNullException(nameof(rowBuilder));
        }

        public ITableColumnBuilder<TSource> Row => _rowBuilder.Row;

        public ITableRowBuilder<TSource> Apply()
            => _rowBuilder.Apply();

        public ITableColumnBuilder<TSource> Select(Func<TSource, object> selector, Action<ColumnOptionsBuilder<TSource>> action = null)
        {
            _rowBuilder.AddColumn(selector, action);
            return this;
        }

        public ITableColumnBuilder<TSource> Skip(int value)
        {
            _rowBuilder.AddColumn(null, options => options.Colspan(value));
            return this;
        }
    }
}
