using System;

namespace simte.Table
{
    public interface ITableColumnBuilder<TSource> : ITableRowBuilder<TSource>
    {
        ITableColumnBuilder<TSource> Select(Func<TSource, object> selector, Action<ColumnOptionsBuilder<TSource>> action = null);
        ITableColumnBuilder<TSource> Skip(int value);
    }
}
