using simte.Common;

namespace simte.Table
{
    public interface ITableRowBuilder<TSource> : IApplyable<ITableRowBuilder<TSource>>
    {
        ITableColumnBuilder<TSource> Row { get; }
    }
}
