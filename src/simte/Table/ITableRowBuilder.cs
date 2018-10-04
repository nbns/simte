using System;

namespace simte.Table
{
    public interface ITableRowBuilder 
    {
        ITableRowBuilder RowSettings(Action<RowOptions> options);
        ITableColumnBuilder Row { get; }
    }
}