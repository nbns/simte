using simte.Common;
using System;
using System.Collections.Generic;

namespace simte.Table
{
    public interface ITableBuilder : IAttachable<IExcelPackage>
    {
        TableOptions Options { get; }

        ITableBuilder AddRows(Action<ITableRowBuilder> builderAction);
        ITableBuilder AddRows<TModel>(Func<IEnumerable<TModel>> dataFunc, Action<ITableRowBuilder<TModel>> builderAction);
    }
}
