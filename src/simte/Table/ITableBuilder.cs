using simte.Common;
using System;
using System.Collections.Generic;

namespace simte.Table
{
    public class TableOptions
    {
        public Position TopLeft { get; set; }
        public bool Border { get; set; }
        public bool Autofit { get; set; }

        public Position? FreezePane { get; set; }
        public bool WrapText { get; set; }
    }

    public interface ITableBuilder : IAttachable<IExcelPackage>
    {
        TableOptions Options { get; }

        ITableBuilder AddRows(Action<ITableRowBuilder> builderAction);
        ITableBuilder AddRows<TModel>(Func<IEnumerable<TModel>> dataFunc, Action<ITableRowBuilder<TModel>> builderAction);
    }
}
