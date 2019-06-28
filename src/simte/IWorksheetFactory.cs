using System;
using simte.Common;
using simte.RichText;
using simte.Table;

namespace simte
{
    public interface IWorksheetFactory : IAttachable<IExcelPackage>
    {
        ITableBuilder Table(TableOptions options);
        IRichTextBuilder RichText(Position pos);
        IWorksheetFactory Text(string text, Position pos, Action<ColumnOptionsBuilder> action = null, double? rowHeight = null);
        
        /// <summary>
        /// Indicates where the last action was
        /// </summary>
        int LastRow { get; }
    }
}
