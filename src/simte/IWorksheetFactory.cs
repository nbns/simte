using simte.Common;
using simte.Table;

namespace simte
{
    public interface IWorksheetFactory : IAttachable<IExcelPackage>
    {
        ITableBuilder Table(TableOptions options);
    }
}
