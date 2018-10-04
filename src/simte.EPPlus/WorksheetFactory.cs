using OfficeOpenXml;
using simte.EPPlus.Table;
using simte.Table;
using System;

namespace simte.EPPlus
{
    public class WorksheetFactory : IWorksheetFactory
    {
        private readonly IExcelPackage _package;
        protected internal readonly ExcelWorksheet ws;

        // ctor
        public WorksheetFactory(IExcelPackage package, ExcelWorksheet excelWorksheet)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            ws = excelWorksheet ?? throw new ArgumentNullException(nameof(excelWorksheet));
        }

        public ITableBuilder Table(TableOptions options)
            => new TableBuilder(this, options);

        public IExcelPackage Attach()
            => _package;
    }
}
