using OfficeOpenXml;

namespace simte.EPPlus
{
    public sealed class Package : IExcelPackage
    {
        private readonly ExcelPackage _package = new ExcelPackage();

        // ctor
        public Package()
        {
        }

        public IWorksheetFactory AddWorksheet(string nameSheet)
            => new WorksheetFactory(this, _package.Workbook.Worksheets.Add(nameSheet));

        public void Save(string filename, string password)
        {
            //package.Workbook.Calculate();
            //package.Workbook.FullCalcOnLoad = true;
            _package.SaveAs(new System.IO.FileInfo(filename), password);
        }

        public void Save(string filename)
        {
            _package.SaveAs(new System.IO.FileInfo(filename));
        }

        public byte[] AsByteArray(string password = "")
            => _package.GetAsByteArray(password);

        public void Dispose()
            => _package?.Dispose();
    }
}