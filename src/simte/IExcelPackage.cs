using System;

namespace simte
{
    public interface IExcelPackage : IDisposable
    {
        IWorksheetFactory AddWorksheet(string nameSheet);
        
        void Save(string filename, string password);
        void Save(string filename);
        
        byte[] AsByteArray(string password = "");
    }
}
