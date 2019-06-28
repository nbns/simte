using simte.Common;

namespace simte.RichText
{
    public interface IRichTextBuilder : IAttachable<IExcelPackage>
    {
        IRichTextBuilder Add(string text /*, style*/);
    }
}