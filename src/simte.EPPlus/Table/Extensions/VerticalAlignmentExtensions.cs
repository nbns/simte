using OfficeOpenXml.Style;
using simte.Table;

namespace simte.EPPlus.Table.Extensions
{
    internal static class VerticalAlignmentExtensions
    {
        public static ExcelVerticalAlignment ToEpPlusVerticalAlignment(this VerticalAlignment alignment)
            => (ExcelVerticalAlignment)(int)alignment;
    }
}
