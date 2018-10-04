using OfficeOpenXml.Style;
using simte.Table;

namespace simte.EPPlus.Table.Extensions
{
    internal static class HorizontalAlignmentExtensions
    {
        public static ExcelHorizontalAlignment ToEPPlusHorizontalAligment(this HorizontalAlignment alignment)
            => (ExcelHorizontalAlignment)(int)alignment;
    }
}
