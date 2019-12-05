using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;

namespace simte.EPPlus.Extensions
{
    internal static class ExcelRangeExtensions
    {
        /// <summary>
        /// There are 72 point in inch. Pixels = points * (1/72.0) * DPI
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int GetHeightInPixels(this ExcelRange range)
        {
            var sum = 0.0;
            using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                for (var currentRow = range.Start.Row; currentRow < range.End.Row; ++currentRow)
                {
                    var dpiY = graphics.DpiY;
                    sum += (range.Worksheet.Row(currentRow).Height * (1 / 72.0) * dpiY);
                }
            }

            return (int) sum;
        }

        /// <summary>
        /// https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int GetWidthInPixels(this ExcelRange range)
        {
            var columnWidth = Enumerable.Range(range.Start.Column, range.End.Column - range.Start.Column).Sum(x => range.Worksheet.Column(x).Width);
            // https://stackoverflow.com/questions/50352133/epplus-position-image-in-a-cell
            var font = new Font(range.Style.Font.Name, range.Style.Font.Size, FontStyle.Regular);
            var pxBaseline = Math.Round(measureString("1234567890", font) / 10);

            return (int) (columnWidth * pxBaseline);
        }

        private static float measureString(string s, Font font)
        {
            using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                return graphics.MeasureString(s, font, int.MaxValue, StringFormat.GenericTypographic).Width;
            }
        }
    }
}