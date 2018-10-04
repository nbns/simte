using OfficeOpenXml;
using OfficeOpenXml.Style;
using simte.EPPlus.Table.Extensions;
using simte.Table;
using System;
using System.Drawing;

namespace simte.EPPlus.Table
{
    public abstract class BaseColumnBuilder
    {
        protected readonly ExcelWorksheet _ws;

        // ctor
        public BaseColumnBuilder(ExcelWorksheet ws)
        {
            _ws = ws ?? throw new ArgumentNullException(nameof(ws));
        }

        protected void setRowHeight(int height, Position pos)
            => _ws.Row(pos.Row).Height = height;

        protected void setExcelRange<T>(T value, Position pos, ColumnOptions options)
        {
            using (var range = _ws.Cells[pos.Row, pos.Col, pos.Row + options.Rowspan - 1, pos.Col + options.Colspan - 1])
            {
                range.Merge = options.Colspan > 1 || options.Rowspan > 1;
                if (!string.IsNullOrEmpty(options.Formula)) range.Formula = options.Formula; else range.Value = value;

                if (options.Width.HasValue)
                    _ws.Column(pos.Col).Width = options.Width.Value;

                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = options.HorizontalAligment.ToEPPlusHorizontalAligment();
                range.Style.VerticalAlignment = options.VerticalAligment.ToEPPlusVerticalAlignment();

                // border
                range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                if (options.VerticalText)
                {
                    range.Style.TextRotation = 90;
                }

                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Font.Color.SetColor(options.TextColor);
                range.Style.Fill.BackgroundColor.SetColor(options.BackgroundColor);

                //////MeasureTextHeight(value.ToString(), range.Style.Font, options.Colspan);
                //////_ws.Column(pos.Col).Width = Math.Min(size.width * 72.0 / 96.0, 409);
                //var size = measureText(value?.ToString() , range.Style.Font, options.Colspan);// * 72.0 / 96.0;
                //_ws.Row(pos.Row).Height = size.height; //Math.Min(size.height * 72.0 / 96.0, 409);
            }
        }


        private (double height, double width) measureText(string text, ExcelFont excelFont, int width)
        {
            if (string.IsNullOrEmpty(text)) return (0, 0);

            Font drawFont = null;
            SolidBrush drawBrush = null;
            Graphics drawGraphics = null;
            Bitmap textBitmap = null;
            try
            {
                // start with empty bitmap, get it's graphic's object
                // and choose a font
                textBitmap = new Bitmap(1, 1);
                drawGraphics = Graphics.FromImage(textBitmap);
                drawFont = new Font(excelFont.Name, excelFont.Size);


                width = (int)(width * 7.5) * text.Length;
                // see how big the text will be
                var size = drawGraphics.MeasureString(text, drawFont, width);


                //// recreate the bitmap and graphic object with the new size
                //textBitmap = new Bitmap(textBitmap, Width, Height);
                //drawGraphics = Graphics.FromImage(textBitmap);


                //// get the drawing brush and where we're going to draw
                //drawBrush = new SolidBrush(Color.Black);
                //PointF DrawPoint = new PointF(0, 0);


                // draw
                //// clear the graphic white and draw the string
                //DrawGraphics.Clear(Color.White);
                //DrawGraphics.DrawString(TheText, DrawFont, DrawBrush, DrawPoint);


                //TextBitmap.Save("bmp2\\" + Guid.NewGuid().ToString() + ".bmp");
                return (size.Height, size.Width);
            }
            finally
            {
                // don't dispose the bitmap, the caller needs it.
                drawFont?.Dispose();
                drawBrush?.Dispose();
                drawGraphics?.Dispose();
            }

        }
    }
}
