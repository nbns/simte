using OfficeOpenXml;
using simte.EPPlus.Table;
using simte.Table;
using System;
using System.Drawing;
using System.IO;
using simte.EPPlus.Extensions;
using simte.EPPlus.Table.Extensions;
using simte.RichText;

namespace simte.EPPlus
{
    public class WorksheetFactory : IWorksheetFactory
    {
        private readonly IExcelPackage _package;
        protected internal readonly ExcelWorksheet ws;

        public int LastRow => ws.Dimension.End.Row;

        // ctor
        public WorksheetFactory(IExcelPackage package, ExcelWorksheet excelWorksheet)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            ws = excelWorksheet ?? throw new ArgumentNullException(nameof(excelWorksheet));
        }

        public ITableBuilder Table(TableOptions options)
            => new TableBuilder(this, options);

        public IWorksheetFactory Text(string text, Position pos, Action<ColumnOptionsBuilder> action = null,
            double? rowHeight = null)
        {
            var columnOptionsBuilder = new ColumnOptionsBuilder();
            action?.Invoke(columnOptionsBuilder);
            ColumnOptions options = columnOptionsBuilder;
            ;

            using (var range = ws.Cells[pos.Row, pos.Col, pos.Row + options.Rowspan - 1, pos.Col + options.Colspan - 1])
            {
                range.Merge = options.Colspan > 1 || options.Rowspan > 1;
                range.Value = text;

                if (options.Width.HasValue)
                    ws.Column(pos.Col).Width = options.Width.Value;

                if (rowHeight.HasValue)
                    ws.Row(pos.Row).Height = rowHeight.Value;

                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = options.HorizontalAlignment.ToEPPlusHorizontalAligment();
                range.Style.VerticalAlignment = options.VerticalAlignment.ToEpPlusVerticalAlignment();

                range.Style.Font.Bold = options.FontBold;
                range.Style.Font.Size = options.FontSize ?? 11;
                // border
                //range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                if (options.VerticalText)
                {
                    range.Style.TextRotation = 90;
                }

                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Font.Color.SetColor(options.TextColor);
                range.Style.Fill.BackgroundColor.SetColor(options.BackgroundColor);
            }

            return this;
        }

        public IRichTextBuilder RichText(Position pos)
        {
            throw new NotImplementedException();
        }

        public Position Picture(string name, Position pos, Stream stream) =>
            Picture(name, pos, Image.FromStream(stream));

        public Position Picture(string name, Position pos, Image image)
        {
            var pic = ws.Drawings.AddPicture(name, image);
            pic.From.Row = pos.Row - 1;
            pic.From.Column = pos.Col - 1;

            return new Position(pic.To.Row, pic.To.Column);
        }

        public Position Picture(string name, Position from, Position to, Image image)
        {
            var pic = ws.Drawings.AddPicture(name, image);

            pic.From.Row = from.Row - 1;
            pic.From.Column = from.Col - 1;
            pic.To.Row = to.Row - 1;
            pic.To.Column = to.Col- 1;

            var size = GetPixelsSizeOfCellRange(from, to);
            pic.SetSize(size.Width, size.Height);

            return new Position(pic.To.Row, pic.To.Column);
        }

        public Size GetPixelsSizeOfCell(Position pos)
        {
            var cell = ws.Cells[pos.Row, pos.Col, pos.Row + 1, pos.Col + 1];
            return new Size(cell.GetWidthInPixels(), cell.GetHeightInPixels());
        }

        public Size GetPixelsSizeOfCellRange(Position from, Position to)
        {
            var cell = ws.Cells[from.Row, from.Col, to.Row, to.Col];
            return new Size(cell.GetWidthInPixels(), cell.GetHeightInPixels());
        }

        public IExcelPackage Attach()
            => _package;
    }
}