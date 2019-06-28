using OfficeOpenXml;
using simte.EPPlus.Table;
using simte.Table;
using System;
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
                range.Style.VerticalAlignment = options.VerticalAlignment.ToEPPlusVerticalAlignment();

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
        
        public IExcelPackage Attach()
            => _package;
    }
}
