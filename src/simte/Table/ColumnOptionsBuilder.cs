using System;
using System.Drawing;

namespace simte.Table
{
    public class ColumnOptionsBuilder
    {
        private readonly ColumnOptions _options;

        public ColumnOptionsBuilder VerticalText
        {
            get
            {
                _options.VerticalText = true;
                return this;
            }
        }

        // ctor
        public ColumnOptionsBuilder() : this(new ColumnOptions()) { }
        public ColumnOptionsBuilder(ColumnOptions defaultOptions)
            => _options = defaultOptions;

        public ColumnOptionsBuilder TextColor(Color color) => WithOptions(opt => opt.TextColor = color);
        public ColumnOptionsBuilder BackgroundColor(Color color) => WithOptions(opt => opt.BackgroundColor = color);
        public ColumnOptionsBuilder Rowspan(int value) => WithOptions(opt => opt.Rowspan = value);
        public ColumnOptionsBuilder Colspan(int value) => WithOptions(opt => opt.Colspan = value);
        public ColumnOptionsBuilder HorizontalAlignment(HorizontalAlignment alignment) => WithOptions(opt => opt.HorizontalAlignment = alignment);
        public ColumnOptionsBuilder VerticalAlignment(VerticalAlignment alignment) => WithOptions(opt => opt.VerticalAlignment = alignment);
        public ColumnOptionsBuilder NumberFormat(string format) => WithOptions(opt => opt.NumberFormat = format);
        public ColumnOptionsBuilder Width(double width) => WithOptions(opt => opt.Width = width);

        public ColumnOptionsBuilder FontSize(int size) => WithOptions(opt => opt.FontSize = size);
        public ColumnOptionsBuilder FontBold(bool value) => WithOptions(opt => opt.FontBold = value);
        
        public ColumnOptionsBuilder Formula(string formula) => WithOptions(opt => opt.Formula = formula);

        public static implicit operator ColumnOptions(ColumnOptionsBuilder columnOptionsBuilder)
            => columnOptionsBuilder._options;

        protected ColumnOptionsBuilder WithOptions(Action<ColumnOptions> action)
        {
            action(_options);
            return this;
        }
    }

    //public ISheetTextFormatting SumFormula(Position pos, Position beg, Position end)
    //{
    //    using (var cell = _sheet.Cells[pos.Row, pos.Column])
    //    {
    //        cell.Formula = string.Format(
    //            "Sum({0})", new ExcelAddress(beg.Row, beg.Column, end.Row, end.Column).Address
    //            );

    //        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
    //        cell.Style.Font.Bold = true;
    //    }

    //    return this;
    //}

    //public ISheetTextFormatting Formula(Position pos, string formula)
    //{
    //    if (string.IsNullOrEmpty(formula))
    //        throw new ArgumentNullException("formulas null or op is null or empty");

    //    using (var cell = _sheet.Cells[pos.Row, pos.Column])
    //    {
    //        cell.Formula = formula;

    //        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
    //        cell.Style.Font.Bold = true;
    //        return this;
    //    }
    //}


}
