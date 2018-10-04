using System.Drawing;

namespace simte.Table
{
    public class ColumnOptions
    {
        public Color TextColor { get; set; } = Color.Black;
        public Color BackgroundColor { get; set; } = Color.White;
        public int Rowspan { get; set; } = 1;
        public int Colspan { get; set; } = 1;
        public string Formula { get; set; } = null;
        public bool VerticalText { get; set; } = false;
        public string NumberFormat { get; set; }
        public double? Width { get; set; } = null;

        public HorizontalAlignment HorizontalAligment { get; set; } = HorizontalAlignment.Center;
        public VerticalAlignment VerticalAligment { get; set; } = VerticalAlignment.Center;
    }
}
