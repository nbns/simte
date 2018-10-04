using System.Collections.Generic;
using System.Drawing;

namespace simte.Table
{
    public class RowOptions
    {
        public double? Height { get; set; } = null;
        public Color TextColor { get; set; } = Color.Black;
        public Color BackgroundColor { get; set; } = Color.White;

        public Dictionary<int, double> HeightByIndex { get; }
            = new Dictionary<int, double>();
    }
}
