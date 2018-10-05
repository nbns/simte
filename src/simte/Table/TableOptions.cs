using System.Collections.Generic;

namespace simte.Table
{
    public class TableOptions
    {
        public Position TopLeft { get; set; }
        public bool Border { get; set; }
        public bool Autofit { get; set; }

        public Position? FreezePane { get; set; }
        public bool WrapText { get; set; }

        public Dictionary<int, double> RowHeightMap => new Dictionary<int, double>();
    }
}
