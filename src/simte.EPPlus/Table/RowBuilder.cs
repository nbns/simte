using OfficeOpenXml;
using simte.SeedWork;
using simte.Table;
using System;
using System.Linq;

namespace simte.EPPlus.Table
{
    internal sealed class RowBuilder : ITableRowBuilder
    {
        private readonly ExcelWorksheet _ws;
        private readonly Position _topLeft;
        private readonly TablePositionFinder _tablePositionFinder;

        private int _currentColumn = 0;

        // ctor
        public RowBuilder(ExcelWorksheet ws, 
            Position topLeft,
            TablePositionFinder tablePositionFinder)
        {
            _ws = ws ?? throw new ArgumentNullException(nameof(ws));

            _topLeft = topLeft;
            _tablePositionFinder = tablePositionFinder
                ?? throw new ArgumentNullException(nameof(tablePositionFinder));
        }

        public ITableRowBuilder RowSettings(Action<RowOptions> settings)
        {
            var rowIndex = _tablePositionFinder.GetLastPosition().Row;
            var currentRow = _ws.Row(rowIndex);

            var options = new RowOptions();
            settings(options);

            if (options.Height.HasValue)
            {
                currentRow.Height = options.Height.Value;
            }

            currentRow.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            currentRow.Style.Font.Color.SetColor(options.TextColor);
            currentRow.Style.Fill.BackgroundColor.SetColor(options.BackgroundColor);

            return this;
        }

        public ITableColumnBuilder Row
        {
            get
            {
                _currentColumn = _tablePositionFinder.GetColumnForNewRow();
                return new TableColumnBuilder(this, _ws);
            }
        }

        internal Position GetPositionForCurrentColumn(int colspan, int rowspan)
        {
            return _tablePositionFinder.GetNewPosition(_currentColumn, colspan, rowspan);
        }

        internal void NextColumn(int colspan)
        {
            _currentColumn = _currentColumn + colspan;

            if (_currentColumn < 1 && _currentColumn > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("Column number out of bounds");
            }
        }
    }
}
