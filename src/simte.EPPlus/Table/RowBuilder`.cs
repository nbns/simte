using OfficeOpenXml;
using simte.SeedWork;
using simte.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace simte.EPPlus.Table
{
    internal class RowBuilder<TSource> : BaseColumnBuilder, ITableRowBuilder<TSource>
    {
        private int _currentRow = 0;
        private readonly Func<IEnumerable<TSource>> _dataFunc;
        private readonly TablePositionFinder _tablePositionFinder;
        private Position _topLeft;

        private class ColumnSource
        {
            public Func<TSource, object> SelectorFunc { get; set; }
            public Action<ColumnOptionsBuilder<TSource>> ActionOptions { get; set; }

            public int Row { get; set; }
        }

        private List<ColumnSource> _columnSources = new List<ColumnSource>();

        // ctor
        public RowBuilder(
            ExcelWorksheet ws,
            Position topLeft,
            TablePositionFinder tablePositionFinder,
            Func<IEnumerable<TSource>> dataFunc) : base(ws)
        {
            _topLeft = topLeft;
            _tablePositionFinder = tablePositionFinder ?? throw new ArgumentNullException(nameof(tablePositionFinder));
            _dataFunc = dataFunc ?? throw new ArgumentNullException(nameof(dataFunc));
        }

        internal void AddColumn(Func<TSource, object> selector, Action<ColumnOptionsBuilder<TSource>> action = null)
            => _columnSources.Add(new ColumnSource
            {
                SelectorFunc = selector,
                ActionOptions = action,
                Row = _currentRow,
            });

        public ITableColumnBuilder<TSource> Row
        {
            get
            {
                ++_currentRow;
                return new TableColumnBuilder<TSource>(this);
            }
        }

        public ITableRowBuilder<TSource> Apply()
        {
            int indexColumn = 0;
            var data = _dataFunc().ToList();

            foreach (var item in data)
            {
                foreach (var row in _columnSources.GroupBy(x => x.Row).ToList())
                {                    
                    indexColumn = _tablePositionFinder.GetColumnForNewRow();
                    foreach (var column in row)
                    {
                        var columnOptionsBuilder = new ColumnOptionsBuilder<TSource>();
                        column.ActionOptions?.Invoke(columnOptionsBuilder);
                        ColumnOptions options = columnOptionsBuilder;

                        // set color if present function
                        options.TextColor = columnOptionsBuilder.TextColorFunc?.Invoke(item) ?? options.TextColor;
                        options.BackgroundColor = columnOptionsBuilder.BackgroundColorFunc?.Invoke(item) ?? options.BackgroundColor;

                        var pos = _tablePositionFinder.GetNewPosition(indexColumn, options.Colspan, options.Rowspan);

                        setExcelRange(column.SelectorFunc?.Invoke(item) ?? "", pos, options);
                        indexColumn = indexColumn + options.Colspan;
                    }
                }
            }

            return this;
        }
    }
}
