using simte.SeedWork;
using simte.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace simte.EPPlus.Table
{
    public class TableBuilder : ITableBuilder
    {
        private readonly TablePositionFinder _tablePositionFinder;
        private readonly WorksheetFactory _worksheetFactory;

        public TableOptions Options { get; }

        // ctor
        public TableBuilder(WorksheetFactory worksheetFactory, TableOptions options)
        {
            _worksheetFactory = worksheetFactory ?? throw new ArgumentNullException(nameof(worksheetFactory));
            Options = options ?? throw new ArgumentNullException(nameof(options));
            _tablePositionFinder = new TablePositionFinder(Options.TopLeft);

            prepare();
        }

        private void prepare()
        {
            if (Options.FreezePane.HasValue)
            {
                _worksheetFactory.ws.View.FreezePanes(
                    Options.FreezePane.Value.Row,
                    Options.FreezePane.Value.Col
                );

                // set row height
                if (Options.RowHeightMap.Any())
                {
                    var ws = _worksheetFactory.ws;
                    foreach (var pair in Options.RowHeightMap)
                    {
                        ws.Row(pair.Key).Height = pair.Value;
                    }
                }
            }
        }

        public ITableBuilder AddRows(Action<ITableRowBuilder> builderAction)
        {
            var rowBuilder = new RowBuilder(_worksheetFactory.ws, Options.TopLeft, _tablePositionFinder);
            builderAction(rowBuilder);

            return this;
        }

        public ITableBuilder AddRows<TModel>(Func<IEnumerable<TModel>> dataFunc,
            Action<ITableRowBuilder<TModel>> builderAction)
        {
            var rowBuilder = new RowBuilder<TModel>(_worksheetFactory.ws, Options.TopLeft, _tablePositionFinder, dataFunc);
            builderAction(rowBuilder);

            return this;
        }

        public IExcelPackage Attach()
            => _worksheetFactory.Attach();

    }
}
