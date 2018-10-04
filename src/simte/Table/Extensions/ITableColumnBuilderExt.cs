using System;

namespace simte.Table.Extensions
{
    public static class ITableColumnBuilderExt
    {
        public static ITableRowBuilder Formula(this ITableColumnBuilder source, string formula,
            Action<ColumnOptionsBuilder> action = null)
        {
            var opt = action == null
                ? ((ColumnOptionsBuilder builder) => builder.Formula(formula))
                : action += x => x.Formula(formula);

            return source.Column(opt);
        }

        //public static ITableRowBuilder Sum(this ITableRowBuilder tableRowBuilder, 
        //    Action<ColumnOptionsBuilder> action = null)
        //{

        //    return Formula("", action);
        //}
    }
}
