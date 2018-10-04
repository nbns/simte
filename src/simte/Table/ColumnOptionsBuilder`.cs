using System;
using System.Drawing;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("simte.EPPlus")]

namespace simte.Table
{
    public class ColumnOptionsBuilder<TSource> : ColumnOptionsBuilder
    {
        internal Func<TSource, Color> TextColorFunc { get; set; }
        protected internal Func<TSource, Color> BackgroundColorFunc { get; set; }

        public ColumnOptionsBuilder<TSource> TextColorIf(Func<TSource, Color> selector)
        {
            TextColorFunc = selector;
            return this;
        }

        public ColumnOptionsBuilder<TSource> BackgroundColorIf(Func<TSource, Color> selector)
        {
            BackgroundColorFunc = selector;
            return this;
        }
    }
}
