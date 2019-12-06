using System;
using System.Drawing;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("simte.EPPlus")]

namespace simte.Table
{
    public class ColumnOptionsBuilder<TSource> : ColumnOptionsBuilder
    {
        protected internal Func<TSource, Color> TextColorFunc { get; private set; }
        protected internal Func<TSource, Color> BackgroundColorFunc { get; private set; }

        public ColumnOptionsBuilder<TSource> TextColorIf(Func<TSource, Color> selector)
        {
            TextColorFunc = selector;
            return this;
        }

        public ColumnOptionsBuilder<TSource> TextColorIf(Func<TSource, bool> predicate, Color color)
        {
            TextColorFunc = x => predicate(x) ? color : Color.Black;
            return this;
        }

        public ColumnOptionsBuilder<TSource> BackgroundColorIf(Func<TSource, Color> selector)
        {
            BackgroundColorFunc = selector;
            return this;
        }

        public ColumnOptionsBuilder<TSource> BackgroundColorIf(Func<TSource, bool> predicate, Color color)
        {
            BackgroundColorFunc = x => predicate(x) ? color : Color.White;
            return this;
        }
    }
}