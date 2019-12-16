using System;

namespace simte.Common
{
    public interface ISupportSelect<out TSource, out TResult>
    {
        TResult Select(Func<TSource, object> selector);
    }
}
