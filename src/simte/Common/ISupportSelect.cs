using System;

namespace simte.Common
{
    public interface ISupportSelect<TSource, TResult>
    {
        TResult Select(Func<TSource, object> selector);
    }
}
