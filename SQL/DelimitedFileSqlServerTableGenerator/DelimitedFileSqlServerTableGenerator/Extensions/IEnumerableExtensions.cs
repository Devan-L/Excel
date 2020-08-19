using System.Collections.Generic;

namespace DelimitedFileSqlServerTableGenerator.Extensions
{
    internal static class IEnumerableExtensions
    {
        internal static string Join<T>(this IEnumerable<T> collection, string separator)
        {
            return string.Join(separator, collection);
        }
    }
}
