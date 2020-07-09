using System;

namespace DelimitedFileSqlServerTableGenerator.Extensions
{
    internal static class TypeExtensions
    {
        internal static bool Is<T>(this Type input)
        {
            return input == typeof(T);
        }
    }
}
