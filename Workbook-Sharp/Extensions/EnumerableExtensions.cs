namespace WorkbookSharp;

public static class EnumerableExtensions
{
    public static string StringJoin<T>(this IQueryable<T> q, string separator) => string.Join(separator, q);
    public static string StringJoin<T>(this IEnumerable<T> q, string separator) => string.Join(separator, q);    

}
