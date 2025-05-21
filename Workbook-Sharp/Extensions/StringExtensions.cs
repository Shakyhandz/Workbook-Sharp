using System.Text.RegularExpressions;

namespace WorkbookSharp;

/// <summary>
/// Extensions to the .NET string class
/// </summary>
public static class StringExtensions
{
    private static readonly Regex _matchEmail = new Regex(@"^([a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// Clean the string from chars 
    /// </summary>
    /// <param name="str">String to clean</param>
    /// <param name="charsToClean">Chars that should be removed from the string</param>
    /// <returns>A cleaned string</returns>    
    public static string Clean(this string? str, string charsToClean) => Clean(str, [.. charsToClean]);

    /// <summary>
    /// Clean the string from chars 
    /// </summary>
    /// <param name="str">String to clean</param>
    /// <param name="charsToClean">Chars that should be removed from the string</param>
    /// <returns>A cleaned string</returns>  
    public static string Clean(this string? str, HashSet<char> charsToClean) => str.IsNothing() ? "" : new string(str!.Where(c => !charsToClean.Contains(c)).ToArray());
    
    /// <summary>
    /// Validates a string to see if it's an email address
    /// </summary>
    /// <param name="str">String to check</param>
    /// <returns>true if the string is an email address</returns>
    public static bool IsEmailAddress(this string? str) => str.IsSome() && _matchEmail.IsMatch(str!);

    /// <summary>
    /// Checks if the string is null, empty or just contains whitespaces
    /// </summary>
    /// <param name="str">String to check</param>
    /// <returns>True if nothing</returns>
    public static bool IsNothing(this string? str) => str == null || string.IsNullOrEmpty(str.Trim());

    /// <summary>
    /// Checks that the string is not null, empty or just containing whitespaces
    /// </summary>
    /// <param name="str">String to check</param>
    /// <returns>True if the string is something</returns>
    public static bool IsSome(this string? str) => !str.IsNothing();

    /// <summary>
    /// Check if the string has only numeric characters. Use case: for really long numbers where long.TryParse fails
    /// </summary>
    /// <param name="str">String to check</param>
    /// <returns>True if the string just has numeric characters</returns>
    public static bool IsNumeric(this string? str) => str.IsSome() && str!.All(char.IsDigit);

    /// <summary>
    /// Truncate a string if it exceeds a threshold value
    /// </summary>
    /// <param name="str">String to truncate</param>
    /// <param name="length">Length of the resulting string</param>
    /// <param name="add"></param>
    /// <returns>Truncated string</returns>
    public static string Truncate(this string? str, int length, string add = "") => str.IsNothing() || str!.Length <= length ? str ?? "" :
                                                                                    str.Substring(0, length) + add;
                                                
    /// <summary>
    /// Parse a string and convert it to a 32 bit integer
    /// </summary>
    /// <param name="str">String to parse</param>
    /// <returns>null if the string wasn't a integer, otherwise the parsed integer</returns>
    public static int? ToInt32(this string? str) => str.IsSome() && int.TryParse(str, out var i) ? i : null;
    
    /// <summary>
    /// Parse a string and convert it to a decimal
    /// </summary>
    /// <param name="str">String to parse</param>
    /// <returns>null if the string wasn't a decimal, otherwise the parsed decimal</returns>
    public static decimal? ToDecimal(this string? str) => str.IsSome() && decimal.TryParse(str, out var i) ? i : null;
    
    /// <summary>
    /// Parse a string and convert it to a 64 bit long
    /// </summary>
    /// <param name="str">String to parse</param>
    /// <returns>null if the string wasn't a long, otherwise the parsed long</returns>
    public static long? ToInt64(this string? str) => str.IsSome() && long.TryParse(str, out var i) ? i : null;
    
    /// <summary>
    /// If the string is nothing throw a argument exception
    /// </summary>
    /// <param name="str">string to check</param>
    /// <param name="message">Exception message</param>
    /// <param name="args">Format parameters</param>
    public static void ThrowOnNothing(this string? str, string message, params object[] args)
    {
        if (str.IsNothing())
            throw new ArgumentException(string.Format(message, args));
    }

    /// <summary>
    /// Replace all occurrences of a string with another string
    /// </summary>
    /// <param name="str">The string</param>
    /// <param name="strings">The strings to replace</param>
    /// <returns></returns>
    public static string Remove(this string? str, IEnumerable<string> strings) => str.IsNothing() ? "" : strings.Aggregate(str!, (current, s) => current.Replace(s, ""));

    /// <summary>
    /// Replaces the last occurrence of a specified substring with a new value.
    /// </summary>
    /// <param name="str">The source string to operate on.</param>
    /// <param name="search">The substring to find and replace.</param>
    /// <param name="replace">The replacement string.</param>
    /// <returns>
    /// A new string with the last occurrence of <paramref name="search"/> replaced by <paramref name="replace"/>,
    /// or the original string if <paramref name="search"/> is not found.
    /// Returns an empty string if <paramref name="str"/> is null or empty.
    /// </returns>
    public static string ReplaceLastOccurrence(this string? str, string search, string replace)
    {
        if (str.IsNothing() || string.IsNullOrEmpty(search))
            return "";

        int pos = str!.LastIndexOf(search, StringComparison.Ordinal);

        if (pos < 0)
            return str;

        return string.Concat(str.AsSpan(0, pos),
                             replace,
                             str.AsSpan(pos + search.Length));
    }

    /// <summary>
    /// Replaces the last occurrence of a specified substring with a new value.
    /// </summary>
    /// <param name="str">The source string to operate on.</param>
    /// <param name="search">The substring to find and replace.</param>
    /// <param name="replace">The replacement string.</param>
    /// <returns>
    /// A new string with the first occurrence of <paramref name="search"/> replaced by <paramref name="replace"/>,
    /// or the original string if <paramref name="search"/> is not found.
    /// Returns an empty string if <paramref name="str"/> is null or empty.
    /// </returns>
    public static string ReplaceFirstOccurrence(this string? str, string search, string replace)
    {
        if (str.IsNothing() || string.IsNullOrEmpty(search))
            return "";

        int pos = str!.IndexOf(search, StringComparison.Ordinal);

        if (pos < 0)
            return str;

        return string.Concat(str.AsSpan(0, pos),
                             replace,
                             str.AsSpan(pos + search.Length));
    }
}
