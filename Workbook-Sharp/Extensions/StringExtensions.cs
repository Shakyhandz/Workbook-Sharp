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
    /// <param name="input">String to clean</param>
    /// <param name="charsToClean">Chars that should be removed from the string</param>
    /// <returns>A cleaned string</returns>    
    public static string Clean(this string? input, string charsToClean) => Clean(input, [.. charsToClean]);

    /// <summary>
    /// Clean the string from chars 
    /// </summary>
    /// <param name="input">String to clean</param>
    /// <param name="charsToClean">Chars that should be removed from the string</param>
    /// <returns>A cleaned string</returns>  
    public static string Clean(this string? input, HashSet<char> charsToClean) => input.IsNothing() ? "" : new string(input!.Where(c => !charsToClean.Contains(c)).ToArray());
    
    /// <summary>
    /// Validates a string to see if it's an email address
    /// </summary>
    /// <param name="email">String to check</param>
    /// <returns>true if the string is an email address</returns>
    public static bool IsEmailAddress(this string email) => email.IsSome() && _matchEmail.IsMatch(email);

    /// <summary>
    /// Checks if the string is null, empty or just contains whitespaces
    /// </summary>
    /// <param name="thisValue">String to check</param>
    /// <returns>True if nothing</returns>
    public static bool IsNothing(this string? thisValue) => thisValue == null || string.IsNullOrEmpty(thisValue.Trim());

    /// <summary>
    /// Checks that the string is not null, empty or just containing whitespaces
    /// </summary>
    /// <param name="thisValue">String to check</param>
    /// <returns>True if the string is something</returns>
    public static bool IsSome(this string? thisValue) => !thisValue.IsNothing();

    /// <summary>
    /// Check if the string has only numeric characters. Use case: for really long numbers where long.TryParse fails
    /// </summary>
    /// <param name="thisValue">String to check</param>
    /// <returns>True if the string just has numeric characters</returns>
    public static bool IsNumeric(this string? value) => value.IsSome() && value!.All(char.IsDigit);

    /// <summary>
    /// Truncate a string if it exceeds a threshold value
    /// </summary>
    /// <param name="thisValue">String to truncate</param>
    /// <param name="length">Length of the resulting string</param>
    /// <param name="add"></param>
    /// <returns>Truncated string</returns>
    public static string Truncate(this string? thisValue, int length, string add = "") => thisValue.IsNothing() ? "" :
                                                                                          thisValue!.Length <= length ? thisValue : 
                                                                                          thisValue.Substring(0, length) + add;
                                                
    /// <summary>
    /// Parse a string and convert it to a 32 bit integer
    /// </summary>
    /// <param name="thisValue">String to parse</param>
    /// <returns>null if the string wasn't a integer, otherwise the parsed integer</returns>
    public static int? ToInt32(this string? thisValue) => thisValue.IsSome() && int.TryParse(thisValue, out var i) ? i : null;
    
    /// <summary>
    /// Parse a string and convert it to a decimal
    /// </summary>
    /// <param name="thisValue">String to parse</param>
    /// <returns>null if the string wasn't a decimal, otherwise the parsed decimal</returns>
    public static decimal? ToDecimal(this string? thisValue) => thisValue.IsSome() && decimal.TryParse(thisValue, out var i) ? i : null;
    
    /// <summary>
    /// Parse a string and convert it to a 64 bit long
    /// </summary>
    /// <param name="thisValue">String to parse</param>
    /// <returns>null if the string wasn't a long, otherwise the parsed long</returns>
    public static long? ToInt64(this string thisValue) => thisValue.IsSome() && long.TryParse(thisValue, out var i) ? i : null;
    
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
    /// <param name="thisValue">The string</param>
    /// <param name="strings">The strings to replace</param>
    /// <returns></returns>
    public static string Remove(this string? thisValue, IEnumerable<string> strings) => thisValue.IsNothing() ? "" : strings.Aggregate(thisValue!, (current, s) => current.Replace(s, ""));

    /// <summary>
    /// Replace the last occurrence of a string with another string
    /// </summary>
    /// <param name="text"></param>
    /// <param name="search"></param>
    /// <param name="replace"></param>
    /// <returns></returns>
    public static string ReplaceLastOccurrence(this string? text, string search, string replace)
    {
        if (text.IsNothing())
            return "";

        int pos = text!.LastIndexOf(search);

        if (pos < 0)
            return text;

        return text.Substring(0, pos) + 
               replace + 
               text.Substring(pos + search.Length);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="text"></param>
    /// <param name="search"></param>
    /// <param name="replace"></param>
    /// <returns></returns>
    public static string ReplaceFirstOccurrence(this string? text, string search, string replace)
    {
        if (text.IsNothing())
            return "";

        int pos = text!.IndexOf(search);

        if (pos < 0)
            return text;

        return text.Substring(0, pos) + 
               replace + 
               text.Substring(pos + search.Length);
    }
}
