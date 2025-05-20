using System.Text.RegularExpressions;

namespace WorkbookSharp;

/// <summary>
/// Extensions to the .NET string class
/// </summary>
public static class StringExtensions
{
    /// <summary>
    /// Clean the string from chars that don't belong
    /// </summary>
    /// <param name="input">String to clean</param>
    /// <param name="charsToClean">chars that should be removed from the string</param>
    /// <returns>A cleaned string</returns>
    public static string Clean(this string input, string charsToClean)
    {
        if (input == null)
            throw new ArgumentException("input can't be null");

        if (charsToClean == null)
            throw new ArgumentException("chars To Clean can't be null");

        if (!String.IsNullOrEmpty(input))
        {
            for (int i = 0; i < charsToClean.Length; i++)
            {
                input = input.Replace(charsToClean.Substring(i, 1), String.Empty);
            }
        }
        return input;
    }

    /// <summary>
    /// Validates a string to see if it's an email address
    /// </summary>
    /// <param name="email">String to check</param>
    /// <returns>true if the string is an email address</returns>
    public static bool IsEmailAddress(this string email)
    {
        if (string.IsNullOrEmpty(email)) return false;

        //string emailPattern = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" +
        //    @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
        //    @".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";
        string emailPattern = @"^([a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)$";

        return Regex.Match(email, emailPattern, RegexOptions.IgnoreCase).Success;
    }

    /// <summary>
    /// Checks if the string is null, empty or just contains whitespaces
    /// </summary>
    /// <param name="thisValue">String to check</param>
    /// <returns>True if nothing</returns>
    public static bool IsNothing(this string? thisValue)
    {
        if (thisValue == null) 
            return true;

        return string.IsNullOrEmpty(thisValue.Trim());
    }

    /// <summary>
    /// Checks that the string is not null, empty or just containing whitespaces
    /// </summary>
    /// <param name="thisValue">String to check</param>
    /// <returns>True if the string is something</returns>
    public static bool IsSome(this string? thisValue)
    {
        return !thisValue.IsNothing();
    }

    /// <summary>
    /// Check if the string just have numeric characters
    /// </summary>
    /// <param name="thisValue">String to check</param>
    /// <returns>True if the string just has numeric characters</returns>
    public static bool IsNumeric(this string thisValue)
    {
        if (string.IsNullOrEmpty(thisValue)) return false;

        string nbrs = @"0123456789";

        return (from c in thisValue.ToCharArray()
                join n in nbrs.ToCharArray() on c equals n into j
                where j.Count() == 0
                select c).Count() <= 0;
    }
    
    /// <summary>
    /// Truncate a string if it's exceeds a threashold value
    /// </summary>
    /// <param name="thisValue">String to truncate</param>
    /// <param name="length">Length of the resulting string</param>
    /// <param name="add"></param>
    /// <returns>Truncated string</returns>
    public static string Truncate(this string thisValue, int length, string add = "")
    {
        if (thisValue == null)
            throw new ArgumentException("input can't be null");

        if (thisValue.Length <= length)
            return thisValue;
        return thisValue.Substring(0, length) + add;
    }
    
    /// <summary>
    /// Parse a string and convert it to a 32 bit integer
    /// </summary>
    /// <param name="thisValue">String to parse</param>
    /// <returns>null if the string wasn't a integer, otherwise the parsed integer</returns>
    public static int? ToInt32(this string thisValue)
    {
        if (thisValue.IsNothing()) return null;
        int i;
        var ok = Int32.TryParse(thisValue, out i);
        return ok ? (int?)i : null;
    }
    
    /// <summary>
    /// Parse a string and convert it to a decimal
    /// </summary>
    /// <param name="thisValue">String to parse</param>
    /// <returns>null if the string wasn't a decimal, otherwise the parsed decimal</returns>
    public static decimal? ToDecimal(this string thisValue)
    {
        if (thisValue.IsNothing()) return null;
        decimal i;
        var ok = decimal.TryParse(thisValue, out i);
        return ok ? (decimal?)i : null;
    }
    
    /// <summary>
    /// Parse a string and convert it to a 64 bit long
    /// </summary>
    /// <param name="thisValue">String to parse</param>
    /// <returns>null if the string wasn't a long, otherwise the parsed long</returns>
    public static long? ToInt64(this string thisValue)
    {
        if (thisValue.IsNothing()) return null;
        long i;
        var ok = long.TryParse(thisValue, out i);
        return ok ? (long?)i : null;
    }
    
    /// <summary>
    /// if the string is nothing throw a komon argument exception
    /// </summary>
    /// <param name="obj">string to check</param>
    /// <param name="message">Exception message</param>
    /// <param name="args">Format parameters</param>
    public static void ThrowOnNothing(this string obj, string message, params object[] args)
    {
        if (obj.IsNothing())
        {
            throw new ArgumentException(string.Format(message, args));
        }
    }
    
    /// <summary>
    /// 
    /// </summary>
    /// <param name="thisValue"></param>
    /// <param name="strings"></param>
    /// <returns></returns>
    public static string Remove(this string thisValue, IEnumerable<string> strings)
    {
        if (thisValue == null)
            throw new ArgumentException("input can't be null");

        if (strings == null)
            throw new ArgumentException("strings can't be null");

        foreach (var item in strings)
        {
            thisValue = thisValue.Replace(item, "");
        }
        return thisValue;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="text"></param>
    /// <param name="search"></param>
    /// <param name="replace"></param>
    /// <returns></returns>
    public static string ReplaceLastOccurrence(this string text, string search, string replace)
    {
        int pos = text.LastIndexOf(search);
        if (pos < 0)
            return text;

        return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="text"></param>
    /// <param name="search"></param>
    /// <param name="replace"></param>
    /// <returns></returns>
    public static string ReplaceFirstOccurrence(this string text, string search, string replace)
    {
        int pos = text.IndexOf(search);
        if (pos < 0)
            return text;

        return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
    }
}
