using System.Text.RegularExpressions;

namespace WorkbookSharp.Cells;

public static class ExcelFormulaValidator
{
    private static readonly Regex CellRangeWithSpace = new Regex(@"[A-Z]+\d+ [A-Z]+\d+", RegexOptions.Compiled);
    private static readonly Regex SheetRefNoQuotes = new Regex(@"\b([^\s'\[]+\s[^\s'\[]+)!([A-Z]+\d+)", RegexOptions.Compiled);
    private static readonly Regex ExternalRef = new Regex(@"\[(.+\.xlsx)\]", RegexOptions.Compiled);
    private static readonly Regex FormattedNumber = new Regex(@"(?<!\d),(?=\d{3})", RegexOptions.Compiled); // commas in numbers
    private static readonly Regex FunctionCall = new Regex(@"\b[A-Z][A-Z0-9]*\(", RegexOptions.Compiled);

    public static List<string> ValidateFormula(string? formula)
    {
        var errors = new List<string>();

        if (formula.IsNothing())
        {
            errors.Add("Formula is empty");
            return errors;
        }

        // Parenthesis check
        int parentheses = 0;

        foreach (char c in formula!)
        {
            if (c == '(') 
                parentheses++;
            else if (c == ')') 
                parentheses--;

            if (parentheses < 0)
            {
                errors.Add("Parentheses closed before opening");
                break;
            }
        }

        if (parentheses > 0)
            errors.Add("Missing closing parenthesis");

        // Cell ranges using space
        if (CellRangeWithSpace.IsMatch(formula))
            errors.Add("Cell ranges must use ':' not space");

        // Sheet reference not wrapped in quotes
        if (SheetRefNoQuotes.IsMatch(formula))
            errors.Add("Sheet names with spaces must be wrapped in single quotes");

        // External reference detection
        if (formula.Contains(":\\") && !ExternalRef.IsMatch(formula))
            errors.Add("External references must include workbook name in brackets (e.g., [Workbook.xlsx])");

        // Comma in number
        if (FormattedNumber.IsMatch(formula))
            errors.Add("Numbers should not contain commas for formatting (e.g., use 1000 not 1,000)");

        // Count nesting depth (simple approximation)
        int depth = 0;

        foreach (Match match in FunctionCall.Matches(formula))
        {
            depth++;
        }

        if (depth > 64)
            errors.Add("Formula nests more than 64 functions");

        return errors;
    }

}
