namespace WorkbookSharp;

internal static class WorksheetExtensions
{
    private const int EXCEL_MAX_SHEET_NAME_LENGTH = 31;
    internal static string GetNewSheetNameSafe(this List<Worksheet> sheets, string sheetName)
    {
        // Remove illegal characters
        var illegalCharacters = new[] { '\\', '/', '*', '[', ']', ':', '?', }.ToList();
        sheetName = sheetName.Where(x => !illegalCharacters.Contains(x)).StringJoin("");

        // Sheet name can be max 31 characters 
        sheetName = sheetName.Truncate(EXCEL_MAX_SHEET_NAME_LENGTH);

        // Increment sheet name if it already exists (yes, start with 2)
        var sheetCounter = 2;

        while (sheets.Any(x => x.SheetName.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase)))
        {
            var suffix = $" ({sheetCounter})";

            // Remove counter suffix
            if (sheetName.EndsWith(suffix))
            {
                sheetName = sheetName.ReplaceLastOccurrence(suffix, "");
                suffix = $" ({++sheetCounter})";
            }

            sheetName = sheetName.Truncate(EXCEL_MAX_SHEET_NAME_LENGTH - suffix.Length) + suffix;
        }

        return sheetName;
    }
}
