using WorkbookSharp.Styles;

namespace WorkbookSharp;

public static class WorkbookExtensions
{
    public static async Task<byte[]> ToExcel<T>(
        this IEnumerable<T> collection,
        string sheetName = "",
        uint headerRow = 1,
        uint startColumn = 1,
        Style? headerStyle = null,
        Style? itemsStyle  = null,
        Dictionary<uint, Style>? columnStyles = null,
        XlFontFamily? fontFamily = null,
        double? fontSize = null)
    {
        if (headerRow < 1)
            throw new ArgumentException($"Header row ({headerRow}) can't be less than 1");

        if (startColumn < 1)
            throw new ArgumentException($"Start column ({startColumn}) can't be less than 1");

        var workbook = new Workbook();
        var worksheet = workbook.AddWorksheet(sheetName);

        if (fontFamily != null)
            worksheet.FontFamily = fontFamily.Value;

        worksheet.FontSize = fontSize;

        var props = typeof(T).GetProperties();

        // Headers                
        var column = startColumn;
        var row = headerRow;
        
        headerStyle ??= new Style
        {
            FontDecoration = XlFontDecoration.Bold,
            FontSize = (fontSize ?? 11) + 1,
            FillColor = System.Drawing.Color.FromArgb(0xE6, 0xE6, 0xE6),
            Border = XlBorder.Around,
            HorizontalAlignment = XlHorizontalAlignment.Center
        };

        itemsStyle ??= new Style
        {
            Border = XlBorder.Around,
        };

        foreach (var prop in props)
            worksheet.SetValue((headerRow, column++), prop.Name, headerStyle.Clone());

        row++;

        var lastRow = (uint)collection.Count() + headerRow;

        // Items
        collection.Select(x => new
        {
            Values = props.Select((y, index) => new
            {
                index = (uint)index,
                value = y.GetValue(x, null)
            })
            .ToList()
        })
        .ToList()
        .ForEach(x =>
        {
            x.Values.ForEach(y =>
            {
                var columnIndex = y.index + startColumn;
                var cellStyle = columnStyles != null && columnStyles.TryGetValue(columnIndex, out var style)
                                ? style.Clone()
                                : itemsStyle.Clone();

                worksheet.SetValue((row, columnIndex), y.value, cellStyle);
            });

            row++;
        });

        return await workbook.Save();
    }

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
