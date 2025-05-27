using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp;

internal static class WorksheetElements
{
    // Correct Order of Elements in a Worksheet
    private static readonly List<Type> _orderOfElements = new()
    {
        typeof(SheetProperties),
        typeof(SheetDimension),
        typeof(SheetViews),
        typeof(SheetFormatProperties),
        typeof(Columns),
        typeof(SheetData),
        typeof(SheetProtection),
        typeof(ProtectedRanges),
        typeof(Scenarios),
        typeof(AutoFilter),
        typeof(SortState),
        typeof(DataConsolidate),
        typeof(CustomSheetViews),
        typeof(MergeCells),
        typeof(PhoneticProperties),
        typeof(ConditionalFormatting),
        typeof(DataValidations),
        typeof(Hyperlinks),
        typeof(PrintOptions),
        typeof(PageMargins),
        typeof(PageSetup),
        typeof(HeaderFooter),
        typeof(RowBreaks),
        typeof(ColumnBreaks),
        typeof(CustomProperties),
        typeof(CellWatches),
        typeof(IgnoredErrors),
        //typeof(SmartTags),
        typeof(Drawing),
        typeof(LegacyDrawing),
        typeof(LegacyDrawingHeaderFooter),
        typeof(Picture),
        typeof(OleObjects),
        typeof(Controls),
        typeof(WebPublishItems),
        typeof(TableParts),
        typeof(WorksheetExtensionList)
    };

    /// <summary>
    /// Returns the first matching child element of type T from the worksheet.
    /// If it doesn't exist, a new one is inserted in schema-compliant order.
    /// </summary>
    internal static T GetOrInsertWorksheetElement<T>(this DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet) where T : OpenXmlElement, new()
    {
        // Check if element already exists
        var existing = worksheet.Elements<T>().FirstOrDefault();

        if (existing is not null)
            return existing;

        // Validate type is allowed
        int newIndex = _orderOfElements.IndexOf(typeof(T));

        if (newIndex == -1)
            throw new InvalidOperationException($"Element type {typeof(T).Name} is not valid for Worksheet children.");

        var newElement = new T();

        // Insert in correct position
        foreach (var child in worksheet.Elements())
        {
            int childIndex = _orderOfElements.IndexOf(child.GetType());

            if (childIndex > newIndex)
            {
                worksheet.InsertBefore(newElement, child);
                return newElement;
            }
        }

        // If no such child exists, append to the end
        worksheet.Append(newElement);
        return newElement;
    }
}
