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

    internal static void InsertWorksheetElementInOrder(this DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, OpenXmlElement newElement)
    {
        var newElementType = newElement.GetType();
        int newIndex = _orderOfElements.IndexOf(newElementType);

        if (newIndex == -1)
            throw new InvalidOperationException("Element type is not valid for Worksheet children.");

        // Find the first existing child that comes after the new element
        foreach (var child in worksheet.Elements())
        {
            int childIndex = _orderOfElements.IndexOf(child.GetType());

            if (childIndex > newIndex)
            {
                worksheet.InsertBefore(newElement, child);
                return;
            }
        }

        // If no such child exists, append to the end
        worksheet.Append(newElement);
    }

}
