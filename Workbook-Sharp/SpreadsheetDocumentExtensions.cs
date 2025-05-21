using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp;

internal static class SpreadsheetDocumentExtensions
{
    internal static int GetSharedStringIndex(this SpreadsheetDocument document, WorksheetPart worksheetPart, string text)
    {
        // Get the SharedStringTablePart and add the result to it
        var shareStringPart = document.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().Count() > 0
                              ? document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First()
                              : document.WorkbookPart!.AddNewPart<SharedStringTablePart>();

        // Insert the result into the SharedStringTablePart
        return InsertSharedStringItem(text, shareStringPart);
    }

    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create it.
        if (shareStringPart.SharedStringTable is null)
            shareStringPart.SharedStringTable = new SharedStringTable();

        int index = 0;

        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
                return index; // The text already exists in the part. Return its index

            index++;
        }

        // The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

        return index;
    }

}
