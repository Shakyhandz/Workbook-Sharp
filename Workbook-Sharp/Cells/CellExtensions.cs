using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp.Cells;

internal static class CellExtensions
{
    internal static string? GetValue(this Cell? cell)
    {
        if (cell == null)
            return null;

        if (cell.DataType == null)
            return cell.InnerText;

        var value = cell.InnerText;

        if (cell.DataType.Value == CellValues.SharedString) // For shared strings, look up the value in the shared strings table.
        {
            // Get worksheet from cell
            var parent = cell?.Parent;

            while (parent?.Parent != null && parent.Parent != parent && string.Compare(parent.LocalName, "worksheet", true) != 0)
            {
                parent = parent.Parent;
            }

            if (string.Compare(parent?.LocalName, "worksheet", true) != 0)
            {
                throw new Exception("Unable to find parent worksheet.");
            }

            if (parent is DocumentFormat.OpenXml.Spreadsheet.Worksheet ws &&
                ws?.WorksheetPart?.OpenXmlPackage is SpreadsheetDocument ssDoc)
            {
                var sstPart = ssDoc?.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                // lookup value in shared string table
                if (sstPart != null && sstPart.SharedStringTable != null)
                {
                    value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
        }
        else if (cell.DataType.Value == CellValues.Boolean)
        {
            value = value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                    ? "false"
                    : "true";
        }

        return value;
    }
}
