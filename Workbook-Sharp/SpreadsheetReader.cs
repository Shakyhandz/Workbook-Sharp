using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp;

public class SpreadsheetReader
{

    //https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-parse-and-read-a-large-spreadsheet?tabs=cs-2%2Ccs-3%2Ccs-4%2Ccs


    public void Read(string fileName)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false);

        var workBookPart = spreadsheetDocument.WorkbookPart;

        foreach (var sheet in workBookPart?.Workbook?.Descendants<Sheet>() ?? [])
        {
            if (sheet?.Id != null && workBookPart?.GetPartById(sheet.Id!) is WorksheetPart wsPart)
            {
                foreach (Row row in wsPart.Worksheet.Descendants<Row>())
                {
                    List<object> rowData = new List<object>();
                    string? value;

                    foreach (Cell c in row.Elements<Cell>())
                    {
                        value = GetCellValue(c);
                        rowData.Add(value ?? "");
                    }
                }
            }
        }
    }

    public static string? GetCellValue(Cell cell)
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
                    ? "FALSE"
                    : "TRUE";
        }

        return value;
    }
}