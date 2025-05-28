using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorkbookSharp.Cells;

namespace WorkbookSharp;

internal static class WorksheetPartExtensions
{
    internal static void AutoSizeCells(this WorksheetPart worksheetPart, Dictionary<uint, double> maxColumnWidths)
    {
        var columns = worksheetPart.Worksheet.GetOrInsertWorksheetElement<Columns>();

        foreach (var kvp in maxColumnWidths.OrderBy(x => x.Key))
        {
            columns.Append(new Column
            {
                Min = kvp.Key,
                Max = kvp.Key,
                Width = kvp.Value,
                CustomWidth = true
            });
        }
    }

    internal static void ShowGridlines(this WorksheetPart worksheetPart, bool showGridlines)
    {
        var sheetViews = worksheetPart.Worksheet.GetOrInsertWorksheetElement<SheetViews>();

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
        
        if (sheetView == null)
        {
            sheetView = new SheetView
            {
                WorkbookViewId = 0,
                ShowGridLines = showGridlines
            };

            sheetViews.Append(sheetView);
        }
        else
        {
            sheetView.ShowGridLines = showGridlines;
        }
    }

    internal static Cell GetOrInsertCellInWorksheet(this WorksheetPart worksheetPart, CellReference cellReference)
    {
        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

        // Get or create the row
        var row = sheetData.Elements<Row>().Where(r => r.RowIndex?.Value == cellReference.RowIndex).FirstOrDefault();
        row ??= sheetData.AppendChild(new Row { RowIndex = cellReference.RowIndex });

        // Get or create the cell
        var cell = row.Elements<Cell>().Where(c => c.CellReference?.Value == cellReference.Address).FirstOrDefault();
        if (cell != null)
            return cell;

        // Find the correct insertion point to keep cells in order
        var refCell = row.Elements<Cell>().Where(c => string.Compare(c.CellReference?.Value, cellReference.Address, StringComparison.OrdinalIgnoreCase) > 0).FirstOrDefault();

        var newCell = new Cell { CellReference = cellReference.Address };
        row.InsertBefore(newCell, refCell);

        return newCell;
    }
}
