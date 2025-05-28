using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorkbookSharp.Cells;

namespace WorkbookSharp;

internal static class WorksheetPartExtensions
{
    internal static void SetCellValue(this WorksheetPart worksheetPart, SpreadsheetDocument document, CellObject cellObject)
    {
        CellValue? cellValue = null;
        EnumValue<CellValues>? dataType = null;

        // Set the value of the cell
        if (cellObject.Value is int i)
        {
            cellValue = new CellValue(i);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellObject.Value is decimal dec)
        {
            cellValue = new CellValue(dec);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellObject.Value is double d)
        {
            cellValue = new CellValue(d);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellObject.Value is long l)
        {
            cellValue = new CellValue((decimal)l);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (cellObject.Value is bool b)
        {
            cellValue = new CellValue(b);
            dataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (cellObject.Value is DateTime dt)
        {
            cellValue = new CellValue(dt);
            dataType = new EnumValue<CellValues>(CellValues.Date);
        }
        else
        {
            // TODO: Add null value if null?
            var index = document.GetSharedStringIndex(worksheetPart, cellObject.Value?.ToString() ?? "");
            cellValue = new CellValue(index.ToString());
            dataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        // Insert the cell value        
        Cell result = GetOrInsertCellInWorksheet(worksheetPart, cellObject.CellReference);

        result.CellValue = cellValue;
        result.DataType = dataType;

        if (cellObject.StyleIndex != null)
            result.StyleIndex = cellObject.StyleIndex;
    }

    internal static void SetCellFormula(this WorksheetPart worksheetPart, Cells.CellFormula formula)
    {
        var cell = GetOrInsertCellInWorksheet(worksheetPart, formula.CellReference);

        cell.CellReference = formula.CellReference.Address; // optional but helps with structure
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula { Text = formula.ParseFormula() };
        cell.StyleIndex = formula.StyleIndex;
    }

    internal static void SetCellRichText(this WorksheetPart worksheetPart, CellRichText richText)
    {
        var cell = GetOrInsertCellInWorksheet(worksheetPart, richText.CellReference);
        cell.CellReference = richText.CellReference.Address; 
        cell.InlineString = richText.InlineString;
        cell.DataType = CellValues.InlineString;
    }

    internal static void SetCellStyle(this WorksheetPart worksheetPart, Cells.CellStyle xlStyle)
    {
        // Set the style for the cell
        var cell = GetOrInsertCellInWorksheet(worksheetPart, xlStyle.CellReference);
        cell.StyleIndex = xlStyle.StyleIndex;
    }

    internal static void MergeCells(this WorksheetPart worksheetPart, CellMerge merge)
    {
        // Verify if the specified cells exist, and if they do not exist, create them.
        GetOrInsertCellInWorksheet(worksheetPart, merge.CellReference);
        GetOrInsertCellInWorksheet(worksheetPart, merge.ToCellReference);

        var mergeCells = worksheetPart.Worksheet.GetOrInsertWorksheetElement<MergeCells>();

        // Create the merged cell and append it to the MergeCells collection.
        MergeCell mergeCell = new MergeCell
        {
            Reference = new StringValue($"{merge.CellReference.Address}:{merge.ToCellReference.Address}")
        };

        mergeCells.Append(mergeCell);
    }

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

    private static Cell GetOrInsertCellInWorksheet(WorksheetPart worksheetPart, CellReference cellReference)
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
