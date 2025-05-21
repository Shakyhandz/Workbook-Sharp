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
        Cell result = GetCellInWorksheet(worksheetPart, cellObject.CellReference.ColumnName, cellObject.CellReference.RowIndex);

        result.CellValue = cellValue;
        result.DataType = dataType;

        if (cellObject.StyleIndex != null)
            result.StyleIndex = cellObject.StyleIndex;
    }

    internal static void SetCellFormula(this WorksheetPart worksheetPart, Cells.CellFormula formula)
    {
        var cell = GetCellInWorksheet(worksheetPart, formula.CellReference.ColumnName, formula.CellReference.RowIndex);

        cell.CellReference = formula.CellReference.Address; // optional but helps with structure
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula { Text = formula.ParseFormula() };
        cell.StyleIndex = formula.StyleIndex;
    }

    internal static void SetCellStyle(this WorksheetPart worksheetPart, Cells.CellStyle xlStyle)
    {
        // Set the style for the cell
        var cell = GetCellInWorksheet(worksheetPart, xlStyle.CellReference.ColumnName, xlStyle.CellReference.RowIndex);
        cell.StyleIndex = xlStyle.StyleIndex;
    }

    internal static void MergeCells(this WorksheetPart worksheetPart, CellMerge merge)
    {
        // Verify if the specified cells exist, and if they do not exist, create them.
        GetCellInWorksheet(worksheetPart, merge.CellReference.ColumnName, merge.CellReference.RowIndex);
        GetCellInWorksheet(worksheetPart, merge.ToCellReference.ColumnName, merge.ToCellReference.RowIndex);

        MergeCells mergeCells;
        var worksheet = worksheetPart.Worksheet;

        if (worksheet.Elements<MergeCells>().Count() > 0)
        {
            mergeCells = worksheet.Elements<MergeCells>().First();
        }
        else
        {
            mergeCells = new MergeCells();
            worksheet.InsertWorksheetElementInOrder(mergeCells);
        }

        // Create the merged cell and append it to the MergeCells collection.
        MergeCell mergeCell = new MergeCell
        {
            Reference = new StringValue($"{merge.CellReference.Address}:{merge.ToCellReference.Address}")
        };

        mergeCells.Append(mergeCell);
    }

    internal static void AutoSizeCells(this WorksheetPart worksheetPart, Dictionary<uint, double> maxColumnWidths)
    {
        var columns = new Columns();

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

        // TODO: Check if element already exists
        worksheetPart.Worksheet.InsertWorksheetElementInOrder(columns);
    }

    private static Cell GetCellInWorksheet(WorksheetPart worksheetPart, string columnName, uint rowIndex)
    {
        DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;

        if (sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.
        if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell? refCell = null;

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell
            {
                CellReference = cellReference
            };

            row.InsertBefore(newCell, refCell);

            return newCell;
        }
    }
}
