using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp.Cells;

internal class CellObject : CellAction
{
    private object? _value;
    internal object? Value => _value;

    internal CellObject(string address, object? value, uint? styleIndex) : base(address, styleIndex)
    {
        _value = value;
    }

    internal CellObject((uint row, uint column) cellReference, object? value, uint? styleIndex) : base(cellReference, styleIndex)
    {
        _value = value;
    }

    internal override void AddToWorksheetPart(WorksheetPart worksheetPart, SpreadsheetDocument document)
    {
        CellValue? cellValue = null;
        EnumValue<CellValues>? dataType = null;

        // Set the value of the cell
        if (Value is int i)
        {
            cellValue = new CellValue(i);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (Value is decimal dec)
        {
            cellValue = new CellValue(dec);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (Value is double d)
        {
            cellValue = new CellValue(d);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (Value is long l)
        {
            cellValue = new CellValue((decimal)l);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (Value is bool b)
        {
            cellValue = new CellValue(b);
            dataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (Value is DateTime dt)
        {
            cellValue = new CellValue(dt);
            dataType = new EnumValue<CellValues>(CellValues.Date);
        }
        else
        {
            // TODO: Add null value if null?
            var index = document.GetSharedStringIndex(worksheetPart, Value?.ToString() ?? "");
            cellValue = new CellValue(index.ToString());
            dataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        // Insert the cell value        
        Cell cell = worksheetPart.GetOrInsertCellInWorksheet(CellReference);

        cell.CellValue = cellValue;
        cell.DataType = dataType;

        if (StyleIndex != null)
            cell.StyleIndex = StyleIndex;
    }
}
