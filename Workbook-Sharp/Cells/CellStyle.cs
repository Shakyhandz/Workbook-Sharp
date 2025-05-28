using DocumentFormat.OpenXml.Packaging;

namespace WorkbookSharp.Cells;

internal class CellStyle : CellAction
{
    internal CellStyle(string address, uint? styleIndex) : base(address, styleIndex)
    {
    }

    internal CellStyle((uint row, uint column) cellReference, uint? styleIndex) : base(cellReference, styleIndex)
    {
    }

    internal override void AddToWorksheetPart(WorksheetPart worksheetPart, SpreadsheetDocument document)
    {
        var cell = worksheetPart.GetOrInsertCellInWorksheet(CellReference);
        cell.StyleIndex = StyleIndex;
    }
}
