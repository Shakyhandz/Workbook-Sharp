using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookSharp.Cells;

internal class CellMerge : CellAction
{
    private CellReference _toCellReference;
    internal CellReference ToCellReference => _toCellReference;

    internal CellMerge(string fromAddress, string toAddress, uint? styleIndex = null) : base(fromAddress, styleIndex)
    {
        _toCellReference = new CellReference(toAddress);
    }

    internal CellMerge((uint row, uint column) fromCell, (uint row, uint column) toCell, uint? styleIndex = null) : base(fromCell, styleIndex)
    {
        _toCellReference = new CellReference(toCell);
    }

    internal override (uint startRow, uint startCol, uint endRow, uint endCol) GetKey() => (CellReference.RowIndex, CellReference.ColumnIndex, ToCellReference.RowIndex, ToCellReference.ColumnIndex);

    internal override void AddToWorksheetPart(WorksheetPart worksheetPart, SpreadsheetDocument document)
    {
        worksheetPart.GetOrInsertCellInWorksheet(CellReference);
        worksheetPart.GetOrInsertCellInWorksheet(ToCellReference);

        var mergeCells = worksheetPart.Worksheet.GetOrInsertWorksheetElement<MergeCells>();

        // Create the merged cell and append it to the MergeCells collection.
        MergeCell mergeCell = new MergeCell
        {
            Reference = new StringValue($"{CellReference.Address}:{ToCellReference.Address}")
        };

        mergeCells.Append(mergeCell);
    }
}