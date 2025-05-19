namespace WorkbookSharp.Cells;

internal class CellStyle : CellAction
{
    public CellStyle(string address, uint? styleIndex) : base(address, styleIndex)
    {
    }

    public CellStyle((uint row, uint column) cellReference, uint? styleIndex) : base(cellReference, styleIndex)
    {
    }
}
