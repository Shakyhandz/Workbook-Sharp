using WorkbookSharp.Styles;

namespace WorkbookSharp.Cells;

internal class CellAction
{
    private CellReference _cellReference;
    public CellReference CellReference => _cellReference;
    public uint? StyleIndex { get; set; }
    //public Style? Style { get; set; }

    //public CellAction(string address, Style? style)
    //{
    //    _cellReference = new CellReference(address);
    //    Style = style;
    //}

    //public CellAction((uint row, uint column) cell, Style? style)
    //{
    //    _cellReference = new CellReference(cell);
    //    Style = style;
    //}

    public CellAction(string address, uint? styleIndex)
    {
        _cellReference = new CellReference(address);
        StyleIndex = styleIndex;
    }

    public CellAction((uint row, uint column) cell, uint? styleIndex)
    {
        _cellReference = new CellReference(cell);
        StyleIndex = styleIndex;
    }

    public virtual (uint startRow, uint startCol, uint endRow, uint endCol) GetKey()
    {
        return (CellReference.RowIndex, CellReference.ColumnIndex, CellReference.RowIndex, CellReference.ColumnIndex);
    }
}
