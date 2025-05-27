using WorkbookSharp.Styles;

namespace WorkbookSharp.Cells;

internal class CellAction
{
    private CellReference _cellReference;
    internal CellReference CellReference => _cellReference;
    internal uint? StyleIndex { get; set; }
    //internal Style? Style { get; set; }

    //internal CellAction(string address, Style? style)
    //{
    //    _cellReference = new CellReference(address);
    //    Style = style;
    //}

    //internal CellAction((uint row, uint column) cell, Style? style)
    //{
    //    _cellReference = new CellReference(cell);
    //    Style = style;
    //}

    internal CellAction(string address, uint? styleIndex)
    {
        _cellReference = new CellReference(address);
        StyleIndex = styleIndex;
    }

    internal CellAction((uint row, uint column) cell, uint? styleIndex)
    {
        _cellReference = new CellReference(cell);
        StyleIndex = styleIndex;
    }

    internal virtual (uint startRow, uint startCol, uint endRow, uint endCol) GetKey()
    {
        return (CellReference.RowIndex, CellReference.ColumnIndex, CellReference.RowIndex, CellReference.ColumnIndex);
    }
}
