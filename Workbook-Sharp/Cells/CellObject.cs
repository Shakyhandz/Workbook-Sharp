namespace WorkbookSharp.Cells;

internal class CellObject : CellAction
{
    private object _value;
    public object Value => _value;

    public CellObject(string address, object value, uint? styleIndex) : base(address, styleIndex)
    {
        _value = value;
    }

    public CellObject((uint row, uint column) cellReference, object value, uint? styleIndex) : base(cellReference, styleIndex)
    {
        _value = value;
    }
}
