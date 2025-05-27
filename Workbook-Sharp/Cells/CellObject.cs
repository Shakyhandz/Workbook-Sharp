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
}
