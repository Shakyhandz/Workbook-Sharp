using WorkbookSharp.Styles;

namespace WorkbookSharp;

public interface ICellRange
{
    object? Value { get; set; }
    string? Formula { get; set; }
    bool Merge { set; }
    Style? Style { set; }
    ICellRange this[uint row, uint col] { get; }
    ICellRange this[uint fromRow, uint fromCol, uint toRow, uint toCol] { get; }
    ICellRange this[string address] { get; }
}
