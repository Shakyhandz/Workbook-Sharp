using System.Text.RegularExpressions;

namespace WorkbookSharp.Cells;

internal class CellReference
{
    private static readonly Regex _regexRow = new Regex(@"\d+", RegexOptions.Compiled);
    private static readonly Regex _regexColumn = new Regex(@"[A-Za-z]+", RegexOptions.Compiled);

    public const uint MAX_COLUMN = 16384;
    public const uint MAX_ROW = 1048576;

    public uint ColumnIndex { get; set; }
    public uint RowIndex { get; set; }
    public string ColumnName { get; set; } = "";
    public string Address => $"{ColumnName}{RowIndex}";

    internal CellReference(string address)
    {
        (ColumnName, RowIndex, ColumnIndex) = ParseAddress(address);

        ValidateCell(RowIndex, ColumnIndex);
    }

    internal CellReference((uint row, uint col) cell)
    {
        RowIndex = cell.row;
        ColumnIndex = cell.col;
        ColumnName = GetColumnName(cell.col);

        ValidateCell(RowIndex, ColumnIndex);
    }

    internal static (string columnName, uint rowIndex, uint columnIndex) ParseAddress(string address)
    {
        Match matchRow = _regexRow.Match(address);
        var rowIndex = uint.Parse(matchRow.Value);

        Match matchColumn = _regexColumn.Match(address);
        var columnName = matchColumn.Value.ToUpper();

        int columnIndex = 0;

        foreach (char c in columnName)
        {
            columnIndex *= 26;
            columnIndex += (c - 'A' + 1);
        }

        return (columnName, rowIndex, (uint)columnIndex);
    }

    internal static string GetColumnName(uint columnIndex)
    {
        string columnName = "";

        while (columnIndex > 0)
        {
            columnIndex--; // Adjust for 0-based indexing
            columnName = (char)('A' + (columnIndex % 26)) + columnName;
            columnIndex /= 26;
        }

        return columnName;
    }

    internal static List<CellReference> GetAllCells(CellReference from, CellReference to)
    {
        var res = new List<CellReference>();

        for (uint row = from.RowIndex; row <= to.RowIndex; row++)
        {
            for (uint col = from.ColumnIndex; col <= to.ColumnIndex; col++)
            {
                res.Add(new CellReference((row, col)));
            }
        }

        return res;
    }

    internal static void ValidateCell(uint row, uint col)
    {
        if (row < 1 || row > MAX_ROW)
            throw new ArgumentOutOfRangeException($"Row {row} is out of range. Valid range is 1 to {MAX_ROW}.");

        if (col < 1 || col > MAX_COLUMN)
            throw new ArgumentOutOfRangeException($"Column {col} is out of range. Valid range is 1 to {MAX_COLUMN}.");
    }
}
