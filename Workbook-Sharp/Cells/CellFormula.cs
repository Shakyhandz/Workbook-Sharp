using System.Text.RegularExpressions;

namespace WorkbookSharp.Cells;

internal class CellFormula : CellAction
{
    // Placed here to be able to reuse the same compiled instance
    private static readonly Regex R1C1Regex = new Regex(@"R(?<row>(\[\-?\d+\])|\d+)?C(?<col>(\[\-?\d+\])|\d+)?", RegexOptions.Compiled);

    private string? _formula;
    public string? Formula => _formula;
    
    private bool _isRelative;
    public bool IsRelative => _isRelative;
    
    public CellFormula(string address, string? formula, bool isRelative, uint? styleIndex) : base(address, styleIndex)
    {
        _formula = formula?.TrimStart('='); // Remove leading '='
        _isRelative = isRelative;
    }

    public CellFormula((uint row, uint column) cellReference, string? formula, bool isRelative, uint? styleIndex) : base(cellReference, styleIndex)
    {
        _formula = formula?.TrimStart('='); // Remove leading '='
        _isRelative = isRelative;
    }

    internal string ParseFormula()
    {
        if (Formula.IsSome() && IsRelative)
        {
            try
            {

                return R1C1Regex.Replace(Formula ?? "", match =>
                {
                    string rowPart = match.Groups["row"].Value;
                    string colPart = match.Groups["col"].Value;

                    uint row = CellReference.RowIndex;
                    uint col = CellReference.ColumnIndex;

                    // Row handling
                    if (rowPart.IsSome())
                    {
                        rowPart = rowPart.Trim('[', ']');
                        row = (uint)(CellReference.RowIndex + int.Parse(rowPart));
                    }

                    // Column handling
                    if (!string.IsNullOrEmpty(colPart))
                    {
                        colPart = colPart.Trim('[', ']');
                        col = (uint)(CellReference.ColumnIndex + int.Parse(colPart));
                    }

                    return $"{CellReference.GetColumnName(col)}{row}";
                });
            }
            catch
            {
                throw new ArgumentException($"Could not parse relative formula {Formula} in cell {CellReference.Address}");
            }
        }
        else
        {
            return Formula ?? "";
        }
    }
}
