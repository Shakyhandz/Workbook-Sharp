using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace WorkbookSharp.Cells;

internal class CellFormula : CellAction
{
    // Placed here to be able to reuse the same compiled instance
    private static readonly Regex R1C1Regex = new Regex(@"R(?<row>(\[\-?\d+\])|\d+)?C(?<col>(\[\-?\d+\])|\d+)?", RegexOptions.Compiled);

    private string? _formula;
    internal string? Formula => _formula;
    
    private bool _isRelative;
    internal bool IsRelative => _isRelative;
    
    internal CellFormula(string address, string? formula, bool isRelative, uint? styleIndex) : base(address, styleIndex)
    {
        _formula = formula?.TrimStart('='); // Remove leading '='
        _isRelative = isRelative;
    }

    internal CellFormula((uint row, uint column) cellReference, string? formula, bool isRelative, uint? styleIndex) : base(cellReference, styleIndex)
    {
        _formula = formula?.TrimStart('='); // Remove leading '='
        _isRelative = isRelative;
    }

    internal override void AddToWorksheetPart(WorksheetPart worksheetPart, SpreadsheetDocument document)
    {
        var cell = worksheetPart.GetOrInsertCellInWorksheet(CellReference);

        cell.CellReference = CellReference.Address; // optional but helps with structure
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula { Text = ParseFormula() };
        cell.StyleIndex = StyleIndex;
    }

    internal string ParseFormula()
    {
        var parsedFormula = ParseFormulaInternal();
        var errors = ExcelFormulaValidator.ValidateFormula(parsedFormula);

        if (errors.Count > 0)
            throw new ArgumentException($"Invalid formula {Formula} in cell {CellReference.Address}:\r\n{errors.StringJoin("\r\n")}");

        return parsedFormula!;
    }

    private string? ParseFormulaInternal()
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

        return Formula;
    }
}
