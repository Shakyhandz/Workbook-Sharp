using WorkbookSharp.Cells;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

public interface IWorksheet
{
    XlFontFamily FontFamily { get; set; }
    double? FontSize { get; set; }
    bool AutoFitColumns { get; set; }
    bool ShowGridlines { get; set; }
    ICellRange Cells { get; }
    ICellRange Dimension { get; }
    void SetValue(string cellReference, object? value, Style? style = null);
    void SetValue((uint row, uint col) cellReference, object? value, Style? style = null);
    void SetRichText(string cellReference, params (string text, Style style)[] runs);
    void SetRichText((uint row, uint col) cellReference, params (string text, Style style)[] runs);
    void SetFormula(string cellReference, string? formula, bool isRelative, Style? style = null);
    void SetFormula((uint row, uint col) cellReference, string? formula, bool isRelative, Style? style = null);
    void MergeCells(string startCell, string endCell, Style? style = null);
    void MergeCells((uint row, uint col) startCell, (uint row, uint col) endCell, Style? style = null);
    void SetStyle(string cellReference, Style style);
    void SetStyle((uint row, uint col) cellReference, Style? style);
    void SetStyle((uint row, uint col) startCell, (uint row, uint col) endCell, Style? style);
}