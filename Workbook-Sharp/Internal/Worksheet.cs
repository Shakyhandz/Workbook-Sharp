using WorkbookSharp.Cells;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

internal class Worksheet : IWorksheet
{
    internal Workbook _workbook;    
    private string _sheetName;
    private readonly CellRange _cells;

    internal string SheetName => _sheetName;
    internal Dictionary<(uint startRow, uint startCol, uint endRow, uint endCol), CellAction> Actions { get; set; } = [];
    internal Dictionary<uint, double> MaxColumnWidths = [];

    public XlFontFamily FontFamily { get; set; } = XlFontFamily.Default;
    public double? FontSize { get; set; }
    public bool AutoFitColumns { get; set; } = true;
    public bool ShowGridlines { get; set; } = true;

    internal Worksheet(Workbook workbook, string name)
    {
        _cells = new CellRange(this);
        _workbook = workbook;
        _sheetName = name;        
    }

    internal CellRange Cells => _cells; // internal for internal code
    ICellRange IWorksheet.Cells => _cells; // public interface exposure
    public ICellRange Dimension => Actions.Count == 0
                                   ? Cells[1, 1, 1, 1] // default to A1 if no cells are touched
                                   : Cells[Actions.Keys.Min(k => k.startRow), Actions.Keys.Min(k => k.startCol), Actions.Keys.Max(k => k.endRow), Actions.Keys.Max(k => k.endCol)];

    // TODO: add bool "KeepExistingStyle" - at the moment it's kept as default?
    public void SetValue(string cellReference, object? value, Style? style = null) => 
        AddCellObject(new CellObject(cellReference, value, GetStyleIndex(style, value)), style);
    public void SetValue((uint row, uint col) cellReference, object? value, Style? style = null) => 
        AddCellObject(new CellObject(cellReference, value, GetStyleIndex(style, value)), style);

    private void AddCellObject(CellObject obj, Style? style)
    {
        UpdateMaxColumnWidths(obj, style);

        if (Actions.TryGetValue(obj.GetKey(), out var action))
        {
            // Keep style of existing cell if not set here
            if (style == null)
                obj.StyleIndex = action.StyleIndex;

            Actions[obj.GetKey()] = obj;
        }
        else
        {
            Actions.Add(obj.GetKey(), obj);
        }
    }

    private void UpdateMaxColumnWidths(CellAction action, Style? style)
    {
        if (AutoFitColumns)
        {
            var col = action.CellReference.ColumnIndex;
            var estimatedWidth = action.EstimateColumnWidth(FontSize, FontFamily, style);

            if (estimatedWidth != null)
            {
                if (MaxColumnWidths.TryGetValue(col, out var currentMax))
                {
                    if (estimatedWidth > currentMax)
                        MaxColumnWidths[col] = estimatedWidth.Value;
                }
                else
                {
                    MaxColumnWidths[col] = estimatedWidth.Value;
                }
            }
        }
    }

    public void SetRichText(string cellReference, params (string text, Style style)[] runs) =>
        AddRichText(new CellRichText(cellReference, runs));
    public void SetRichText((uint row, uint col) cellReference, params (string text, Style style)[] runs) =>
        AddRichText(new CellRichText(cellReference, runs));

    private void AddRichText(CellRichText obj)
    {
        UpdateMaxColumnWidths(obj, null);

        if (Actions.TryGetValue(obj.GetKey(), out var action))
        {
            // Keep style of existing cell
            obj.StyleIndex = action.StyleIndex;
            Actions[obj.GetKey()] = obj;
        }
        else
        {
            Actions.Add(obj.GetKey(), obj);
        }
    }

    public void SetFormula(string cellReference, string? formula, bool isRelative, Style? style = null) =>
        AddFormula(new CellFormula(cellReference, formula, isRelative, GetStyleIndex(style)));
    public void SetFormula((uint row, uint col) cellReference, string? formula, bool isRelative, Style? style = null) =>
        AddFormula(new CellFormula(cellReference, formula, isRelative, GetStyleIndex(style)));

    private void AddFormula(CellFormula obj)
    {
        if (Actions.ContainsKey(obj.GetKey()))
            Actions[obj.GetKey()] = obj;
        else
            Actions.Add(obj.GetKey(), obj);
    }

    public void MergeCells(string startCell, string endCell, Style? style = null) => AddCellMerge(new CellMerge(startCell, endCell, GetStyleIndex(style)));
    public void MergeCells((uint row, uint col) startCell, (uint row, uint col) endCell, Style? style = null) => AddCellMerge(new CellMerge(startCell, endCell, GetStyleIndex(style)));

    private void AddCellMerge(CellMerge obj)
    {
        // TODO: validate order of cells?

        if (Actions.ContainsKey(obj.GetKey()))
            Actions[obj.GetKey()] = obj;
        else
            Actions.Add(obj.GetKey(), obj);

        // Set style to all cells
        foreach (var cell in CellReference.GetAllCells(obj.CellReference, obj.ToCellReference))
        {
            AddStyle(new CellStyle(cell.Address, obj.StyleIndex));
        }        
    }

    internal void UnMergeCells(CellMerge obj)
    {
        // Remove merge action
        if (Actions.ContainsKey(obj.GetKey()))
            Actions.Remove(obj.GetKey());
    }

    public void SetStyle(string cellReference, Style style) => AddStyle(new CellStyle(cellReference, GetStyleIndex(style)));
    public void SetStyle((uint row, uint col) cellReference, Style? style) => AddStyle(new CellStyle(cellReference, GetStyleIndex(style)));
    
    public void SetStyle((uint row, uint col) startCell, (uint row, uint col) endCell, Style? style)
    {
        var from = new CellReference(startCell);
        var to = new CellReference(endCell);

        // Set style to all cells
        foreach (var cell in CellReference.GetAllCells(from, to))
        {
            AddStyle(new CellStyle(cell.Address, GetStyleIndex(style)));
        }
    }

    private void AddStyle(CellStyle style)
    {
        if (Actions.TryGetValue(style.GetKey(), out var a))
            a.StyleIndex = style.StyleIndex;
        else
            Actions.Add(style.GetKey(), style);
    }

    private uint? GetStyleIndex(Style? style, object? value = null)
    {
        // Default to worksheet font
        style ??= new Style
        {
            FontFamily = FontFamily,
            FontSize = FontSize,
        };

        // If font not defined for cell, set it to worksheet font
        if (style.FontFamily == XlFontFamily.Default)
            style.FontFamily = FontFamily;

        if (style.FontSize == null)
            style.FontSize = FontSize;

        // Default to date format if value is DateTime
        if (value != null && value is DateTime dt && style.DateFormat == null)
        {
            style.DateFormat = dt.Second > 0 || dt.Minute > 0 || dt.Hour > 0 
                               ? XlDateFormat.DateHoursMinutesSeconds 
                               : XlDateFormat.Date;
        }

        return _workbook.styleManager.GetStyleIndex(style);
    }

}