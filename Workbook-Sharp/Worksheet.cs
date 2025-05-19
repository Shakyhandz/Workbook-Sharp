using WorkbookSharp.Cells;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

public class Worksheet
{
    private Workbook _workbook;    
    private string _sheetName;
    internal string SheetName => _sheetName;
    //internal List<CellAction> Actions { get; set; } = [];
    internal Dictionary<(uint startRow, uint startCol, uint endRow, uint endCol), CellAction> Actions { get; set; } = [];
    public XlFontFamily FontFamily { get; set; } = XlFontFamily.Default;
    public double? FontSize { get; set; }
    /// <summary>
    /// Defaults to true
    /// </summary>
    public bool AutoFitColumns { get; set; } = true;

    internal Worksheet(Workbook workbook, string name)
    {
        _workbook = workbook;
        _sheetName = name;        
    }

    public CellRange Cells => new CellRange(this, 1, 1, CellReference.MAX_ROW, CellReference.MAX_COLUMN);

    public void SetValue(string cellReference, object value, Style? style = null) => AddCellObject(new CellObject(cellReference, value, GetStyleIndex(style, value)));
    public void SetValue((uint row, uint col) cellReference, object value, Style? style = null) => AddCellObject(new CellObject(cellReference, value, GetStyleIndex(style, value)));

    private void AddCellObject(CellObject obj)
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

    public void SetStyle(string cellReference, Style style) => AddStyle(new CellStyle(cellReference, GetStyleIndex(style)));
    public void SetStyle((uint row, uint col) cellReference, Style? style) => AddStyle(new CellStyle(cellReference, GetStyleIndex(style)));

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
        if (value != null && value is DateTime && style.DateFormat == null)
            style.DateFormat = XlDateFormat.Date;

        return _workbook.styleManager.GetStyleIndex(style);
    }

}