using WorkbookSharp.Cells;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

internal class Worksheet : IWorksheet
{
    private Workbook _workbook;    
    private string _sheetName;
    
    internal string SheetName => _sheetName;
    internal Dictionary<(uint startRow, uint startCol, uint endRow, uint endCol), CellAction> Actions { get; set; } = [];
    internal Dictionary<uint, double> MaxColumnWidths = new();

    public XlFontFamily FontFamily { get; set; } = XlFontFamily.Default;
    public double? FontSize { get; set; }
    public bool AutoFitColumns { get; set; } = true;
    public bool ShowGridlines { get; set; } = true;

    internal Worksheet(Workbook workbook, string name)
    {
        _workbook = workbook;
        _sheetName = name;        
    }

    public CellRange Cells => new CellRange(this);

    public void SetValue(string cellReference, object? value, Style? style = null) => 
        AddCellObject(new CellObject(cellReference, value, GetStyleIndex(style, value)), style);
    public void SetValue((uint row, uint col) cellReference, object? value, Style? style = null) => 
        AddCellObject(new CellObject(cellReference, value, GetStyleIndex(style, value)), style);

    private void AddCellObject(CellObject obj, Style? style)
    {
        if (AutoFitColumns && obj.Value != null)
        {
            var text = obj.Value.ToString() ?? "";

            if (obj.Value is DateTime dt)
            {
                if (style?.DateFormat == null  || style.DateFormat == XlDateFormat.Date)
                    text = dt.ToShortDateString();
                else if (style.DateFormat == XlDateFormat.DateHoursMinutesSeconds)
                    text = dt.ToString("yyyy-MM-dd HH:mm:ss");
                else if (style.DateFormat == XlDateFormat.DateHoursMinutes)
                    text = dt.ToString("yyyy-MM-dd HH:mm");
                else if (style.DateFormat == XlDateFormat.DateHours)
                    text = dt.ToString("yyyy-MM-dd HH");
                else if (style.DateFormat == XlDateFormat.HoursMinutesSeconds)
                    text = dt.ToString("HH:mm:ss");
                else if (style.DateFormat == XlDateFormat.HoursMinutes)
                    text = dt.ToString("HH:mm");
            }
            double estimatedWidth = EstimateColumnWidth(text, style); 
            uint col = obj.CellReference.ColumnIndex;
            
            if (MaxColumnWidths.TryGetValue(col, out var currentMax))
            {
                if (estimatedWidth > currentMax)
                    MaxColumnWidths[col] = estimatedWidth;
            }
            else
            {
                MaxColumnWidths[col] = estimatedWidth;
            }
        }

        if (Actions.ContainsKey(obj.GetKey()))
            Actions[obj.GetKey()] = obj;
        else
            Actions.Add(obj.GetKey(), obj);
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

    private double EstimateColumnWidth(string text, Style? cellStyle)
    {
        if (text.IsNothing())
            return 0;

        // Average character widths, roughly based on Calibri 11pt
        double total = 0;

        foreach (char c in text)
        {
            total += c switch
            {
                'i' or 'l' or 'I' or '|' => 0.5,
                'W' or 'M' => 1.5,
                ' ' => 0.5,
                '-' or '_' => 0.75,
                _ => 1.0
            };
        }

        // Scale for font size (default width scale is based on 11pt)
        var fontSize = cellStyle?.FontSize ?? FontSize ?? 11.0;
        var fontSizeScale = fontSize / 11.0;
        var fontFamilyScale = (cellStyle?.FontFamily ?? FontFamily) switch
        {
            XlFontFamily.Calibri => 1.0,
            XlFontFamily.Arial => 1.08,  // Arial is ~8% wider than Calibri
            _ => 1.0
        };

        return Math.Min(255, total * fontSizeScale * fontFamilyScale + 2); // +2 for padding
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