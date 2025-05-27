using WorkbookSharp.Styles;

namespace WorkbookSharp;

public class SpreadsheetExportOptions
{
    public string SheetName { get; set; } = "";
    public uint HeaderRow { get; set; } = 1;
    public uint StartColumn { get; set; } = 1;
    public Style? HeaderStyle { get; set; }
    public Style? ItemsStyle { get; set; }
    public Dictionary<uint, Style>? ColumnStyles { get; set; }
    public XlFontFamily? FontFamily { get; set; }
    public double? FontSize { get; set; }
}