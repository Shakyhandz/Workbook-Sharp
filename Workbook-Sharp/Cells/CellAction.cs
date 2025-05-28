using DocumentFormat.OpenXml.Packaging;
using WorkbookSharp.Styles;

namespace WorkbookSharp.Cells;

internal abstract class CellAction
{
    private CellReference _cellReference;
    internal CellReference CellReference => _cellReference;
    internal uint? StyleIndex { get; set; }

    internal CellAction(string address, uint? styleIndex)
    {
        _cellReference = new CellReference(address);
        StyleIndex = styleIndex;
    }

    internal CellAction((uint row, uint column) cell, uint? styleIndex)
    {
        _cellReference = new CellReference(cell);
        StyleIndex = styleIndex;
    }

    internal virtual (uint startRow, uint startCol, uint endRow, uint endCol) GetKey()
    {
        return (CellReference.RowIndex, CellReference.ColumnIndex, CellReference.RowIndex, CellReference.ColumnIndex);
    }

    internal abstract void AddToWorksheetPart(WorksheetPart worksheetPart, SpreadsheetDocument document);

    internal double? EstimateColumnWidth(double? sheetFontSize, XlFontFamily sheetFontFamily, Styles.Style? style)
    {
        if (this is CellObject obj)
        {
            var text = obj?.Value?.ToString() ?? "";

            if (obj?.Value is DateTime dt)
            {
                if (style?.DateFormat == null || style.DateFormat == XlDateFormat.Date)
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

            return EstimateColumnWidth([(text, style)], sheetFontSize, sheetFontFamily);
        }
        else if (this is CellRichText rt)
        {
            return EstimateColumnWidth(rt.Runs, sheetFontSize, sheetFontFamily);
        }

        return null;
    }

    private double EstimateColumnWidth(List<(string text, Styles.Style? style)> runs, double? sheetFontSize, XlFontFamily sheetFontFamily)
    {
        double estimate = 0;

        foreach (var run in runs)
        {
            double total = 0;

            foreach (char c in run.text)
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
            var fontSize = run.style?.FontSize ?? sheetFontSize ?? 11.0;
            var fontSizeScale = fontSize / 11.0;
            
            var fontFamilyScale = (run.style?.FontFamily ?? sheetFontFamily) switch
            {
                XlFontFamily.Calibri => 1.0,
                XlFontFamily.Arial => 1.08,  // Arial is ~8% wider than Calibri
                _ => 1.0
            };

            estimate += total * fontSizeScale * fontFamilyScale;
        }

        // 255 is max allowed column width (+2 for padding)
        return Math.Min(255, estimate + 2);
    }
}
