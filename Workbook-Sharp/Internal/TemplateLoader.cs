using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

internal class TemplateLoader : ITemplateLoader
{
    public IEnumerable<string> GetEmbeddedExcelTemplates(Assembly assembly)
    {
        var resourceNames = assembly.GetManifestResourceNames();

        foreach (var name in resourceNames)
        {
            yield return name;
        }
    }

    public IWorkbook LoadWorkbookFromTemplate(Assembly assembly, string resourcePath)
    {
        using Stream? resourceStream = assembly.GetManifestResourceStream(resourcePath) ?? throw new FileNotFoundException($"Embedded resource not found: {resourcePath}");

        using var memStream = new MemoryStream();
        resourceStream.CopyTo(memStream);
        memStream.Position = 0;

        using var document = SpreadsheetDocument.Open(memStream, false);
        var workbookPart = document.WorkbookPart!;
        var styleSheet = workbookPart.WorkbookStylesPart?.Stylesheet;

        var newWorkbook = new Workbook();

        foreach (var sheet in workbookPart?.Workbook?.Sheets?.OfType<Sheet>() ?? [])
        {
            var worksheetPart = (WorksheetPart)workbookPart!.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            if ((sheet?.Name?.Value).IsNothing())
                throw new InvalidOperationException("Sheet name cannot be null or empty.");

            var newWorksheet = new Worksheet(newWorkbook, sheet!.Name!);

            foreach (var row in sheetData?.Elements<Row>() ?? [])
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value;
                    
                    if (cellRef == null) 
                        continue;

                    Style? style = null;
                    
                    if (cell.StyleIndex != null && styleSheet != null)
                        style = styleSheet.ConvertToStyle(cell.StyleIndex.Value, workbookPart);

                    if (cell.CellFormula != null)
                    {
                        string formula = cell.CellFormula.Text ?? string.Empty;
                        bool isRelative = cell.CellFormula.Reference == null;
                        newWorksheet.SetFormula(cellRef, formula, isRelative, style);
                    }
                    else
                    {
                        var value = cell.GetValue(workbookPart);
                        newWorksheet.SetValue(cellRef, value, style);
                    }
                }
            }

            // Handle merged cells
            var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();

            if (mergeCells != null)
            {
                foreach (var mergeCell in mergeCells.Elements<MergeCell>())
                {
                    var range = mergeCell.Reference?.Value;
            
                    if (!string.IsNullOrWhiteSpace(range) && range.Contains(":"))
                    {
                        var parts = range.Split(':');
                        newWorksheet.MergeCells(parts[0], parts[1]);
                    }
                }
            }

            newWorkbook.AddExistingWorksheet(newWorksheet);
        }

        return newWorkbook;
    }
}

public static class OpenXmlExtensions
{
    public static object? GetValue(this Cell cell, WorkbookPart workbookPart)
    {
        var raw = cell.CellValue?.Text;

        if (string.IsNullOrWhiteSpace(raw))
            return null;

        var styleSheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        uint? formatId = null;

        if (styleSheet != null && styleSheet.CellFormats != null && cell.StyleIndex != null)
        {
            var styleIndex = (int)cell.StyleIndex.Value;

            if (styleIndex < (styleSheet?.CellFormats?.Count ?? 0))
            {
                var cellFormat = styleSheet!.CellFormats.ElementAt(styleIndex) as CellFormat;
                formatId = cellFormat?.NumberFormatId?.Value;
            }
        }

        if (14 <= formatId && formatId <= 22 && double.TryParse(raw, out var oaDate))
            return DateTime.FromOADate(oaDate);

        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.SharedString)
        {
            int sstIndex = int.TryParse(raw, out var index) ? index : -1;

            return sstIndex >= 0
                   ? workbookPart.SharedStringTablePart?.SharedStringTable.ElementAt(index)?.InnerText
                   : null;
        }
        else if (dataType == CellValues.Boolean)
        {
            return raw == "1" || (raw.IsSome() && raw.Equals("true", StringComparison.OrdinalIgnoreCase));
        }
        else if (dataType == CellValues.Date)
        {
            return double.TryParse(raw, out oaDate)
                   ? DateTime.FromOADate(oaDate) 
                   : raw;
        }
        else if (dataType == CellValues.Number)
        {
            return double.TryParse(raw, out var d) ? d : raw;
        }
        
        return raw;
    }

    public static Style ConvertToStyle(this Stylesheet stylesheet, uint styleIndex, WorkbookPart workbookPart)
    {
        var cellFormat = stylesheet.CellFormats?.ElementAt((int)styleIndex) as CellFormat;
        if (cellFormat == null)
            return new Style();

        var style = new Style();

        // Extract font info
        if (cellFormat.FontId != null && stylesheet.Fonts != null)
        {
            var font = stylesheet.Fonts.ElementAt((int)cellFormat.FontId.Value) as Font;

            if (font != null)
            {
                style.FontSize = double.TryParse(font.FontSize?.Val?.ToString(), out var sz) ? sz : null;

                var fontName = font.FontName?.Val?.Value;

                style.FontFamily = fontName switch
                {
                    "Arial" => XlFontFamily.Arial,
                    "Calibri" => XlFontFamily.Calibri,
                    _ => XlFontFamily.Default
                };
            }
        }

        // Extract number format
        if (cellFormat.NumberFormatId != null)
        {
            var nId = cellFormat.NumberFormatId.Value;

            style.DateFormat = nId switch
            {
                14 => XlDateFormat.Date,
                22 => XlDateFormat.DateHoursMinutesSeconds,
                165 => XlDateFormat.DateHours,
                _ => null
            };
        }

        return style;
    }
}
