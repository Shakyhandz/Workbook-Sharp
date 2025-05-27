using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorkbookSharp.Cells;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

internal class Workbook : IWorkbook
{
    private List<Worksheet> _worksheets = [];
    internal StyleManager styleManager = new();
    
    public IWorksheet AddWorksheet(string name = "")
    {
        if (string.IsNullOrEmpty(name.Trim()))
            name = $"Sheet{_worksheets.Count + 1}";

        var newSheetName = _worksheets.GetNewSheetNameSafe(name);
        var sheet = new Worksheet(this, newSheetName);
        _worksheets.Add(sheet);

        return sheet;
    }

    public async Task Save(string fileName)
    {
        var bytes = await Save();
        File.WriteAllBytes(fileName, bytes);
    }

    public async Task<byte[]> Save()
    {
        await Task.Yield();

        using var stream = new MemoryStream();
        using SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, false);
        
        // Add a WorkbookPart to the document.
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

        // Add styles to the WorkbookPart
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = styleManager.BuildStylesheet();
        stylesPart.Stylesheet.Save();

        // Add Sheets to the Workbook.
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        uint sheetId = 1;

        foreach (var ws in _worksheets)
        {
            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());

            // Append a new worksheet and associate it with the workbook.
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = ws.SheetName,
            };

            sheets.Append(sheet);

            // Very important to set values in the correct order, row by row, column by column
            var actions = ws.Actions
                            .Values
                            .OrderBy(x => x.CellReference.RowIndex)
                            .ThenBy(x => x.CellReference.ColumnIndex)
                            .ToList();

            foreach (var action in actions)
            {
                if (action is CellObject xlObj)
                {
                    worksheetPart.SetCellValue(document, xlObj);
                }
                else if (action is Cells.CellFormula xlFormula)
                {
                    worksheetPart.SetCellFormula(xlFormula);
                }
                else if (action is Cells.CellStyle xlStyle)
                {
                    worksheetPart.SetCellStyle(xlStyle);
                }
                else if (action is CellMerge xlMerge)
                {
                    worksheetPart.MergeCells(xlMerge);
                }
            }

            if (ws.AutoFitColumns)
            {
                worksheetPart.AutoSizeCells(ws.MaxColumnWidths);
            }

            if (!ws.ShowGridlines)
            {
                worksheetPart.ShowGridlines(ws.ShowGridlines);
            }
        }

        document.Save();
        return stream.ToArray();
    }


    /* TODO: set default row height for one font size

    // Rough approximation based on Excel's behavior
    var rowHeight = fontSize * 1.33;

    // Insert in correct order
    InsertWorksheetElementInOrder 
    new SheetFormatProperties
    {
        DefaultRowHeight = rowHeight,
        CustomHeight = true
    });
    */


}
