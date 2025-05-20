using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using WorkbookSharp.Cells;
using WorkbookSharp.Styles;

namespace WorkbookSharp;

public class Workbook
{
    // Placed here to be able to reuse the same compiled instance
    private static readonly Regex R1C1Regex = new Regex(@"R(?<row>(\[\-?\d+\])|\d+)?C(?<col>(\[\-?\d+\])|\d+)?", RegexOptions.Compiled);

    private List<Worksheet> _worksheets = [];
    internal StyleManager styleManager = new();
    
    public Worksheet AddWorksheet(string name = "")
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
                    SetValue(document, worksheetPart, xlObj.CellReference, xlObj.Value, xlObj.StyleIndex);
                }
                else if (action is Cells.CellFormula formula)
                {
                    if (formula.Formula == null)
                    {
                        SetValue(document, worksheetPart, formula.CellReference, formula.Formula, formula.StyleIndex);
                    }
                    else
                    {
                        var cell = InsertCellInWorksheet(worksheetPart, formula.CellReference.ColumnName, formula.CellReference.RowIndex);

                        cell.CellReference = formula.CellReference.Address; // optional but helps with structure
                        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula { Text = formula.ParseFormula(R1C1Regex) };
                        cell.StyleIndex = formula.StyleIndex;                        
                    }
                }
                else if (action is Cells.CellStyle xlStyle)
                {
                    // Set the style for the cell
                    var cell = InsertCellInWorksheet(worksheetPart, xlStyle.CellReference.ColumnName, xlStyle.CellReference.RowIndex);
                    cell.StyleIndex = xlStyle.StyleIndex;
                }
                else if (action is CellMerge xlMerge)
                {
                    MergeCellsInternal(worksheetPart, xlMerge.CellReference, xlMerge.ToCellReference);
                }
            }

            if (ws.AutoFitColumns)
            {
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var columns = AutoSizeCells(sheetData!, ws.MaxColumnWidths); // Is not null, added above in worksheet constructor
                InsertWorksheetElementInOrder(worksheetPart.Worksheet, columns);
            }
        }

        document.Save();
        return stream.ToArray();
    }

    //private uint? GetStyleIndex(XlFontFamily fontFamily, double? fontSize, Style? style, object? value = null)
    //{
    //    // Default to worksheet font
    //    style ??= new Style
    //    {
    //        FontFamily = fontFamily,
    //        FontSize = fontSize,
    //    };

    //    // Default to date format if value is DateTime
    //    if (value != null && value is DateTime && style.DateFormat == null)
    //        style.DateFormat = XlDateFormat.Date;

    //    return styleManager.GetStyleIndex(style);
    //}

    private Columns AutoSizeCells(SheetData sheetData, Dictionary<uint, double> maxColumnWidths)
    {
        var columns = new Columns();

        foreach (var kvp in maxColumnWidths.OrderBy(x => x.Key))
        {
            columns.Append(new Column
            {
                Min = kvp.Key,
                Max = kvp.Key,
                Width = kvp.Value,
                CustomWidth = true
            });
        }

        return columns;
    }

    private void SetValue(SpreadsheetDocument document, WorksheetPart worksheetPart, CellReference resultCell, object? value, uint? styleIndex)
    {
        CellValue? cellValue = null;
        EnumValue<CellValues>? dataType = null;

        // Set the value of the cell
        if (value is int i)
        {
            cellValue = new CellValue(i);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (value is decimal dec)
        {
            cellValue = new CellValue(dec);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (value is double d)
        {
            cellValue = new CellValue(d);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (value is long l)
        {
            cellValue = new CellValue((decimal)l);
            dataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (value is bool b)
        {
            cellValue = new CellValue(b);
            dataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (value is DateTime dt)
        {
            cellValue = new CellValue(dt);
            dataType = new EnumValue<CellValues>(CellValues.Date);
        }

        // Insert the cell value
        if (cellValue != null && dataType != null)
        {
            AddCellValue(worksheetPart, resultCell, cellValue, dataType, styleIndex);
        }
        else 
        {
            // TODO: Add null value if null?
            InsertText(document, worksheetPart, resultCell, value?.ToString() ?? "", styleIndex);
        }
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

    /* TODO: show gridlines
     var sheetViews = new SheetViews(
        new SheetView
        {
            WorkbookViewId = 0,
            ShowGridLines = false
        }
    );
    
    worksheet.Append(sheetViews);
    */

    private void InsertText(SpreadsheetDocument document, WorksheetPart worksheetPart, CellReference cellReference, string text, uint? styleIndex)
    {
        // Get the SharedStringTablePart and add the result to it
        var shareStringPart = document.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().Count() > 0
                              ? document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First()
                              : document.WorkbookPart!.AddNewPart<SharedStringTablePart>();

        // Insert the result into the SharedStringTablePart
        int index = InsertSharedStringItem(text, shareStringPart);

        var cellValue = new CellValue(index.ToString());
        var dataType = new EnumValue<CellValues>(CellValues.SharedString);
        AddCellValue(worksheetPart, cellReference, cellValue, dataType, styleIndex);
    }

    private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create it.
        if (shareStringPart.SharedStringTable is null)
            shareStringPart.SharedStringTable = new SharedStringTable();

        int index = 0;

        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)            
                return index; // The text already exists in the part. Return its index

            index++;
        }

        // The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

        return index;
    }

    private void AddCellValue(WorksheetPart worksheetPart, CellReference cellReference, CellValue? cellValue, EnumValue<CellValues>? dataType, uint? styleIndex = null)
    {
        Cell result = InsertCellInWorksheet(worksheetPart, cellReference.ColumnName, cellReference.RowIndex);

        // Set the value of the cell
        result.CellValue = cellValue;
        result.DataType = dataType;

        if (styleIndex != null)
            result.StyleIndex = styleIndex;
    }

    private void MergeCellsInternal(WorksheetPart worksheetPart, CellReference startCell, CellReference endCell)
    {
        // Verify if the specified cells exist, and if they do not exist, create them.
        InsertCellInWorksheet(worksheetPart, startCell.ColumnName, startCell.RowIndex);
        InsertCellInWorksheet(worksheetPart, endCell.ColumnName, endCell.RowIndex);

        MergeCells mergeCells;
        var worksheet = worksheetPart.Worksheet;

        if (worksheet.Elements<MergeCells>().Count() > 0)
        {
            mergeCells = worksheet.Elements<MergeCells>().First();
        }
        else
        {
            mergeCells = new MergeCells();
            InsertWorksheetElementInOrder(worksheet, mergeCells);
        }

        // Create the merged cell and append it to the MergeCells collection.
        MergeCell mergeCell = new MergeCell
        {
            Reference = new StringValue($"{startCell.Address}:{endCell.Address}")
        };

        mergeCells.Append(mergeCell);
    }

    // Correct Order of Elements in a Worksheet
    private static readonly List<Type> _orderOfElements = new()
    {
        typeof(SheetProperties),
        typeof(SheetDimension),
        typeof(SheetViews),
        typeof(SheetFormatProperties),
        typeof(Columns),
        typeof(SheetData),
        typeof(SheetProtection),
        typeof(ProtectedRanges),
        typeof(Scenarios),
        typeof(AutoFilter),
        typeof(SortState),
        typeof(DataConsolidate),
        typeof(CustomSheetViews),
        typeof(MergeCells),
        typeof(PhoneticProperties),
        typeof(ConditionalFormatting),
        typeof(DataValidations),
        typeof(Hyperlinks),
        typeof(PrintOptions),
        typeof(PageMargins),
        typeof(PageSetup),
        typeof(HeaderFooter),
        typeof(RowBreaks),
        typeof(ColumnBreaks),
        typeof(CustomProperties),
        typeof(CellWatches),
        typeof(IgnoredErrors),
        //typeof(SmartTags),
        typeof(Drawing),
        typeof(LegacyDrawing),
        typeof(LegacyDrawingHeaderFooter),
        typeof(Picture),
        typeof(OleObjects),
        typeof(Controls),
        typeof(WebPublishItems),
        typeof(TableParts),
        typeof(WorksheetExtensionList)
    };

    private static void InsertWorksheetElementInOrder(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, OpenXmlElement newElement)
    {
        var newElementType = newElement.GetType();
        int newIndex = _orderOfElements.IndexOf(newElementType);

        if (newIndex == -1)
            throw new InvalidOperationException("Element type is not valid for Worksheet children.");

        // Find the first existing child that comes after the new element
        foreach (var child in worksheet.Elements())
        {
            int childIndex = _orderOfElements.IndexOf(child.GetType());

            if (childIndex > newIndex)
            {
                worksheet.InsertBefore(newElement, child);
                return;
            }
        }

        // If no such child exists, append to the end
        worksheet.Append(newElement);
    }

    private Cell InsertCellInWorksheet(WorksheetPart worksheetPart, string columnName, uint rowIndex)
    {
        DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;

        if (sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex is not null && r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.
        if (row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference is not null && c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell? refCell = null;

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference?.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell
            {
                CellReference = cellReference
            };

            row.InsertBefore(newCell, refCell);

            return newCell;
        }
    }
}
