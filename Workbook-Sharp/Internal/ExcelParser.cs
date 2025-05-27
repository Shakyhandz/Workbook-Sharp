using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Dynamic;
using WorkbookSharp.Cells;

namespace WorkbookSharp;

internal class ExcelParser : IExcelParser, IDisposable
{
    public string SheetName { get; set; } = "";   
    public string FilePath { get; set; } = "";
    public uint HeaderRow { get; set; } = 1;    
    public uint HeaderStartColumn { get; set; } = 1;
    public int? HeaderLength { get; set; }
    public uint? LastRow { get; set; }    
    
    private Dictionary<string, int> mHeaders = [];

    public IEnumerable<dynamic> Execute()
    {
        Init();

        var fi = new FileInfo(FilePath);

        if (!fi.Exists)
            throw new ArgumentException($"File {FilePath} does not exist");

        using var spreadsheetDocument = SpreadsheetDocument.Open(FilePath, false);

        var workBookPart = spreadsheetDocument.WorkbookPart;
        DocumentFormat.OpenXml.Spreadsheet.Worksheet? worksheet = null;

        foreach (var sheet in workBookPart?.Workbook?.Descendants<Sheet>() ?? [])
        {
            if (sheet?.Id != null && workBookPart?.GetPartById(sheet.Id!) is WorksheetPart wsPart)
            {
                if (SheetName.IsNothing() || sheet.Name == SheetName)
                {
                    worksheet = wsPart.Worksheet;
                    break;
                }
            }
        }

        if (worksheet == null)
        {
            if (SheetName.IsNothing())
                throw new ArgumentException($"File {FilePath} does not contain any sheets");
            else
                throw new ArgumentException($"File {FilePath} does not contain sheet with name {SheetName}");
        }

        // Header
        var headerRow = worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex?.Value == HeaderRow);

        if (headerRow == null)
            throw new ArgumentException($"File {FilePath} does not contain header row {HeaderRow}");

        var headerIndex = 0;
        var column = HeaderStartColumn;
        mHeaders = new Dictionary<string, int>();
        var currentAddress = CellReference.GetColumnName(column) + HeaderRow.ToString();

        foreach (Cell c in headerRow.Elements<Cell>())
        {
            if (c.CellReference == currentAddress)
            {
                var header = GetCellValue(c)?.Replace(" ", "").ToUpper() ?? ("Col_" + column);
                mHeaders.Add(header, headerIndex++);
                currentAddress = CellReference.GetColumnName(++column) + HeaderRow.ToString();

                if (HeaderLength != null && mHeaders.Count >= HeaderLength)
                    break;
            }
        }

        // Data
        var currentRow = HeaderRow + 1;

        foreach (Row row in worksheet.Descendants<Row>())
        {
            if (row.RowIndex?.Value == currentRow)
            {
                var cells = new List<string>();

                for (uint col = HeaderStartColumn; col < HeaderStartColumn + mHeaders.Count; col++)
                {
                    var c = row.Elements<Cell>().Where(x => x.CellReference == CellReference.GetColumnName(col) + currentRow.ToString()).FirstOrDefault();
                    cells.Add(GetCellValue(c) ?? "");
                }

                // Break if empty row
                if (cells.Count == 0 || cells.All(x => x.IsNothing()))
                    break;

                yield return new ExcelDynamicRow(cells, mHeaders);
                currentRow++;
            }

            if (LastRow != null && currentRow > LastRow)
                break;
        }
    
        yield break;
    }

    private string? GetCellValue(Cell? cell)
    {
        if (cell == null)
            return null;

        if (cell.DataType == null)
            return cell.InnerText;

        var value = cell.InnerText;

        if (cell.DataType.Value == CellValues.SharedString) // For shared strings, look up the value in the shared strings table.
        {
            // Get worksheet from cell
            var parent = cell?.Parent;

            while (parent?.Parent != null && parent.Parent != parent && string.Compare(parent.LocalName, "worksheet", true) != 0)
            {
                parent = parent.Parent;
            }

            if (string.Compare(parent?.LocalName, "worksheet", true) != 0)
            {
                throw new Exception("Unable to find parent worksheet.");
            }

            if (parent is DocumentFormat.OpenXml.Spreadsheet.Worksheet ws &&
                ws?.WorksheetPart?.OpenXmlPackage is SpreadsheetDocument ssDoc)
            {
                var sstPart = ssDoc?.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                // lookup value in shared string table
                if (sstPart != null && sstPart.SharedStringTable != null)
                {
                    value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
        }
        else if (cell.DataType.Value == CellValues.Boolean)
        {
            value = value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                    ? "false"
                    : "true";
        }

        return value;
    }

    private void Init()
    {
        if (FilePath.IsNothing())
            throw new ArgumentException("FilePath must be set");

        if (HeaderRow < 1)
            throw new ArgumentException($"HeaderRow can't be less than 1 ({FilePath})");

        if (HeaderStartColumn < 1)
            throw new ArgumentException($"HeaderStartColumn can't be less than 1 ({FilePath})");

        if (HeaderLength != null && HeaderLength < 1)
            throw new ArgumentException($"HeaderLength can't be set to less than 1 ({FilePath})");
    }

    public void Dispose()
    {

    }
}
