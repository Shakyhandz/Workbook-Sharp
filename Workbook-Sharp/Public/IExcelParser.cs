namespace WorkbookSharp;

public interface IExcelParser
{
    /// <summary>
    /// If left empty the first sheet will be used.
    /// </summary>
    string SheetName { get; set; }
    /// <summary>
    /// Path and file name to an open xml Excel file
    /// </summary>
    string FilePath { get; set; }
    /// <summary>
    /// Set if header row isn't row 1
    /// </summary>
    uint HeaderRow { get; set; }
    /// <summary>
    /// Set if header doesn't start at column 1
    /// </summary>
    uint HeaderStartColumn { get; set; }
    /// <summary>
    /// Set if there are columns after the data being parsed
    /// </summary>
    int? HeaderLength { get; set; }
    /// <summary>
    /// Set if there are rows after the data being parsed
    /// </summary>
    uint? LastRow { get; set; }

    IEnumerable<dynamic> Execute();
}
