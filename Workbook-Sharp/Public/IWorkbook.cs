namespace WorkbookSharp;

public interface IWorkbook
{
    IWorksheet AddWorksheet(string name = "");
    Task Save(string fileName);
    Task<byte[]> Save();
    IWorksheet this[string sheetName] { get; }
}