namespace WorkbookSharp;

public class WorkbookFactory
{
    public static IWorkbook CreateWorkbook() => new Workbook();
    public static IExcelParser CreateExcelParser() => new ExcelParser();
}
