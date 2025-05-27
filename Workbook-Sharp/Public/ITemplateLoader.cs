using System.Reflection;

namespace WorkbookSharp;

public interface ITemplateLoader
{
    IEnumerable<string> GetEmbeddedExcelTemplates(Assembly assembly);
    IWorkbook LoadWorkbookFromTemplate(Assembly assembly, string resourcePath);
    //IWorkbook LoadWorkbookFromTemplate(Stream stream);
    //IWorkbook LoadWorkbookFromFile(string filePath);
}
