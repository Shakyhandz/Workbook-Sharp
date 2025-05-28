using System.Reflection;
using TestConsoleApp;

// TODO:
// if worksheet font family and size set, use them for all cells and set row height of sheet 
// custom column widths
// custom row heights
// Extension methods: add amount (thousand separator no of decimals)
// Autofit formula columns (HARD to do)
// Autofit not working 100% yet (check test_to_Excel)
// Autofit not working for merged cells (check test_to_Excel)
// Hyperlinks

var test = new Test();

//test.GetEmbeddedResources();
//await test.TestTemplates();
//await test.TestToExcel($"test_to_Excel_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
//await test.TestToExcelMultipleSheets($"test_to_Excel_multi_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
await test.TestSheets($"test_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
//await test.TestReadExcelFile("test.xlsx");
//await test.TestParseExcelFile("test_parser.xlsx");


