using TestConsoleApp;

// TODO:
// if worksheet font family and size set, use them for all cells and set row height of sheet 
// font color
// gridlines
// custom column widths
// custom row heights
// formulas
// Extension methods: add amount (thousand separator no of decimals)
// Autofit formula columns (HARD to do)
// Autofit not working 100% yet (check test_to_Excel)
// Autofit not working for merged cells (check test_to_Excel)

var test = new Test();

await test.TestToExcel($"test_to_Excel_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
await test.TestToExcelMultipleSheets($"test_to_Excel_multi_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
//await test.TestSheets($"test_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
//await test.TestReadExcelFile("test.xlsx");
//await test.TestParseExcelFile("test_parser.xlsx");


