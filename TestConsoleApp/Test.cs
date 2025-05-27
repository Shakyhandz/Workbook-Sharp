using System.Reflection;
using WorkbookSharp;
using WorkbookSharp.Styles;
using static WorkbookSharp.WorkbookExtensions;

namespace TestConsoleApp;

internal class Test
{
    internal void GetEmbeddedResources()
    {
        var loader = WorkbookFactory.CreateTemplateLoader();
        foreach (var t in loader.GetEmbeddedExcelTemplates(Assembly.GetExecutingAssembly()))
            Console.WriteLine(t);
    }

    internal async Task TestTemplates()
    {
        var resource = "TestConsoleApp.Templates.test_template.xlsx";
        var loader = WorkbookFactory.CreateTemplateLoader();
        var workbook = loader.LoadWorkbookFromTemplate(Assembly.GetExecutingAssembly(), resource);

        var sheet1 = workbook["Sheet1"];        
        sheet1.SetValue("A1", "Updated!");
        await workbook.Save($"template_copy_{DateTime.Now.Ticks}.xlsx");
    }


    internal async Task TestToExcel(string fileName)
    {
        var testData = new[]
        {
            new { Id = 1, TypeId = new Guid("9B657229-DCA0-44DE-BC68-01B67D0C5030"), Name = "Doppel", IsActive = false, Date = new DateTime(2019, 08, 01) },
            new { Id = 2, TypeId = new Guid("2C44A758-3944-45E9-84CD-07DC3C99ED2E"), Name = "Marketing", IsActive = false, Date = new DateTime(2019, 09, 23) },
            new { Id = 3, TypeId = new Guid("B4171BB4-CEED-4DA4-8BDC-1010A39E33FA"), Name = "Native", IsActive = true, Date = new DateTime(2019, 09, 23) },
            new { Id = 4, TypeId = new Guid("9CA58FEE-4F2A-4401-80CD-25E6B7D09037"), Name = "Alpro", IsActive = true, Date = new DateTime(2019, 11, 05) },
            new { Id = 5, TypeId = new Guid("F21F1829-5144-460F-8F86-35F30D4B3D05"), Name = "Europe", IsActive = false, Date = new DateTime(2019, 11, 04) },
            new { Id = 6, TypeId = new Guid("9F09CAA5-9E55-40F2-9FFB-46C78BDCD01D"), Name = "Samsung", IsActive = false, Date = new DateTime(2019, 11, 11) },
            new { Id = 7, TypeId = new Guid("03474FDD-1436-480C-8ECB-52F185DAD5D8"), Name = "Signal", IsActive = false, Date = new DateTime(2019, 10, 03) },
            new { Id = 8, TypeId = new Guid("7DB1BBA4-5828-4D02-8F6B-54544ACB8408"), Name = "Social", IsActive = true, Date = new DateTime(2019, 11, 11) },
            new { Id = 9, TypeId = new Guid("B70637D2-E109-46FD-98BB-90F35D831561"), Name = "Hub", IsActive = false, Date = new DateTime(2019, 11, 11) },
            new { Id = 10, TypeId = new Guid("0D63613B-A496-45D8-BEDA-9F6882DBA900"), Name = "Iduna", IsActive = false, Date = new DateTime(2019, 10, 03) },
            new { Id = 11, TypeId = new Guid("38310602-253A-42AB-AB86-A302290C5245"), Name = "Clark", IsActive = false, Date = new DateTime(2019, 10, 04) },
            new { Id = 12, TypeId = new Guid("77D59C80-C276-4CB8-B5EC-AC84F72013AA"), Name = "Editorial", IsActive = false, Date = new DateTime(2020, 02, 07) },
            new { Id = 13, TypeId = new Guid("C8CAF791-30A0-45AB-9244-B3B48B4E3109"), Name = "Juice", IsActive = false, Date = new DateTime(2020, 01, 10) },
            new { Id = 14, TypeId = new Guid("34045D2C-7944-4995-973A-C1318361A3EA"), Name = "Plus", IsActive = false, Date = new DateTime(2020, 01, 10) },
            new { Id = 15, TypeId = new Guid("DE47E97C-14B4-4A05-8465-C97C44BB2A02"), Name = "Avira", IsActive = false, Date = new DateTime(2020, 02, 21) },
            new { Id = 16, TypeId = new Guid("9B1A31DE-8E9F-4E45-B129-CD3BC4FAC91D"), Name = "Truck", IsActive = false, Date = new DateTime(2020, 02, 27) },
            new { Id = 17, TypeId = new Guid("C2C87405-C049-4CB4-A5BD-DB6B46220495"), Name = "Forman", IsActive = true, Date = new DateTime(2020, 03, 16) },
            new { Id = 18, TypeId = new Guid("C5EB3C14-07FC-4E4B-8B61-EA7164B13DAF"), Name = "Penny", IsActive = false, Date = new DateTime(2020, 03, 30) },
            new { Id = 19, TypeId = new Guid("700DBB7F-AD75-4F5B-A782-F84C8EB59C9E"), Name = "Phone", IsActive = false, Date = new DateTime(2020, 02, 18) },
            new { Id = 20, TypeId = new Guid("30AB4A9B-F21B-4B07-B380-FA426745C81C"), Name = "Water", IsActive = false, Date = new DateTime(2020, 02, 18) },
        };

        var bytes = await testData.ToExcel(new SpreadsheetExportOptions
        {
            SheetName = "Report",
            FontFamily = XlFontFamily.Calibri,
            FontSize = 9,
            ColumnStyles = new()
            {
                { 5, new Style { Border = XlBorder.Around, DateFormat = XlDateFormat.DateHoursMinutes } }
            }
        });

        File.WriteAllBytes(fileName, bytes);
    }

    internal async Task TestToExcelMultipleSheets(string fileName)
    {
        var testData = new[]
        {
            new { Id = 1, TypeId = new Guid("9B657229-DCA0-44DE-BC68-01B67D0C5030"), Name = "Doppel", IsActive = false, Date = new DateTime(2019, 08, 01) },
            new { Id = 2, TypeId = new Guid("2C44A758-3944-45E9-84CD-07DC3C99ED2E"), Name = "Marketing", IsActive = false, Date = new DateTime(2019, 09, 23) },
            new { Id = 3, TypeId = new Guid("B4171BB4-CEED-4DA4-8BDC-1010A39E33FA"), Name = "Native", IsActive = true, Date = new DateTime(2019, 09, 23) },
            new { Id = 4, TypeId = new Guid("9CA58FEE-4F2A-4401-80CD-25E6B7D09037"), Name = "Alpro", IsActive = true, Date = new DateTime(2019, 11, 05) },
            new { Id = 5, TypeId = new Guid("F21F1829-5144-460F-8F86-35F30D4B3D05"), Name = "Europe", IsActive = false, Date = new DateTime(2019, 11, 04) },
            new { Id = 6, TypeId = new Guid("9F09CAA5-9E55-40F2-9FFB-46C78BDCD01D"), Name = "Samsung", IsActive = false, Date = new DateTime(2019, 11, 11) },
            new { Id = 7, TypeId = new Guid("03474FDD-1436-480C-8ECB-52F185DAD5D8"), Name = "Signal", IsActive = false, Date = new DateTime(2019, 10, 03) },
            new { Id = 8, TypeId = new Guid("7DB1BBA4-5828-4D02-8F6B-54544ACB8408"), Name = "Social", IsActive = true, Date = new DateTime(2019, 11, 11) },
            new { Id = 9, TypeId = new Guid("B70637D2-E109-46FD-98BB-90F35D831561"), Name = "Hub", IsActive = false, Date = new DateTime(2019, 11, 11) },
            new { Id = 10, TypeId = new Guid("0D63613B-A496-45D8-BEDA-9F6882DBA900"), Name = "Iduna", IsActive = false, Date = new DateTime(2019, 10, 03) },
            new { Id = 11, TypeId = new Guid("38310602-253A-42AB-AB86-A302290C5245"), Name = "Clark", IsActive = false, Date = new DateTime(2019, 10, 04) },
            new { Id = 12, TypeId = new Guid("77D59C80-C276-4CB8-B5EC-AC84F72013AA"), Name = "Editorial", IsActive = false, Date = new DateTime(2020, 02, 07) },
            new { Id = 13, TypeId = new Guid("C8CAF791-30A0-45AB-9244-B3B48B4E3109"), Name = "Juice", IsActive = false, Date = new DateTime(2020, 01, 10) },
            new { Id = 14, TypeId = new Guid("34045D2C-7944-4995-973A-C1318361A3EA"), Name = "Plus", IsActive = false, Date = new DateTime(2020, 01, 10) },
            new { Id = 15, TypeId = new Guid("DE47E97C-14B4-4A05-8465-C97C44BB2A02"), Name = "Avira", IsActive = false, Date = new DateTime(2020, 02, 21) },
            new { Id = 16, TypeId = new Guid("9B1A31DE-8E9F-4E45-B129-CD3BC4FAC91D"), Name = "Truck", IsActive = false, Date = new DateTime(2020, 02, 27) },
            new { Id = 17, TypeId = new Guid("C2C87405-C049-4CB4-A5BD-DB6B46220495"), Name = "Forman", IsActive = true, Date = new DateTime(2020, 03, 16) },
            new { Id = 18, TypeId = new Guid("C5EB3C14-07FC-4E4B-8B61-EA7164B13DAF"), Name = "Penny", IsActive = false, Date = new DateTime(2020, 03, 30) },
            new { Id = 19, TypeId = new Guid("700DBB7F-AD75-4F5B-A782-F84C8EB59C9E"), Name = "Phone", IsActive = false, Date = new DateTime(2020, 02, 18) },
            new { Id = 20, TypeId = new Guid("30AB4A9B-F21B-4B07-B380-FA426745C81C"), Name = "Water", IsActive = false, Date = new DateTime(2020, 02, 18) },
        };

        //var sheets = new[] 
        //{
        //    new SpreadsheetExportSet(testData.Take(7), (SpreadsheetExportOptions?)null),
        //    new SpreadsheetExportSet(testData.Skip(7), null),
        //};
        var sheets = new[] 
        {
            (testData.Take(7), (SpreadsheetExportOptions?)null),
            (testData.Skip(7), null),
        };
        await sheets.ToExcelMulti(fileName);
    }

    internal async Task TestSheets(string fileName)
    {
        var workbook = WorkbookFactory.CreateWorkbook();

        {
            var worksheet1 = workbook.AddWorksheet();
            worksheet1.FontFamily = XlFontFamily.Calibri;
            worksheet1.FontSize = 9;
            worksheet1.AutoFitColumns = true;
            
            worksheet1.SetValue("A1", "Test1");
            worksheet1.SetValue("A2", "Test2", new Style { FontColor = System.Drawing.Color.Red });
            worksheet1.SetValue("A3", 15000, new Style { UseThousandSeparator = true });
            worksheet1.SetValue("A4", true, new Style { FillColor = System.Drawing.Color.FromArgb(0xE7, 0xE6, 0xE6) });
            worksheet1.SetValue("A5", new DateTime(2025, 2, 2), new Style { FillColor = System.Drawing.Color.FromArgb(221, 235, 247) });
            worksheet1.SetValue("A6", 0.2, new Style { FillColor = System.Drawing.Color.DarkBlue });
            worksheet1.Cells[1, 1].Value = "overwrite";
            worksheet1.SetValue("A7", 0.23M);
            worksheet1.SetValue("A8", long.MaxValue);
            worksheet1.SetValue("J5", "Thunder", new Style { Border = XlBorder.Top | XlBorder.Bottom });
            worksheet1.MergeCells("A8", "C9", new Style { Border = XlBorder.Top | XlBorder.Bottom, HorizontalAlignment = XlHorizontalAlignment.Center });
            worksheet1.SetValue("A11", new DateTime(2001, 2, 3, 4, 5, 6), new Style { DateFormat = XlDateFormat.DateHoursMinutesSeconds });
            worksheet1.Cells[10, 5].Value = "Thunder";

            worksheet1.SetValue((6, 12), "what the h is long.MaxValue so long for");
            worksheet1.Cells["A6"].Style = new Style { FillColor = System.Drawing.Color.LimeGreen };

            worksheet1.SetValue("C1", 15000, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetValue("C2", 230000, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetValue("C3", 7000.45, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetFormula("C4", "=SUM(C1:C3)", false, new Style { UseThousandSeparator = true, decimalPlaces = 2 });

            worksheet1.SetValue("D1", 67.5, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetValue("D2", 30, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetValue("D3", 22.42, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetFormula("D4", "=SUM(R[-3]C:R[-1]C)", true, new Style { UseThousandSeparator = true, decimalPlaces = 2 });

            worksheet1.SetValue("D6", 67.5, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetValue("D7", 30, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetValue("D8", 22.42, new Style { UseThousandSeparator = true, decimalPlaces = 2 });
            worksheet1.SetFormula("D5", "=SUM(R1C:R3C)", true, new Style { UseThousandSeparator = true, decimalPlaces = 2 });

        }
        {
            var worksheet2 = workbook.AddWorksheet();
            worksheet2.FontFamily = XlFontFamily.Arial;
            worksheet2.FontSize = 14;
            worksheet2.AutoFitColumns = true;
            worksheet2.ShowGridlines = false;

            worksheet2.Cells[15, 15].Value = 6321456;
            worksheet2.SetValue("J5", "Thunder", new Style { FontSize = 12, });
            worksheet2.SetValue("B2", "Test decoration", new Style { FontDecoration = XlFontDecoration.Bold | XlFontDecoration.Strikeout });
            worksheet2.SetValue("B3", 2500000);
            worksheet2.SetValue("B4", true);
            worksheet2.SetValue("B5", new DateTime(2025, 2, 2));
            worksheet2.SetValue("B6", 0.2, new Style { IsPercentage = true });
            worksheet2.SetValue("B7", 0.23M, new Style { decimalPlaces = 4 });
            worksheet2.SetValue("B8", long.MaxValue);
            worksheet2.MergeCells("B8", "C9");
            worksheet2.SetValue((6, 12), "what the h is long.MaxValue so long for");
        }

        var bytes = await workbook.Save();
        File.WriteAllBytes(fileName, bytes);
    }

    internal async Task TestParseExcelFile(string fileName)
    {
        await Task.Yield();

        var parser = WorkbookFactory.CreateExcelParser();

        parser.FilePath = fileName;
        //parser.HeaderLength = 4;
        parser.LastRow = 11;        

        var col = parser.Execute()
                        .Select(x => new
                        {

                            Id = int.Parse((string)x.Id),
                            TypeId = Guid.Parse((string)x.TypeId),
                            Name = (string)x.Name,
                            IsActive = bool.Parse((string)x.IsActive),
                            Date = x.Date != null ? (string)x.Date : "",
                        })
                        .ToList();

    }
}
