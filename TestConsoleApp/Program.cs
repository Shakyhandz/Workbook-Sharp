using WorkbookSharp;
using WorkbookSharp.Styles;


// TODO:
// font color
// column widths row heights
// gridlines
// auto fit columns



var workbook = new Workbook();

{
    var worksheet1 = workbook.AddWorksheet();
    worksheet1.FontFamily = XlFontFamily.Calibri;
    worksheet1.FontSize = 9;
    worksheet1.AutoFitColumns = true;

    worksheet1.SetValue("A1", "Test1");
    worksheet1.SetValue("A2", "Test2");
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

}
{
    var worksheet2 = workbook.AddWorksheet();
    worksheet2.FontFamily = XlFontFamily.Arial;
    worksheet2.FontSize = 14;
    worksheet2.AutoFitColumns = false;

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
File.WriteAllBytes($"test_{DateTime.Now:yyyyMMddHHmmss}.xlsx", bytes);

//new SpreadsheetReader().Read("test.xlsx");