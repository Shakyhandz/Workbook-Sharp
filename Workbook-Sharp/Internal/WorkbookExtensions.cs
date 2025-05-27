using WorkbookSharp.Styles;

namespace WorkbookSharp;

public static class WorkbookExtensions
{
    //public record SpreadsheetExportSet<T>(IEnumerable<T> Collection, SpreadsheetExportOptions? Options = null);

    //public static async Task ToExcel<T>(this IEnumerable<SpreadsheetExportSet<T>> collections, string fileName, CancellationToken cancellationToken = default)
    public static async Task ToExcelMulti<T>(this IEnumerable<(IEnumerable<T> collection, SpreadsheetExportOptions? options)> collections, string fileName, CancellationToken cancellationToken = default)
    {
        if (fileName.IsNothing())
            throw new ArgumentNullException(nameof(fileName), "File name must be set");

        Workbook? workbook = null;

        foreach (var collection in collections)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (collection.collection == null)
                throw new ArgumentException("Collection can't be null", nameof(collections));

            workbook = await collection.collection.ToWorkbook(workbook, collection.options, cancellationToken);
        }

        if (workbook == null)
            throw new InvalidOperationException("No collections provided to export");

        await workbook.Save(fileName);
    }

    //public static async Task<byte[]> ToExcel<T>(this IEnumerable<SpreadsheetExportSet<T>> collections, CancellationToken cancellationToken = default)
    public static async Task<byte[]> ToExcelMulti<T>(this IEnumerable<(IEnumerable<T> collection, SpreadsheetExportOptions? options)> collections, CancellationToken cancellationToken = default)
    {
        Workbook? workbook = null;

        foreach (var collection in collections)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (collection.collection == null)
                throw new ArgumentException("Collection can't be null", nameof(collections));

            workbook = await collection.collection.ToWorkbook(workbook, collection.options, cancellationToken);
        }

        if (workbook == null)
            throw new InvalidOperationException("No collections provided to export");

        return await workbook.Save();
    }

    public static async Task ToExcel<T>(this IEnumerable<T> collection, string fileName, SpreadsheetExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (fileName.IsNothing())
            throw new ArgumentNullException(nameof(fileName), "File name must be set");

        var workbook = await collection.ToWorkbook(options: options, cancellationToken: cancellationToken);
        await workbook.Save(fileName);
    }

    public static async Task<byte[]> ToExcel<T>(this IEnumerable<T> collection, SpreadsheetExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        var workbook = await collection.ToWorkbook(options: options, cancellationToken: cancellationToken);
        return await workbook.Save();
    }

    // TODO: make public?
    private static async Task<Workbook> ToWorkbook<T>(this IEnumerable<T> collection, Workbook? workbook = null, SpreadsheetExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        await Task.Yield();

        if (typeof(T).IsTuple())
            throw new InvalidOperationException("Tuples are not supported in Excel export. Use a named class or anonymous object instead.");

        options ??= new SpreadsheetExportOptions();

        if (options.HeaderRow < 1)
            throw new ArgumentException($"Header row ({options.HeaderRow}) can't be less than 1");

        if (options.StartColumn < 1)
            throw new ArgumentException($"Start column ({options.StartColumn}) can't be less than 1");

        workbook = workbook ??= new Workbook();
        var worksheet = workbook.AddWorksheet(options.SheetName);

        if (options.FontFamily != null)
            worksheet.FontFamily = options.FontFamily.Value;

        worksheet.FontSize = options.FontSize;

        // Add header row
        var props = typeof(T).GetProperties();

        var column = options.StartColumn;
        var row = options.HeaderRow;

        options.HeaderStyle ??= new Style
        {
            FontDecoration = XlFontDecoration.Bold,
            FontSize = (options.FontSize ?? 11) + 1,
            FillColor = System.Drawing.Color.FromArgb(0xE6, 0xE6, 0xE6),
            Border = XlBorder.Around,
            HorizontalAlignment = XlHorizontalAlignment.Center
        };

        options.ItemsStyle ??= new Style
        {
            Border = XlBorder.Around,
        };

        foreach (var prop in props)
        {
            cancellationToken.ThrowIfCancellationRequested();
            worksheet.SetValue((row, column++), prop.Name, options.HeaderStyle.Clone());
        }

        row++;

        // Add items
        foreach (var item in collection)
        {
            cancellationToken.ThrowIfCancellationRequested();

            uint colIndex = options.StartColumn;

            foreach (var prop in props)
            {
                var value = prop.GetValue(item);

                var style = options.ColumnStyles != null && options.ColumnStyles.TryGetValue(colIndex, out var s)
                            ? s.Clone()
                            : options.ItemsStyle.Clone();

                worksheet.SetValue((row, colIndex++), value, style);
            }

            row++;
        }

        return workbook;
    }

}
