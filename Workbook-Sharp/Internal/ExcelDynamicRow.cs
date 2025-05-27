using System.Dynamic;

namespace WorkbookSharp;

internal class ExcelDynamicRow : DynamicObject
{
    private List<string> mCells;
    private Dictionary<string, int> mHeaders;

    internal ExcelDynamicRow(List<string> cells, Dictionary<string, int> headers)
    {
        if (cells.Count != headers.Count)
            throw new ArgumentException("Cells count must match headers count.");

        mCells = cells;
        mHeaders = headers;
    }

    public override bool TryGetMember(GetMemberBinder binder, out object? result)
    {
        var name = binder.Name.ToUpper();

        if (mHeaders.TryGetValue(name, out var index) && index < mCells.Count)
        {
            result = mCells[index] ?? "";     
        }
        else
        {
            result = null;
        }

        return true;
    }
}