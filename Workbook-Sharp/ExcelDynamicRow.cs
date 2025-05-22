using System.Dynamic;

namespace WorkbookSharp;

internal class ExcelDynamicRow : DynamicObject
{
    private List<string> mCells;
    private Dictionary<string, int> mHeaders;

    internal ExcelDynamicRow(List<string> cells, Dictionary<string, int> headers)
    {
        mCells = cells;
        mHeaders = headers;
    }

    public override bool TryGetMember(GetMemberBinder binder, out object? result)
    {
        var name = binder.Name.ToUpper();

        if (mHeaders.ContainsKey(name))
        {
            result = mCells[mHeaders[name]];
        }
        else
        {
            result = null;
        }

        return true;
    }
}