using WorkbookSharp.Styles;

namespace WorkbookSharp.Cells;

public class CellRange(Worksheet _worksheet, uint _startRow, uint _startColumn, uint _endRow, uint _endColumn)
{
    public object Value
    {
        // TODO: Getter
        set => _worksheet.SetValue((_startRow, _startColumn), value);        
    }

    public bool Merge
    {
        set
        {
            if (_startColumn == _endColumn && _startRow == _endRow)
                throw new ArgumentException("For a merge, the cell range must be more than one cell");

            // TODO: Set Style to left- and up-most column
            //var leftTopCell = _worksheet.Actions.GetValueOrDefault((_startRow, _startColumn, _startRow, _startColumn));
            //_worksheet.MergeCells((_startRow, _startColumn), (_endRow, _endColumn), leftTopCell?.Style);

            _worksheet.MergeCells((_startRow, _startColumn), (_endRow, _endColumn), null);

            // TODO: Un-merge
        }
    }

    public Style? Style
    {
        // TODO: Getter
        //get
        //{
        //    if (_worksheet.Actions.TryGetValue((_startRow, _startColumn, _endRow, _endColumn), out var action))
        //        return action.Style;

        //    return null;
        //}
        set => _worksheet.SetStyle((_startRow, _startColumn), value);        
    }

    public CellRange this[string address]
    {
        get
        {
            if (address.Contains(":"))
            {
                var start = address.Split(':')[0];
                var (_, startRow, startCol) = CellReference.ParseAddress(start);
                
                var end = address.Split(':')[1];
                var (_, endRow, endCol) = CellReference.ParseAddress(end);

                return this[startRow, startCol, endRow, endCol];
            }

            var (_, row, col) = CellReference.ParseAddress(address);
            
            return this[row, col];
        }
    }

    public CellRange this[uint row, uint col]
    {
        get
        {
            CellReference.ValidateCell(row, col);

            _startColumn = col;
            _startRow = row;
            _endColumn = col;
            _endRow = row;

            return this;
        }
    }

    public CellRange this[uint fromRow, uint fromCol, uint toRow, uint toCol]
    {
        get
        {
            CellReference.ValidateCell(fromRow, fromCol);
            CellReference.ValidateCell(toRow, toCol);

            _startColumn = Math.Min(fromCol, toCol);
            _startRow = Math.Min(fromRow, toRow);
            _endColumn = Math.Max(fromCol, toCol);
            _endRow = Math.Max(fromRow, toRow);

            return this;
        }
    }
}
