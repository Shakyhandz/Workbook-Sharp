using WorkbookSharp.Styles;

namespace WorkbookSharp.Cells;

public class CellRange : ICellRange
{
    private Worksheet _worksheet;
    private uint _startRow;
    private uint _startColumn;
    private uint _endRow;
    private uint _endColumn;

    internal CellRange(Worksheet worksheet)
    {
        _worksheet = worksheet;
        _startRow = 1;
        _startColumn = 1;
        _endRow = CellReference.MAX_ROW;
        _endColumn = CellReference.MAX_COLUMN;

        // No need to validate here, as the constructor sets allowed ranges
        //CellReference.ValidateCell(_startRow, _startColumn);
        //CellReference.ValidateCell(_endRow, _endColumn);
    }

    public object? Value
    {
        get => _worksheet.Actions.TryGetValue((_startRow, _startColumn, _startRow, _startColumn), out var action) ? action : null;
        set => _worksheet.SetValue((_startRow, _startColumn), value);        
    }

    public string? Formula
    {
        get => _worksheet.Actions.TryGetValue((_startRow, _startColumn, _startRow, _startColumn), out var action) && action is CellFormula f ? f.Formula : null;
        set => _worksheet.SetFormula((_startRow, _startColumn), value, false);
    }

    public bool Merge
    {
        get => _worksheet.Actions.TryGetValue((_startRow, _startColumn, _endRow, _endColumn), out var action) && action is CellMerge;
        set
        {
            if (_startColumn == _endColumn && _startRow == _endRow)
                throw new ArgumentException("For a merge, the cell range must be more than one cell");

            if (value)
            {
                // Set Style to left- and up-most column
                _worksheet.MergeCells((_startRow, _startColumn), (_endRow, _endColumn), _worksheet.Cells[_startRow, _startColumn].Style);
            }
            else
            {
                // Un-merge if set to false
                if (_worksheet.Actions.TryGetValue((_startRow, _startColumn, _endRow, _endColumn), out var action) && action is CellMerge mergeOp)
                    _worksheet.UnMergeCells(mergeOp);                    
            }
        }
    }

    public Style? Style
    {
        get => _worksheet.Actions.TryGetValue((_startRow, _startColumn, _startRow, _startColumn), out var action) && action.StyleIndex != null 
               ? _worksheet._workbook.styleManager.GetStyleFromIndex(action.StyleIndex.Value) 
               : null;
        set => _worksheet.SetStyle((_startRow, _startColumn), (_endRow, _endColumn), value);
    }

    public ICellRange this[string address]
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

    public ICellRange this[uint row, uint col]
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

    public ICellRange this[uint fromRow, uint fromCol, uint toRow, uint toCol]
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
