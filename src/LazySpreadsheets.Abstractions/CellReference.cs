namespace LazySpreadsheets;

/// <summary>
/// Represents a cell reference i.e. A1.
/// </summary>
public record CellReference
{
    private uint _column;
    private uint _row;
    
    /// <summary>
    /// The Column number i.e. 1, 2, 3, 4.
    /// </summary>
    public int ColumnNumber { get => Convert.ToInt32(_column); set => _column = Convert.ToUInt32(value); }

    /// <summary>
    /// The Column Letter i.e. A, B, C, D.
    /// </summary>
    public string ColumnLetter { get => ToColumnLetter(_column); set => _column = ToColumnNumber(value); }

    /// <summary>
    /// The row number i.e. 1, 2, 3, 4.
    /// </summary>
    public int RowNumber { get => Convert.ToInt32(_row); set => _row = Convert.ToUInt32(value); }

    public CellReference()
    {
        _column = 1u;
        _row = 1u;
    }

    public CellReference(uint column, uint row)
    {
        _column = column;
        _row = row;
    }

    public CellReference(int column, int row)
        : this(Convert.ToUInt32(column), Convert.ToUInt32(row))
    {
    }
 
    public CellReference(string column, uint row)
        :this(ToColumnNumber(column), row)
    {
    }

    public CellReference(string column, int rowNumber)
        : this(ToColumnNumber(column), Convert.ToUInt32(rowNumber))
    {
    }
   
    public CellReference(char column, uint row)
        :this(ToColumnNumber(column.ToString()), row)
    {
    }

    public CellReference(char column, int rowNumber)
        : this(column, Convert.ToUInt32(rowNumber))
    {
    }

    
    public CellReference(string cellRef)
    {
        var columnLetter = GetColumnLetter(cellRef);
        _column = ToColumnNumber(columnLetter);
        _row = GetRowNumber(cellRef);
    }

    public CellReference(CellReference cellRef)
    {
        _column = cellRef._column;
        _row = cellRef._row;
    }

    public static implicit operator CellReference(string text) => new (text);

    /// <summary>
    /// Returns a new Object with the same values.
    /// </summary>
    public CellReference Copy() => new (this);

    /// <summary>
    /// Returns a new object modified by the given values.
    /// </summary>
    public CellReference MutateBy(int columns, int rows)
    {
        var rowNumber = _row + rows;
        var nextRowNumber = rowNumber < 1 ? 1u : Convert.ToUInt32(rowNumber);

        var columnNumber = _column + columns;
        var nextColumnNumber = columnNumber < 1 ? 1u : Convert.ToUInt32(columnNumber);

        return new CellReference(nextColumnNumber, nextRowNumber);
    }

    /// <summary>
    /// Increments the column value by one.
    /// </summary>
    public void NextColumn() => _column += 1u;

    /// <summary>
    /// Increments the row value by one.
    /// </summary>
    public void NextRow() => _row += 1u;

    /// <summary>
    /// Decreases the column value by one.
    /// </summary>
    public void PreviousColumn()
    {
        if (_column > 1u) return;
        _column -= 1u;
    }

    /// <summary>
    /// Decreases the row value by one.
    /// </summary>
    public void PreviousRow()
    {
        if (_row == 1u) return;
        _row -= 1u;
    }

    /// <summary>
    /// Prints the cell reference as a string i.e. "A1".
    /// </summary>
    public override string ToString() => $"{ColumnLetter}{RowNumber}";

    /// <summary>
    /// Returns the row number from a cell reference string.
    /// i.e. "A1" would be "1"
    /// </summary>
    public static uint GetRowNumber(string cellReference)
    {
        var numberChars = cellReference.Where(char.IsNumber).ToArray();
        var rowNumber = new string(numberChars);
        return uint.TryParse(rowNumber, out var result) ? result : 0u;
    }

    /// <summary>
    /// Returns the column letter from a cell reference string.
    /// i.e. "A1" would be "A"
    /// </summary>
    public static string GetColumnLetter(string cellReference)
    {
        var columnChars = cellReference.Where(char.IsLetter).ToArray();
        return columnChars.Any() ? new string(columnChars) : "A";
    }

    /// <summary>
    /// Returns the column number from the given column letter.
    /// i.e. "A" would be "1"
    ///
    /// Do not use on a cell reference string.
    /// Should be chained from .GetColumnLetter()
    /// </summary>
    public static uint ToColumnNumber(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter))
            throw new ArgumentNullException(nameof(columnLetter));

        columnLetter = columnLetter.ToUpperInvariant();

        var sum = 0;

        for (var i = 0; i < columnLetter.Length; i++)
        {
            sum *= 26;
            sum += (columnLetter[i] - 'A' + 1);
        }

        return (uint)sum;
    }

    /// <summary>
    /// Returns the column letter from a column number.
    /// i.e. "1" would be "A"
    /// </summary>
    public static string ToColumnLetter(uint columnNumber)
    {
        var dividend = columnNumber;
        var columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }
}