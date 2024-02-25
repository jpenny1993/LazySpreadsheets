namespace LazySpreadsheets;

/// <summary>
/// Represents a cell reference i.e. A1.
/// </summary>
public record CellReference
{
    /// <summary>
    /// The Column Letter i.e. A, B, C, D.
    /// </summary>
    public string ColumnLetter { get; set; }

    /// <summary>
    /// The row number i.e. 1, 2, 3, 4.
    /// </summary>
    public uint RowNumber { get; set; }

    public CellReference()
    {
        ColumnLetter = "A";
        RowNumber = 1;
    }

    public CellReference(string cellRef)
    {
        ColumnLetter = GetColumnLetter(cellRef);
        RowNumber = GetRowNumber(cellRef);
    }

    public CellReference(string columnLetter, uint rowNumber)
    {
        ColumnLetter = columnLetter;
        RowNumber = rowNumber;
    }

    public static implicit operator CellReference(string text) => new (text);

    public static implicit operator string(CellReference cellRef) => cellRef.ToString();

    /// <summary>
    /// Returns a new Object with the same values.
    /// </summary>
    public CellReference Copy() => new (ColumnLetter, RowNumber);

    /// <summary>
    /// Returns a new object modified by the given values.
    /// </summary>
    public CellReference MutateBy(int columns, int rows)
    {
        var uintValue = Convert.ToUInt32(rows);
        var nextRowNumber = RowNumber + uintValue;

        if (columns != 0)
        {
            var nextColumnNumber = ToColumnNumber(ColumnLetter) + columns;
            var nextColumnLetter = ToColumnLetter(nextColumnNumber);
            return new CellReference(nextColumnLetter, nextRowNumber);
        }

        return new CellReference(ColumnLetter, nextRowNumber);
    }

    /// <summary>
    /// Increments the column value by one.
    /// </summary>
    public void NextColumn()
    {
        var nextColumnNumber = ToColumnNumber(ColumnLetter) + 1;
        ColumnLetter = ToColumnLetter(nextColumnNumber);
    }

    /// <summary>
    /// Increments the row value by one.
    /// </summary>
    public void NextRow()
    {
        RowNumber++;
    }

    /// <summary>
    /// Decreases the column value by one.
    /// </summary>
    public void PreviousColumn(string columnLetter)
    {
        if (columnLetter == "A") return;

        var previousColumnNumber = ToColumnNumber(columnLetter) - 1; 
        ColumnLetter = ToColumnLetter(previousColumnNumber);
    }

    /// <summary>
    /// Decreases the row value by one.
    /// </summary>
    public void PreviousRow(string columnLetter)
    {
        if (RowNumber == 1) return;
        RowNumber--;
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
        return uint.TryParse(rowNumber, out var result) ? result : 0;
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
    public static int ToColumnNumber(string columnLetter)
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

        return sum;
    }

    /// <summary>
    /// Returns the column letter from a column number.
    /// i.e. "1" would be "A"
    /// </summary>
    public static string ToColumnLetter(int columnNumber)
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