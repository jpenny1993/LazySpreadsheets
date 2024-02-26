using ClosedXML.Excel;

namespace LazySpreadsheets.Extensions;

public static class XLCellExtensions
{
    public static CellReference CellReference(this IXLCell cell)
        => new (cell.Address.ColumnNumber, cell.Address.RowNumber);
}