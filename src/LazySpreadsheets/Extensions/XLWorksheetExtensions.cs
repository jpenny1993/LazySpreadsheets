using ClosedXML.Excel;

namespace LazySpreadsheets.Extensions;

public static class XLWorksheetExtensions
{
    public static IXLCell Cell(this IXLWorksheet worksheet, CellReference reference)
        => worksheet.Cell(reference.RowNumber, reference.ColumnNumber);

    public static IXLRange Range(this IXLWorksheet worksheet, CellReference first, CellReference last)
        => worksheet.Range( first.RowNumber, first.ColumnNumber, last.RowNumber, last.ColumnNumber);
}