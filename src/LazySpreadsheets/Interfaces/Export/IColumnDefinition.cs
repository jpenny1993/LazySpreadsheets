using ClosedXML.Excel;

namespace LazySpreadsheets.Interfaces.Export;

/// <summary>
/// Defines a column in a worksheet.
/// This includes how to format and export cell values.
/// </summary>
internal interface IColumnDefinition
{
    int ColumnNumber { get; }

    string ColumnHeader { get; }

    int ColumnWidth { get; }

    void ApplyStyles(IXLRangeColumn column);

    object? GetCellValue(object? item);
}