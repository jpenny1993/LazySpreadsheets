using ClosedXML.Excel;

namespace LazySpreadsheets.Interfaces.Export;

/// <summary>
/// Defines a worksheet in Excel workbook.
/// </summary>
internal interface IWorksheetDefinition
{
    /// <summary>
    /// Builds the worksheet by enumerating the data and applying the column definitions.
    /// </summary>
    void AppendWorksheet(IXLWorkbook workbook);
}