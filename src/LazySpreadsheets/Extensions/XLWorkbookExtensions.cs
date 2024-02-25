using ClosedXML.Excel;

namespace LazySpreadsheets.Extensions;

public static class XLWorkbookExtensions
{
    /// <summary>
    /// Exports the Excel spreadsheet and converts to a byte array.
    /// </summary>
    public static byte[] ToBytes(this IXLWorkbook workbook)
    {
        using var ms = workbook.ToMemoryStream();
        return ms.ToArray();
    }

    /// <summary>
    /// Exports the Excel spreadsheet and converts to a <see cref="MemoryStream"/>.
    /// </summary>
    public static MemoryStream ToMemoryStream(this IXLWorkbook workbook)
    {
        var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return memoryStream;
    }
}