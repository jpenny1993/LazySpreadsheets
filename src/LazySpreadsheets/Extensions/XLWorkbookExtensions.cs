using ClosedXML.Excel;

namespace LazySpreadsheets.Extensions;

public static class XLWorkbookExtensions
{
    /// <summary>
    /// Exports the Excel spreadsheet and converts to a byte array.
    /// </summary>
    public static byte[] ToBytes(this IXLWorkbook workbook, bool disposeWorkbook = false)
    {
        using var ms = workbook.ToMemoryStream(disposeWorkbook);
        return ms.ToArray();
    }

    /// <summary>
    /// Exports the Excel spreadsheet to the file system.
    /// </summary>
    public static void ToFile(this XLWorkbook workbook, string filename, bool disposeWorkbook = false)
    {
        workbook.SaveAs(filename);
        if (disposeWorkbook)
        {
            workbook.Dispose(); // Avoids having to create a using block on workbook
        }
    }

    /// <summary>
    /// Exports the Excel spreadsheet and converts to a <see cref="MemoryStream"/>.
    /// </summary>
    public static MemoryStream ToMemoryStream(this IXLWorkbook workbook, bool disposeWorkbook = false)
    {
        var memoryStream = new MemoryStream();
        workbook.SaveAs(memoryStream);
        if (disposeWorkbook)
        {
            workbook.Dispose(); // Avoids having to create a using block on workbook
        }

        memoryStream.Seek(0, SeekOrigin.Begin);
        return memoryStream;
    }
}