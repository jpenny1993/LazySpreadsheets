using ClosedXML.Excel;
using LazySpreadsheets.Export;

namespace LazySpreadsheets.Extensions;

public static class EnumerableExtensions
{
    /// <summary>
    /// Converts the provided enumerable into an <see cref="IXLWorkbook"/>.
    /// </summary>
    public static IXLWorkbook ToWorkbook<TData>(this IEnumerable<TData> enumerable)
    {
        return new WorkbookBuilder().Sheet(enumerable).ToWorkbook();
    }
}