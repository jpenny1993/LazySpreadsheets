using ClosedXML.Excel;
using LazySpreadsheets.Interfaces.Export;

namespace LazySpreadsheets.Export;

public sealed class WorkbookBuilder
{
    private readonly List<IWorksheetDefinition> _worksheetDefinitions = new ();

    private int GetNextSheetNumber() => _worksheetDefinitions.Count + 1;
    
    /// <summary>
    /// Creates a worksheet definition and defines all properties as columns.
    /// </summary>
    /// <param name="enumerable">A enumerable of <see cref="TData"/>.</param>
    /// <typeparam name="TData">The data record type.</typeparam>
    public WorkbookBuilder Sheet<TData>(IEnumerable<TData> enumerable)
    {
        var builder = new WorksheetBuilder<TData>(enumerable, GetNextSheetNumber());
        builder.DefineAllPropertiesAsColumns();
        _worksheetDefinitions.Add(builder);
        return this;
    }

    /// <summary>
    /// Creates a worksheet definition and defines all properties as columns.
    /// </summary>
    /// <param name="enumerable">A enumerable of <see cref="TData"/>.</param>
    /// <param name="sheetName">The name of the worksheet.</param>
    /// <typeparam name="TData">The data record type.</typeparam>
    public WorkbookBuilder Sheet<TData>(IEnumerable<TData> enumerable, string sheetName)
    {
        var builder = new WorksheetBuilder<TData>(enumerable, GetNextSheetNumber());
            builder.Name(sheetName)
                   .DefineAllPropertiesAsColumns();
        _worksheetDefinitions.Add(builder);
        return this;
    }
    
    /// <summary>
    /// Creates a worksheet definition.
    /// </summary>
    /// <param name="enumerable">A enumerable of <see cref="TData"/>.</param>
    /// <param name="worksheetBuilder">A predicate that defines the contents of the worksheet.</param>
    /// <typeparam name="TData">The data record type.</typeparam>
    public WorkbookBuilder Sheet<TData>(IEnumerable<TData> enumerable, Action<IWorksheetBuilder<TData>> worksheetBuilder)
    {
        var builder = new WorksheetBuilder<TData>(enumerable, GetNextSheetNumber());
        worksheetBuilder(builder);
        _worksheetDefinitions.Add(builder);
        return this;
    }

    /// <summary>
    /// Builds the worksheets and combines into a workbook.
    /// </summary>
    public IXLWorkbook ToWorkbook()
    {
        var workbook = new XLWorkbook();
        foreach (var builder in _worksheetDefinitions)
        {
            builder.AppendWorksheet(workbook);
        }

        return workbook;
    }
}