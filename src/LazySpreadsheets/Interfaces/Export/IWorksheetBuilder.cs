using System.ComponentModel;
using System.Linq.Expressions;
using LazySpreadsheets.Attributes;

namespace LazySpreadsheets.Interfaces.Export;

/// <summary>
/// Builds an excel worksheet from an <see cref="IEnumerable{TData}"/>.
/// </summary>
/// <typeparam name="TData">A list item model.</typeparam>
public interface IWorksheetBuilder<TData>
{
    /// <summary>
    /// Adds a column definition to the workbook builder.
    /// </summary>
    /// <param name="propertySelector">The data to display.</param>
    /// <typeparam name="TProperty">The data type.</typeparam>
    IWorksheetBuilder<TData> Column<TProperty>(Expression<Func<TData, TProperty>> propertySelector);

    /// <summary>
    /// Adds a column definition to the workbook builder.
    /// </summary>
    /// <param name="propertySelector">The data to display.</param>
    /// <param name="columnBuilder">A predicate that defines the contents of the column.</param>
    /// <typeparam name="TProperty">The data type.</typeparam>
    IWorksheetBuilder<TData> Column<TProperty>(Expression<Func<TData, TProperty>> propertySelector, Action<IColumnBuilder<TData, TProperty>> columnBuilder);

    /// <summary>
    /// Adds a computed column definition to the workbook builder.
    /// </summary>
    /// <param name="columnBuilder">A predicate that defines the contents of the column.</param>
    /// <typeparam name="TValue">The data type.</typeparam>
    IWorksheetBuilder<TData> Computed<TValue>(Action<IColumnBuilder<TData, TValue>> columnBuilder);

    /// <summary>
    /// Defines all columns using reflection, columns are ordered by position in the class.
    /// Column headers can be customised using <see cref="DisplayNameAttribute"/>.
    /// Number Formatting can be customised using <see cref="CellFormatAttribute"/>. 
    /// </summary>
    IWorksheetBuilder<TData> DefineAllPropertiesAsColumns();

    /// <summary>
    /// Removes a defined column from the worksheet builder by property selector.
    /// </summary>
    /// <param name="propertySelector">The data to remove.</param>
    /// <typeparam name="TProperty">The data type.</typeparam>
    IWorksheetBuilder<TData> Ignore<TProperty>(Expression<Func<TData, TProperty>> propertySelector);

    /// <summary>
    /// Removes a defined column from the worksheet builder by column header.
    /// </summary>
    /// <param name="columnHeader">Display name of the column.</param>
    IWorksheetBuilder<TData> Ignore(string columnHeader);

    /// <summary>
    /// Sets the worksheet name.
    /// </summary>
    /// <param name="worksheetName">Worksheet name. (default: "Sheet 1")</param>
    IWorksheetBuilder<TData> Name(string worksheetName);
    
    /// <summary>
    /// Sets a freeze pane on the worksheet.
    /// </summary>
    /// <param name="rows">Total rows to freeze. (default: 0)</param>
    /// <param name="columns">Total columns to freeze. (default: 0)</param>
    IWorksheetBuilder<TData> Freeze(int rows, int columns);
    
    /// <summary>
    /// Sets a freeze pane on the worksheet.
    /// </summary>
    IWorksheetBuilder<TData> Freeze(CellReference cellReference);
}