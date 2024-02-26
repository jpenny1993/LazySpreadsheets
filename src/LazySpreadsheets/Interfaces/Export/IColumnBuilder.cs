using ClosedXML.Excel;
using LazySpreadsheets.Enums;

namespace LazySpreadsheets.Interfaces.Export;

/// <summary>
/// Builds a column for a property defined on an item. 
/// </summary>
/// <typeparam name="TData">A list item model.</typeparam>
/// <typeparam name="TProperty">A property member type.</typeparam>
public interface IColumnBuilder<out TData, in TProperty>
{
    /// <summary>
    /// Sets the horizontal alignment on the cell content.
    /// </summary>
    IColumnBuilder<TData, TProperty> Align(HorizontalAlignment horizontalAlignment);

    /// <summary>
    /// Sets the vertical alignment on the cell content.
    /// </summary>
    IColumnBuilder<TData, TProperty> Align(VerticalAlignment verticalAlignment);

    /// <summary>
    /// Sets the alignment on the cell content.
    /// </summary>
    IColumnBuilder<TData, TProperty> Align(
        HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment);

    /// <summary>
    /// Configures conditional formatting for the column cells.
    /// </summary>
    IColumnBuilder<TData, TProperty> ConditionalFormat(Action<IXLConditionalFormat> action);

    /// <summary>
    /// Configures data validation for the column cells.
    /// </summary>
    IColumnBuilder<TData, TProperty> DataValidation(Action<IXLDataValidation> action);
    
    /// <summary>
    /// Sets the column header/name to the provided value.
    /// </summary>
    IColumnBuilder<TData, TProperty> Header(string columnHeader);

    /// <summary>
    /// Sets the cell format to the provided excel compatible format.
    /// </summary>
    IColumnBuilder<TData, TProperty> Format(NumberFormats numberFormats);

    /// <summary>
    /// Sets the cell format to the provided excel compatible format.
    /// </summary>
    IColumnBuilder<TData, TProperty> Format(string cellFormatStr);

    /// <summary>
    /// Sets the cell value as the provided formula.
    /// </summary>
    IColumnBuilder<TData, TProperty> Formula(string formulaA1);

    /// <summary>
    /// Configures a subtotal cell to display above the column.
    /// </summary>
    IColumnBuilder<TData, TProperty> Subtotal();
    
    /// <summary>
    /// Sets the cell value to a static value.
    /// </summary>
    IColumnBuilder<TData, TProperty> Value(TProperty staticValue);

    /// <summary>
    /// Sets the cell value to a calculated value.
    /// </summary>
    IColumnBuilder<TData, TProperty> Value(Func<TData, TProperty> valueSelector);
    
    /// <summary>
    /// Sets the column width to a fixed value.
    /// </summary>
    IColumnBuilder<TData, TProperty> Width(int width);
}