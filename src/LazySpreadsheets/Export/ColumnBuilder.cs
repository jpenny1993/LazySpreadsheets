using System.ComponentModel;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;
using LazySpreadsheets.Attributes;
using LazySpreadsheets.Enums;
using LazySpreadsheets.Extensions;
using LazySpreadsheets.Interfaces.Export;

namespace LazySpreadsheets.Export;

/// <summary>
/// Defines how to populate a column in the workbook. 
/// </summary>
/// <typeparam name="TData">List object type.</typeparam>
/// <typeparam name="TProperty">Displayed property type.</typeparam>
internal sealed class ColumnBuilder<TData, TProperty> : IColumnBuilder<TData, TProperty>, IColumnDefinition
{
    private Func<TData, TProperty>? _valueSelector;
    private int? _numberFormatId;
    private string? _cellFormatStr;
    private string? _formulaA1;
    private HorizontalAlignment? _horizontalAlignment;
    private VerticalAlignment? _verticalAlignment;
    private Action<IXLConditionalFormat>? _conditionalFormatAction;
    private Action<IXLDataValidation>? _dataValidationAction;

    public int ColumnNumber { get; }
    public string ColumnHeader { get; private set; }
    public int ColumnWidth { get; private set; }
    public bool HasSubtotal { get; private set; }

    /// <summary>
    /// Create a ColumnBuilder for a computed or static column.
    /// </summary>
    /// <param name="columnNumber">Column number.</param>
    public ColumnBuilder(int columnNumber)
    {
        ColumnNumber = columnNumber;
        ColumnHeader = $"Column {columnNumber}";
    }

    /// <summary>
    /// Create a ColumnBuilder for the provided property.
    /// </summary>
    /// <param name="columnNumber">Column number.</param>
    /// <param name="propertySelector">Member property selector.</param>
    public ColumnBuilder(int columnNumber, Expression<Func<TData, TProperty>> propertySelector)
        : this(columnNumber, propertySelector, ExpressionExtensions.GetPropertyInfo(propertySelector))
    {
    }

    /// <summary>
    /// Create a ColumnBuilder for the provided generic property.
    /// </summary>
    /// <param name="columnNumber">Column number.</param>
    /// <param name="valueSelector">Member property selector.</param>
    /// <param name="propertyInfo">Member property info.</param>
    public ColumnBuilder(int columnNumber, Expression<Func<TData, TProperty>> valueSelector, MemberInfo propertyInfo)
    {
        _valueSelector = valueSelector.Compile();
        var displayNameAttr = propertyInfo.GetCustomAttribute<DisplayNameAttribute>();
        var cellFormatAttr = propertyInfo.GetCustomAttribute<CellFormatAttribute>();
        var cellAlignmentAttr = propertyInfo.GetCustomAttribute<CellAlignmentAttribute>();
        var columnWidthAttr = propertyInfo.GetCustomAttribute<ColumnWidthAttribute>();
        var subtotalAttr = propertyInfo.GetCustomAttribute<SubtotalAttribute>();
        ColumnNumber = columnNumber;
        ColumnHeader = displayNameAttr?.DisplayName ?? propertyInfo.Name;
        _numberFormatId = cellFormatAttr?.NumberFormatId;
        _cellFormatStr = cellFormatAttr?.CellFormatStr;
        _horizontalAlignment = cellAlignmentAttr?.Horizontal;
        _verticalAlignment = cellAlignmentAttr?.Vertical;
        ColumnWidth = columnWidthAttr?.Width ?? 0;
        HasSubtotal = subtotalAttr != null;
    }

    public IColumnBuilder<TData, TProperty> Align(HorizontalAlignment horizontalAlignment)
    {
        _horizontalAlignment = horizontalAlignment;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Align(VerticalAlignment verticalAlignment)
    {
        _verticalAlignment = verticalAlignment;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Align(HorizontalAlignment horizontalAlignment,
        VerticalAlignment verticalAlignment)
    {
        _horizontalAlignment = horizontalAlignment;
        _verticalAlignment = verticalAlignment;
        return this;
    }

    public IColumnBuilder<TData, TProperty> ConditionalFormat(Action<IXLConditionalFormat> action)
    {
        _conditionalFormatAction = action;
        return this;
    }

    public IColumnBuilder<TData, TProperty> DataValidation(Action<IXLDataValidation> action)
    {
        _dataValidationAction = action;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Format(NumberFormats numberFormats)
    {
        _numberFormatId = (int)numberFormats;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Format(string cellFormatStr)
    {
        _cellFormatStr = cellFormatStr;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Formula(string formulaA1)
    {
        _formulaA1 = formulaA1;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Header(string columnHeader)
    {
        ColumnHeader = columnHeader;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Subtotal()
    {
        HasSubtotal = true;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Value(TProperty staticValue)
    {
        _valueSelector = x => staticValue;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Value(Func<TData, TProperty> valueSelector)
    {
        _valueSelector = valueSelector;
        return this;
    }

    public IColumnBuilder<TData, TProperty> Width(int width)
    {
        ColumnWidth = width;
        return this;
    }

    public void ApplyStyles(IXLRange column)
    {
        if (_numberFormatId != null)
        {
            column.Style.NumberFormat.SetNumberFormatId(_numberFormatId.Value);
        }
        else if (!string.IsNullOrEmpty(_cellFormatStr))
        {
            column.Style.NumberFormat.SetFormat(_cellFormatStr);
        }

        if (_horizontalAlignment != null)
        {
            column.Style.Alignment.Horizontal = _horizontalAlignment.Value.ToClosedXmlValue();
        }

        if (_verticalAlignment != null)
        {
            column.Style.Alignment.Vertical = _verticalAlignment.Value.ToClosedXmlValue();
        }

        if (!string.IsNullOrEmpty(_formulaA1) && 
            column.FirstCell().CellReference().RowNumber > 1) // is not subtotal
        {
            column.FormulaA1 = _formulaA1;
        }

        _conditionalFormatAction?.Invoke(column.AddConditionalFormat());

        _dataValidationAction?.Invoke(column.CreateDataValidation());
    }

    public object? GetCellValue(object? item)
    {
        if (item is null || _valueSelector is null)
            return null;

        var cellValue = _valueSelector.Invoke((TData)item);

        // Convert boolean to number if a custom number format is applied
        if (cellValue is bool b &&
            !string.IsNullOrEmpty(_cellFormatStr))
        {
            return b ? 1 : 0;
        }

        return cellValue;
    }
}