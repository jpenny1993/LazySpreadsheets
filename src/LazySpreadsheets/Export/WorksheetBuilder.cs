﻿using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;
using LazySpreadsheets.Attributes;
using LazySpreadsheets.Interfaces.Export;

namespace LazySpreadsheets.Export;

internal sealed class WorksheetBuilder<TData> : IWorksheetBuilder<TData>, IWorksheetDefinition
{
    private readonly IEnumerable<TData> _datasource;
    private readonly List<IColumnDefinition> _columnDefinitions;
    private string _worksheetName;
    private int _freezeColumns;
    private int _freezeRows;

    internal WorksheetBuilder(IEnumerable<TData> enumerable, int sheetNumber)
    {
        _worksheetName = $"Sheet {sheetNumber}";
        _datasource = enumerable;
        _columnDefinitions = new List<IColumnDefinition>();
    }

    private int GetNextColumnNumber() => _columnDefinitions.Count + 1;

    public void AppendWorksheet(IXLWorkbook workbook)
    {
        var worksheet = workbook.Worksheets.Add(_worksheetName);

        // Write column headers
        worksheet.Cell(1, 1).InsertData(new[] { _columnDefinitions.Select(d => d.ColumnHeader) });

        // Do model transformation using instructions from fluent api 
        var rowData = _datasource.Select(i => _columnDefinitions.Select(d => d.GetCellValue(i)));

        // Write row data, otherwise works fine with just the data source
        var dataRange = worksheet.Cell(2, 1).InsertData(rowData);

        if (dataRange.RowCount() == 0)
        {
            // Create empty table
            worksheet.Range(worksheet.FirstCell(), worksheet.Cell(2, _columnDefinitions.Count)).CreateTable();
        }
        else
        {
            // Customise how column data is displayed
            foreach (var column in _columnDefinitions)
            {
                column.ApplyStyles(dataRange.Column(column.ColumnNumber));
            }

            // Make data filterable in excel
            worksheet.Range(worksheet.FirstCell(), dataRange.LastCell()).CreateTable().SetAutoFilter();
        }

        // Resize columns for long text values
        worksheet.Columns().AdjustToContents();

        foreach (var column in _columnDefinitions.Where(c => c.ColumnWidth > 0))
        {
            worksheet.Column(column.ColumnNumber).Width = column.ColumnWidth;
        }

        // Add freeze pane
        if (_freezeRows > 0 || _freezeColumns > 0)
        {
            worksheet.SheetView.Freeze(_freezeRows, _freezeColumns);
        }
    }

    public IWorksheetBuilder<TData> Column<TProperty>(Expression<Func<TData, TProperty>> propertySelector)
    {
        var columnNumber = GetNextColumnNumber();
        var columnBuilder = new ColumnBuilder<TData, TProperty>(columnNumber, propertySelector);
        _columnDefinitions.Add(columnBuilder);
        return this;
    }

    public IWorksheetBuilder<TData> Column<TProperty>(Expression<Func<TData, TProperty>> propertySelector, Action<IColumnBuilder<TData, TProperty>> columnBuilder)
    {
        var columnNumber = GetNextColumnNumber();
        var builder = new ColumnBuilder<TData, TProperty>(columnNumber, propertySelector);
        columnBuilder(builder);
        _columnDefinitions.Add(builder);
        return this;
    }

    public IWorksheetBuilder<TData> Computed<TValue>(Action<IColumnBuilder<TData, TValue>> columnBuilder)
    {
        var columnNumber = GetNextColumnNumber();
        var builder = new ColumnBuilder<TData, TValue>(columnNumber);
        columnBuilder(builder);
        _columnDefinitions.Add(builder);
        return this;
    }

    public IWorksheetBuilder<TData> DefineAllPropertiesAsColumns()
    {
        var dataType = typeof(TData);
        var orderedProperties = dataType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty)
            .Where(propertyInfo => propertyInfo.GetCustomAttribute<IgnoreColumnAttribute>() == null)
            .Select((value, index) => (
                propertyInfo: value,
                index: value.GetCustomAttribute<ColumnOrderAttribute>()?.Order ?? index))
            .OrderBy(tuple => tuple.index)
            .Select(tuple => tuple.propertyInfo);

        foreach (var propertyInfo in orderedProperties)
        {
            // Skip ignored properties
            var ignoreAttr = propertyInfo.GetCustomAttribute<IgnoreColumnAttribute>();
            if (ignoreAttr != null) continue;
            
            var columnNumber = GetNextColumnNumber();
            var propertyType = propertyInfo.PropertyType;

            // Create property selector
            var parameterExpr = Expression.Parameter(dataType, dataType.Name);
            var propertyExpr = Expression.Property(parameterExpr, propertyInfo.Name);
            var propertyValueSelector = Expression.Lambda(typeof(Func<,>).MakeGenericType(dataType, propertyType), propertyExpr, parameterExpr);

            // Create generic parameters for ColumnBuilder<TData,TProperty>()
            var columnBuilderParams = new object[] { columnNumber, propertyValueSelector, propertyInfo };

            // Create instance of ColumnBuilder<TData,TProperty>()
            var columBuilderBaseType = typeof(ColumnBuilder<,>);
            var columnBuilderType = columBuilderBaseType.MakeGenericType(dataType, propertyType);
            var columnBuilder = (IColumnDefinition)Activator.CreateInstance(columnBuilderType, columnBuilderParams)!;

            _columnDefinitions.Add(columnBuilder);
        }

        return this;
    }

    public IWorksheetBuilder<TData> Name(string worksheetName)
    {
        var nameLength = Math.Min(worksheetName.Length, 31); // Max length is 31 chars in Excel
        _worksheetName = worksheetName.Substring(0, nameLength);
        return this;
    }

    public IWorksheetBuilder<TData> Freeze(int rows, int columns)
    {
        _freezeRows = rows;
        _freezeColumns = columns;
        return this;
    }
}