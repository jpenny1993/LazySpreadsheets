using LazySpreadsheets.Enums;

namespace LazySpreadsheets.Attributes;

/// <summary>
/// Defines how to format the property value in an Excel spreadsheet.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class CellFormatAttribute : Attribute
{
    public CellFormatAttribute(string cellFormatStr)
    {
        NumberFormatId = null;
        CellFormatStr = cellFormatStr;
    }

    public CellFormatAttribute(NumberFormats numberFormat)
    {
        NumberFormatId = (int)numberFormat;
        CellFormatStr = null;
    }

    public int? NumberFormatId { get; }

    public string? CellFormatStr { get; }
}
