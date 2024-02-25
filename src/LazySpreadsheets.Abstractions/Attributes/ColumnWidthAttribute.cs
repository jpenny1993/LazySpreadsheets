namespace LazySpreadsheets.Attributes;

/// <summary>
/// Defines how wide a column will be when the spreadsheet is exported.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ColumnWidthAttribute : Attribute
{
    public ColumnWidthAttribute(int width)
    {
        Width = width;
    }

    public int Width { get; }
}