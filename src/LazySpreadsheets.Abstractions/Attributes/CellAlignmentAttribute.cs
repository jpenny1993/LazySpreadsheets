using LazySpreadsheets.Enums;

namespace LazySpreadsheets.Attributes;

/// <summary>
/// Defines how to align cell content in an Excel spreadsheet.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class CellAlignmentAttribute : Attribute
{
    public CellAlignmentAttribute(HorizontalAlignment horizontal)
    {
        Horizontal = horizontal;
    }

    public CellAlignmentAttribute(VerticalAlignment vertical)
    {
        Vertical = vertical;
    }

    public CellAlignmentAttribute(HorizontalAlignment horizontal, VerticalAlignment vertical)
    {
        Horizontal = horizontal;
        Vertical = vertical;
    }

    public HorizontalAlignment? Horizontal { get; }

    public VerticalAlignment? Vertical { get; }
}