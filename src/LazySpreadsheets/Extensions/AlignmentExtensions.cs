using ClosedXML.Excel;
using LazySpreadsheets.Enums;

namespace LazySpreadsheets.Extensions;

internal static class AlignmentExtensions
{
    internal static XLAlignmentHorizontalValues ToClosedXmlValue(this HorizontalAlignment alignment)
    {
        return alignment switch
        {
            HorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
            HorizontalAlignment.CenterContinuous => XLAlignmentHorizontalValues.CenterContinuous,
            HorizontalAlignment.Distributed => XLAlignmentHorizontalValues.Distributed,
            HorizontalAlignment.Fill => XLAlignmentHorizontalValues.Fill,
            HorizontalAlignment.General => XLAlignmentHorizontalValues.General,
            HorizontalAlignment.Justify => XLAlignmentHorizontalValues.Justify,
            HorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
            HorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null)
        };
    }

    internal static XLAlignmentVerticalValues ToClosedXmlValue(this VerticalAlignment alignment)
    {
        return alignment switch
        {
            VerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
            VerticalAlignment.Center => XLAlignmentVerticalValues.Center,
            VerticalAlignment.Distributed => XLAlignmentVerticalValues.Distributed,
            VerticalAlignment.Justify => XLAlignmentVerticalValues.Justify,
            VerticalAlignment.Top => XLAlignmentVerticalValues.Top,
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null)
        };
    }
}