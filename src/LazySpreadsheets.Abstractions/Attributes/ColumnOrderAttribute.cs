namespace LazySpreadsheets.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public sealed class ColumnOrderAttribute : Attribute
{
    public ColumnOrderAttribute(int order)
    {
        Order = order;
    }

    public int Order { get; }
}