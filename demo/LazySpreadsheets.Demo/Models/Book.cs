using System.ComponentModel;
using LazySpreadsheets.Attributes;
using LazySpreadsheets.Constants;
using LazySpreadsheets.Enums;

namespace LazySpreadsheets.Demo.Models;

public record Book
{
    [IgnoreColumn]
    public Guid Id { get; set; }

    [ColumnOrder(9)]
    [DisplayName("ISBN-10")]
    public string Isbn10 { get; set; } = default!;
    
    [ColumnOrder(10)]
    [DisplayName("ISBN-13")]
    public string Isbn13 { get; set; } = default!;
    
    [ColumnWidth(40)]
    public string Title { get; set; } = default!;

    public string Author { get; set; } = default!;

    [CellFormat(NumberFormats.IntegerWithComma)]
    public int TotalPages { get; set; }
    
    public string Publisher { get; set; } = default!;

    [CellFormat(NumberFormats.Date)]
    public DateTime Published { get; set; }

    [Subtotal]
    [CellFormat(CellFormats.AccountingGBP)]
    public decimal Price { get; set; }

    [ColumnOrder(1)]
    [DisplayName("In Stock")]
    [CellAlignment(HorizontalAlignment.Center)]
    [CellFormat(CellFormats.BooleanTickOnly)]
    public bool IsInStock { get; set; }
}