## Lazy Spreadsheets

Turns an enumerable into a data table export.

## Dependencies

- [ClosedXML](https://github.com/ClosedXML/ClosedXML)

## Usage examples

```csharp
using ClosedXML.Excel;
using LazySpreadsheets.Export;
using LazySpreadsheets.Extensions;
using LazySpreadsheets.Constants;
using LazySpreadsheets.Enums;

public void SaveToFile(IEnumerable<Book> books)
{
    return books.ToWorkbook().SaveAs("my-books.xlsx");
}

public MemoryStream SaveToStream(IEnumerable<Book> books)
{
    return new WorkbookBuilder()
        .Sheet(books)
        .ToWorkbook()
        .ToMemoryStream();
}

public byte[] SaveToBytes(IEnumerable<Book> books)
{
    using var workbook = new WorkbookBuilder()
        .Sheet(books, sheet => sheet
            .Name("Book Prices")
            .Column(book => book.Title)
            .Column(book => book.Published, col => col
                .Header("Publish Date")
                .Format(NumberFormats.DateTime)
            )
            .Column(book => book.Price, col => col
                .Format(CellFormats.AccountingGBP)
                .ConditionalFormat(format => format
                    .ColorScale()
                    .LowestValue(XLColor.Red)
                    .HighestValue(XLColor.Green))
            )
        )
        .ToWorkbook();

    return workbook.ToBytes();
}
```
