using ClosedXML.Excel;
using LazySpreadsheets.Export;
using LazySpreadsheets.Demo.Factories;
using LazySpreadsheets.Demo.Models;
using LazySpreadsheets.Enums;

Console.WriteLine("Generating data...");

IEnumerable<Book> books = BookFactory.Book().Generate(count: 1000);

Console.WriteLine("Building spreadsheet...");

using var workbook = new WorkbookBuilder()
    .Sheet(books, "Books")
    .Sheet(books, sheet => sheet
        .Name("Book Report")
        .Column(book => book.Title)
        .Column(book => book.Published, col => col
            .Header("Publish Date")
            .Format(NumberFormats.DateTime)
        )
        .Column(book => book.Price, col => col
            .Subtotal()
            .ConditionalFormat(format => format
                .ColorScale()
                .LowestValue(XLColor.Red)
                .HighestValue(XLColor.Green))
        )
        .Computed<string?>(c => c
            .Header("Customer Rating")
            .DataValidation(data =>
            {
                data.List("\"1 star,2 stars,3 stars,4 stars,5 stars\"", inCellDropdown: true);
            })
        )
        .Freeze(rows: 1, columns: 0)
    )
    .ToWorkbook();

Console.WriteLine("Saving spreadsheet...");

workbook.SaveAs("test.xlsx");

Console.WriteLine("Spreadsheet saved!");