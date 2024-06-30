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
    using var workbook = return books.ToWorkbook();
    workbook.ToFile("my-books.xlsx");
}

public MemoryStream SaveToStream(IEnumerable<Book> books)
{
    using var workbook = new WorkbookBuilder()
        .Sheet(books)
        .ToWorkbook();

    return workbook.ToMemoryStream();
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
---

## Building a Spreadsheet

### Attributes

#### DisplayNameAttribute

Use the DisplayName attribute to customise the column name for your data.

```csharp
using System.ComponentModel;

public class MyReportItem 
{
    [DisplayName("My Display Name")]
    public string MyProperty { get; set; }   
}
```

#### CellAlignmentAttribute

Use the CellAlignment attribute to configure how your data will be aligned in Excel.

```csharp
using LazySpreadsheets.Attributes;
using LazySpreadsheets.Constants;

public class MyReportItem 
{
    [CellAlignment(HorizontalAlignment.Center, VerticalAlignment.Middle)]
    public string MyProperty { get; set; }   
}
```

#### CellFormatAttribute

Use the CellFormat attribute to configure how your data will display in Excel.

This attribute will accept either a  [NumberFormat as described by ClosedXml](https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table) or a [custom format as a string](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-3.0.1).

```csharp
using LazySpreadsheets.Attributes;
using LazySpreadsheets.Constants;

public class MyReportItem 
{
    [CellFormat(CellFormats.Date)]
    public DateTime? MyProperty { get; set; }   
}
```

### WorkbookBuilder

The following APIs are functions on a WorkbookBuilder.

#### Sheet

Adds an instance of WorksheetBuilder for the provided enumerable.

`.Sheet(IEnumerable<TData> data)`

`.Sheet(IEnumerable<TData> data, string name)`

`.Sheet(IEnumerable<TData> data, Action<IWorksheetBuilder> predicate)`

#### ToWorkbook

Converts the current WorkbookBuilder into an XLWorkbook.

`.ToWorkbook()` returns `XLWorkbook`

### WorksheetBuilder

The following APIs are functions on a WorksheetBuilder, an instance of this can only be created WorkbookBuilder.

#### SheetName

Sets the worksheet name, in Excel the max length of a worksheet is 31 characters.

`.Name(string name)`

#### Column

Adds an instance of ColumnBuilder for the provided class property.

`.Column(Func<TData, TProperty>> propertySelector)`

.`Column(Func<TData, TProperty> propertySelector, Action<IColumnBuilder> predicate)`

#### Computed

Adds an instance of ColumnBuilder where you can define a custom value or value function.

`.Computed<TValue>(Action<IColumnBuilder> predicate)`

#### DefineAllPropertiesAsColumns

Creates a column definition for all properties in the class.

This is the underlying function called when you call `.Sheet(IEnumerable<Data> data)`.

### ColumnBuilder

The following APIs are functions on a ColumnBuilder, an instance of this can only be created WorksheetBuilder.

#### Alignment

Set the horizontal and/or vertical alignment on the cell content.

`.Align(HorizontalAlignment horizontalAlignment)`

`.Align(VerticalAlignment verticalAlignment)`

`.Align(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)`

#### Column Header

Sets the column header to the provided value.

`.Header(string name)`

#### Conditional Format

Add conditional formatting to your column as provided by [Closed XML](https://docs.closedxml.io/en/latest/api/index.html#class-ClosedXML.Excel.XLConditionalFormats).

`.ConditionalFormat(options => ...)`

#### Data Validation

Add data validation to your column as provided by [Closed XML](https://docs.closedxml.io/en/latest/api/index.html#interface-ClosedXML.Excel.IXLDataValidation).

`.DataValidation(options => ...)`

#### Format

Sets the cell format to the provided [excel compatible format](https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table).

Most of the formats you will want from the link above, plus a few extra custom formats are available in [CellFormats](#cell-formats) 

`.Format(string cellFormat)`

`.Format(NumberFormats numberFormat)`

#### Formula

Allows you to set the cell value to a [custom formula](https://docs.closedxml.io/en/latest/features/formulas.html#feature-overview).

`.Formula(string formulaA1)`

#### Subtotal Headers

Adds an extra row above the headers on the workbook. For each colum where subtotal was specified a formula will be added to calculate the subtotal for that column. 

`.Subtotal()`

#### Value

Sets the cell value to a static or calculated value of the column type.

`.Value<TValue>(TValue staticValue)`

`.Value<TValue>(Func<TData, TValue> valueSelector)`

#### Column Width

Sets the column width to a fixed value in NOC format.
Finding your desired width will likely be trial and error.

`.Width(int width)`

> NoC are a non-linear units displayed as a column width in Excel, next to pixels.
> NoC combined with default font of the workbook can express width of the column in pixels and other units.

---

## Constant Values

### Cell Formats

The following values are provided by `LazySpreadsheets.Constants.CellFormats`

If you can't find what you want then try using [NumberFormatId](https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table).

Accounting and Currency are different, don't ask me why Excel has 2 options for this.

| Member Name      | Value                                                                         | Example          |
|------------------|-------------------------------------------------------------------------------|------------------|
| AccountingGBP    | _-\"£\"* #,##0.00_-;\\-\"£\"* #,##0.00_-;_-\"£\"* \"-\"??_-;_-@_-             | £ 30.00          |
| AccountingUSD    | _-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* \"-\"??_ ;_-@_" | $ 30.00          |
| AccountingEUR    | _-[$?-2]\\ * #,##0.00_-;\\-[$?-2]\\ * #,##0.00_-;_-[$?-2]\\ * \"-\"??_-;_-@_- | € 30.00          |
| BooleanYN        | \"Y\";;\"N\";                                                                 | Y or N           |
| BooleanNOnly     | [=0]\"N\";                                                                    | N or blank       |
| BooleanYOnly     | [=1]\"Y\";                                                                    | Y or blank       |
| BooleanYesNo     | \"YES\";;\"NO\";                                                              | YES or NO        |
| BooleanNoOnly    | [=0]\"NO\";                                                                   | NO or blank      |
| BooleanYesOnly   | [=1]\"YES\";                                                                  | YES or blank     |
| BooleanTickCross | \"\u2714\ufe0f\";;\"\u2716\ufe0f\";                                           | ✔ or ❌          |
| BooleanTickOnly  | [=1]\"\u2714\ufe0f\";                                                         | ✔ or blank       |
| BooleanCrossOnly | [=0]\"\u2716\ufe0f\";                                                         | ❌ or blank      |
| CurrencyGBP      | \"£\"#,##0.00                                                                 | £30.00           |
| CurrencyUSD      | [$$-409]#,##0.00                                                              | $30.00           |
| CurrencyEUR      | [$?-2]\\ #,##0.00                                                             | €30.00           |
| Date             | dd/mm/yyyy                                                                    | 30/06/2024       |
| DateTime         | dd/mm/yyyy HH:mm                                                              | 30/06/2024 11:31 |
| LongDate         | [$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy                                         | 30 June 2024     |
| Time             | [$-F400]h:mm:ss\\ AM/PM                                                       | 11:31 AM         |


### Content Types

The following values are provided by `LazySpreadsheets.Constants.ContentTypes`

| Member Name          | Value                                                                |
|----------------------|----------------------------------------------------------------------|
| MacroEnabledWorkbook | application/vnd.ms-excel.sheet.macroenabled.12                       |
| MacroEnabledTemplate | application/vnd.ms-excel.template.macroenabled.12                    |
| Workbook             | application/vnd.openxmlformats-officedocument.spreadsheetml.sheet    |
| Template             | application/vnd.openxmlformats-officedocument.spreadsheetml.template |

### File Extensions

The following values are provided by `LazySpreadsheets.Constants.FileExtensions`

| Member Name          | Value |
|----------------------|-------|
| MacroEnabledWorkbook | .xlsm |
| MacroEnabledTemplate | .xltm |
| Workbook             | .xlsx |
| Template             | .xltx |

---

## Exporting an XLWorkbook

An `XLWorkbook` implements `IDisposable` it is recommended to dispose of it when you're finished.

### ToBytes

Use the `.ToBytes()` extension to convert your spreadsheet into a byte array.

Set the `disposeWorkbook` parameter to true to dispose the workbook after the memory stream has been created, otherwise you will need to dispose the XLWorkbook yourself.

```csharp
using LazySpreadsheets.Constants;
using LazySpreadsheets.Export;
using LazySpreadsheets.Extensions;

byte[] bytes = enumerable
    .ToWorkbook()
    .ToBytes(disposeWorkbook: true);
```

### ToMemoryStream

Use the `.ToMemoryStream()` extension to convert your spreadsheet into a memory stream.

Set the `disposeWorkbook` parameter to true to dispose the workbook after the memory stream has been created, otherwise you will need to dispose the XLWorkbook yourself.

When returning as part of a web request you may also need the content type `ContentTypes.Workbook`.

```csharp
using LazySpreadsheets.Constants;
using LazySpreadsheets.Export;
using LazySpreadsheets.Extensions;

MemoryStream ms = enumerable
    .ToWorkbook()
    .ToMemoryStream(disposeWorkbook: true);

return File(ms, ContentTypes.Workbook);
```

### ToFile

This calls the `.SaveAs(filename)` function provided by ClosedXml to write your spreadsheet as a file in the filesystem. 

Set the `disposeWorkbook` parameter to true to dispose the workbook after the memory stream has been created, otherwise you will need to dispose the XLWorkbook yourself.
 
When setting the filename you will need to include the file extension `FileExtensions.Workbook`.

```csharp
using LazySpreadsheets.Constants;
using LazySpreadsheets.Export;
using LazySpreadsheets.Extensions;

enumerable
    .ToWorkbook()
    .ToFile(filename: "test" + FileExtensions.Workbook, disposeWorkbook: true);
```