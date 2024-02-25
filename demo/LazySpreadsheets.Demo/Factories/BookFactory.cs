using Bogus;
using LazySpreadsheets.Demo.Models;

namespace LazySpreadsheets.Demo.Factories;

public static class BookFactory
{
    public static Faker<Book> Book() => new Faker<Book>()
        .RuleFor(b => b.Id, f => f.Random.Guid())
        .RuleFor(b => b.Isbn10, f => f.Random.ReplaceNumbers("##########"))
        .RuleFor(b => b.Isbn13, f => f.Random.ReplaceNumbers("###-##########"))
        .RuleFor(b => b.Title, f => f.Lorem.Sentence())
        .RuleFor(b => b.Author, f => f.Name.FullName())
        .RuleFor(b => b.Publisher, f => f.Company.CompanyName())
        .RuleFor(b => b.TotalPages, f => f.Random.Number(min: 100, max: 5000))
        .RuleFor(b => b.Price, f => f.Random.Decimal(min: 3, max: 30))
        .RuleFor(b => b.Published, f => f.Date.Past(yearsToGoBack: 30, DateTime.Today))
        .RuleFor(b => b.IsInStock, f => f.Random.Bool());
}