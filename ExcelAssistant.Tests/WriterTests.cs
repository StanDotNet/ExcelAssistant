using static ExcelAssistant.Tests.TestSupport;

namespace ExcelAssistant.Tests;

public class WriterTests
{
    private static readonly List<Report> Reports =
    [
        new("Alice", "alice@mail.com", 1200.50m),
        new("Bob", "bob@mail.com", 850m),
    ];

    private static readonly List<Customer> Customers =
    [
        new(Guid.Parse("0f8fad5b-d9cb-469f-a165-70867728950e"), "Alice", "alice@mail.com"),
        new(Guid.Parse("7c9e6679-7425-40de-944b-e07fc1f90ae7"), "Bob", "bob@mail.com"),
    ];

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Custom_mappings_rename_headers_and_map_back_on_read(ExcelType type)
    {
        using var stream = Write(type, Reports, w =>
        {
            w.AddMapping<Report>(r => r.CustomerName, "Client");
            w.AddMapping<Report>(r => r.Balance, "Account Balance");
        });
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        Assert.Equal(["Client", "Email", "Account Balance"], reader.Read().First().Keys);

        reader.AddMapping<Report>(r => r.CustomerName, "Client");
        reader.AddMapping<Report>(r => r.Balance, "Account Balance");
        Assert.Equal(Reports, reader.Read<Report>("Sheet"));
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Excluded_properties_are_not_written_and_read_as_default(ExcelType type)
    {
        using var stream = Write(type, Customers, w =>
        {
            w.Exclude<Customer>(c => c.Id);
            w.Exclude("Email");
        });
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        Assert.Equal(["Name"], reader.Read().First().Keys);
        Assert.Equal(
            Customers.Select(c => c with { Id = Guid.Empty, Email = null! }),
            reader.Read<Customer>("Sheet"));
    }

    [Fact]
    public void Exclusions_in_the_configuration_are_ignored_by_the_writer()
    {
        using var writer = new ExcelWriter(new ExcelConfiguration { Exclusions = ["Id"] });
        writer.BuildSheet(Customers, "Sheet");
        using var stream = new MemoryStream();
        writer.CopyTo(stream);
        using var reader = Reader(stream, ExcelType.xlsx);
        reader.OpenSheet("Sheet");

        Assert.Equal(["Id", "Name", "Email"], reader.Read().First().Keys);
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Multiple_sheets_are_written_and_read_by_name(ExcelType type)
    {
        using var writer = new ExcelWriter(new ExcelConfiguration { ExcelType = type });
        writer.AddMapping<Report>(r => r.CustomerName, "Customer Name");
        writer.AddMapping<Customer>(c => c.Email, "Email Address");
        writer.BuildSheet(Reports, "Reports");
        writer.BuildSheet(Customers, "Customers");
        using var stream = new MemoryStream();
        writer.CopyTo(stream);

        using var reader = Reader(stream, type);
        reader.AddMapping<Report>(r => r.CustomerName, "Customer Name");
        reader.AddMapping<Customer>(c => c.Email, "Email Address");

        Assert.Equal(Reports, reader.Read<Report>());
        Assert.Equal(Customers, reader.Read<Customer>("Customers"));
        Assert.Equal(Reports, reader.Read<Report>("Reports"));
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void WriteRecords_writes_a_single_sheet(ExcelType type)
    {
        using var writer = new ExcelWriter(new ExcelConfiguration { ExcelType = type });
        using var stream = writer.WriteRecords(new MemoryStream(), Reports);

        Assert.Equal(0, stream.Position);
        Assert.True(stream.Length > 0);

        using var reader = Reader(stream, type);
        Assert.Equal(Reports, reader.Read<Report>());
    }

    [Fact]
    public void WriteRecords_leaves_the_stream_empty_when_there_are_no_records()
    {
        using var writer = new ExcelWriter();
        using var stream = writer.WriteRecords(new MemoryStream(), new List<Report>());

        Assert.Equal(0, stream.Length);
    }

    [Theory]
    [InlineData(ExcelType.xlsx, ExcelConstants.XlsxContentType)]
    [InlineData(ExcelType.xls, ExcelConstants.XlsContentType)]
    public void Content_type_follows_the_excel_type(ExcelType type, string contentType)
    {
        using var writer = new ExcelWriter(new ExcelConfiguration { ExcelType = type });

        Assert.Equal(contentType, writer.ContentType);
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Column_widths_are_clamped_between_min_and_max(ExcelType type)
    {
        var records = new List<Report> { new(new string('x', 100), "a@b.c", 1m) };
        using var writer = new ExcelWriter(new ExcelConfiguration { ExcelType = type });
        writer.BuildSheet(records, "Sheet");
        using var stream = new MemoryStream();
        writer.CopyTo(stream);

        var sheet = NPOI.SS.UserModel.WorkbookFactory.Create(stream).GetSheet("Sheet");

        Assert.Equal(60 * 300, sheet.GetColumnWidth(0));
        Assert.Equal(10 * 300, sheet.GetColumnWidth(1));
        Assert.Equal(10 * 300, sheet.GetColumnWidth(2));
    }
}
