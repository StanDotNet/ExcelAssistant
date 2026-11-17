using static ExcelAssistant.Tests.TestSupport;

namespace ExcelAssistant.Tests;

public class RoundTripTests
{
    private static readonly List<Vehicle> Vehicles =
    [
        new("Sierra 1500", 5, 4, 2024, 12_345_678_901L, 2450.5f, 0.1 + 0.2, 0.015m,
            Guid.Parse("5b8e1c1e-7f6a-4c43-9d55-0c2a3e1f9a10"), TimeSpan.FromDays(1095), new TimeOnly(8, 30),
            2, 1200.50m, "trade-in"),
        new("Outback", 7, 5, 2019, 0, 1700f, 4.75, 19_999.99m,
            Guid.Empty, TimeSpan.FromHours(36), new TimeOnly(17, 0, 15),
            null, null, null),
    ];

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Typed_values_survive_a_round_trip(ExcelType type)
    {
        using var stream = Write(type, Vehicles);
        using var reader = Reader(stream, type);

        var read = reader.Read<Vehicle>("Sheet");

        Assert.Equal(
        [
            Vehicles[0] with { Rating = type == ExcelType.xls ? 0.3 : 0.1 + 0.2 },
            Vehicles[1] with { OpensAt = new TimeOnly(17, 0) },
        ], read);
    }

    [Theory]
    [InlineData(ExcelType.xlsx, "0.30000000000000004")]
    [InlineData(ExcelType.xls, "0.3")]
    public void Raw_cell_text_is_what_typed_reads_parse(ExcelType type, string rating)
    {
        using var stream = Write(type, Vehicles);
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        var rows = reader.Read().ToList();
        var first = rows[0];

        Assert.Equal("Sierra 1500", first["Model"]);
        Assert.Equal("12345678901", first["Mileage"]);
        Assert.Equal("2450.5", first["Weight"]);
        Assert.Equal(rating, first["Rating"]);
        Assert.Equal("0.015", first["Price"]);
        Assert.Equal("1095.00:00:00", first["Warranty"]);
        Assert.Equal("08:30", first["OpensAt"]);
        Assert.Equal("1200.5", first["Discount"]);
        Assert.Equal("trade-in", first["Notes"]);
        Assert.Equal("17:00", rows[1]["OpensAt"]);
        Assert.Equal("", rows[1]["Owners"]);
        Assert.Equal("", rows[1]["Notes"]);
    }

    [Theory]
    [InlineData(ExcelType.xlsx, "0.015", "0.015")]
    [InlineData(ExcelType.xlsx, "0.1", "0.1")]
    [InlineData(ExcelType.xlsx, "1.005", "1.005")]
    [InlineData(ExcelType.xlsx, "-42.42", "-42.42")]
    [InlineData(ExcelType.xlsx, "123456789.123456789", "123456789.12345679")]
#if NET11_0_OR_GREATER
    [InlineData(ExcelType.xlsx, "12345678901234.567", "12345678901234.566")]
#else
    [InlineData(ExcelType.xlsx, "12345678901234.567", "12345678901234.568")]
#endif
    [InlineData(ExcelType.xls, "0.015", "0.015")]
    [InlineData(ExcelType.xls, "0.1", "0.1")]
    [InlineData(ExcelType.xls, "1.005", "1.005")]
    [InlineData(ExcelType.xls, "-42.42", "-42.42")]
    [InlineData(ExcelType.xls, "123456789.123456789", "123456789.1")]
    public void Decimals_go_through_double_and_back_through_text(ExcelType type, string written, string read)
    {
        using var stream = Write(type, [new Amount(decimal.Parse(written))]);
        using var reader = Reader(stream, type);

        var value = Assert.Single(reader.Read<Amount>("Sheet")).Value;

        Assert.Equal(decimal.Parse(read), value);
    }

    [Fact]
    public void Xls_shows_large_numbers_in_scientific_notation_which_decimals_cannot_parse()
    {
        using var stream = Write(ExcelType.xls, [new Amount(12345678901234.567m)]);
        using var reader = Reader(stream, ExcelType.xls);
        reader.OpenSheet("Sheet");

        Assert.Equal("1.23457E+13", reader.Read().Single()["Value"]);
        Assert.Throws<FormatException>(() => reader.Read<Amount>("Sheet"));
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Decimals_round_trip_under_a_comma_culture(ExcelType type)
    {
        using var culture = Culture("pl-PL");
        using var stream = Write(type, [new Amount(1200.55m)]);
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        Assert.Equal("1200,55", reader.Read().Single()["Value"]);
        Assert.Equal(1200.55m, reader.Read<Amount>("Sheet").Single().Value);
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Settable_properties_are_filled_and_get_only_ones_skipped(ExcelType type)
    {
        using var stream = Write(type, [new Settable { Title = "a", Count = 3, Total = 9.99m }]);
        using var reader = Reader(stream, type);

        var read = Assert.Single(reader.Read<Settable>("Sheet"));

        Assert.Equal("a", read.Title);
        Assert.Equal(3, read.Count);
        Assert.Equal(9.99m, read.Total);
        Assert.Equal("a:3", read.Computed);
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Bools_are_written_as_excel_booleans_but_cannot_be_read_into_a_model(ExcelType type)
    {
        using var stream = Write(type, [new Flag("a", true), new Flag("b", false)]);
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        Assert.Equal(["TRUE", "FALSE"], reader.Read().Select(r => r["Enabled"]));
        Assert.Throws<NotSupportedException>(() => reader.Read<Flag>("Sheet"));
    }

    [Theory]
    [InlineData(ExcelType.xlsx, "46343.57326388889")]
    [InlineData(ExcelType.xls, "46343.57326")]
    public void Written_dates_have_no_date_format_and_read_back_as_serial_numbers(ExcelType type, string serial)
    {
        var at = new DateTime(2026, 11, 17, 13, 45, 30);
        using var stream = Write(type, [new Stamp(at)]);
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        Assert.Equal(serial, reader.Read().Single()["At"]);
        Assert.Throws<FormatException>(() => reader.Read<Stamp>("Sheet"));
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Written_date_only_values_read_back_as_serial_numbers(ExcelType type)
    {
        using var stream = Write(type, [new Day(new DateOnly(2026, 1, 2))]);
        using var reader = Reader(stream, type);
        reader.OpenSheet("Sheet");

        Assert.Equal("46024", reader.Read().Single()["On"]);
        Assert.Throws<FormatException>(() => reader.Read<Day>("Sheet"));
    }
}
