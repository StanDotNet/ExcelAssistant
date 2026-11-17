using NPOI.SS.UserModel;
using static ExcelAssistant.Tests.TestSupport;

namespace ExcelAssistant.Tests;

public class ReaderTests
{
    [Theory]
    [InlineData(ExcelType.xlsx, "17-Nov-2026", "02-Jan-2026", "2026-11-17T00:00:00")]
    [InlineData(ExcelType.xls, "2026-11-17 13:45:30", "1/2/26", "2026-11-17T13:45:30")]
    public void Date_formatted_cells_are_read_through_their_display_text(
        ExcelType type, string first, string second, string firstParsed)
    {
        using var stream = Sheet(type, ["At"],
            row => DateCell(row, new DateTime(2026, 11, 17, 13, 45, 30), "yyyy-mm-dd hh:mm:ss"),
            row => DateCell(row, new DateTime(2026, 1, 2), "m/d/yy"));
        using var reader = Reader(stream, type);
        reader.OpenSheet();

        Assert.Equal([first, second], reader.Read().Select(r => r["At"]));
        Assert.Equal(
            [DateTime.Parse(firstParsed), new DateTime(2026, 1, 2)],
            reader.Read<Stamp>().Select(s => s.At));
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Text_dates_are_parsed_with_their_time(ExcelType type)
    {
        using var stream = Sheet(type, ["At", "On"],
            row =>
            {
                row.CreateCell(0).SetCellValue("2026-11-17 08:05:09");
                row.CreateCell(1).SetCellValue("2026-11-17");
            });
        using var reader = Reader(stream, type);

        Assert.Equal(new DateTime(2026, 11, 17, 8, 5, 9), reader.Read<Stamp>().Single().At);
        Assert.Equal(new DateOnly(2026, 11, 17), reader.Read<Day>().Single().On);
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Fuzzy_headers_are_matched_to_properties(ExcelType type)
    {
        using var stream = Sheet(type, ["customer_name", "Email Address", " BALANCE "],
            row =>
            {
                row.CreateCell(0).SetCellValue("Alice");
                row.CreateCell(1).SetCellValue("alice@mail.com");
                row.CreateCell(2).SetCellValue(10.5);
            });
        using var reader = Reader(stream, type);

        Assert.Equal(new Report("Alice", "alice@mail.com", 10.5m), reader.Read<Report>().Single());
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Headers_below_the_matching_percentage_are_left_unmapped(ExcelType type)
    {
        MemoryStream Sheet() => TestSupport.Sheet(type, ["CustomerName", "Emial", "Balance"],
            row =>
            {
                row.CreateCell(0).SetCellValue("Alice");
                row.CreateCell(1).SetCellValue("alice@mail.com");
                row.CreateCell(2).SetCellValue(1);
            });

        using var strict = Reader(Sheet(), type);
        using var loose = Reader(Sheet(), type, matchingPercentage: 75);

        Assert.Null(strict.Read<Report>().Single().Email);
        Assert.Equal("alice@mail.com", loose.Read<Report>().Single().Email);
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Headers_containing_another_property_name_collide(ExcelType type)
    {
        using var stream = Sheet(type, ["Email", "Email Backup"],
            row => row.CreateCell(0).SetCellValue("a@mail.com"));
        using var reader = Reader(stream, type);

        Assert.Throws<ArgumentException>(() => reader.Read<Report>());
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Blank_rows_are_skipped_and_missing_cells_default(ExcelType type)
    {
        using var stream = Sheet(type, ["CustomerName", "Email", "Balance"],
            row => row.CreateCell(0).SetCellValue("Alice"),
            row => row.CreateCell(0, CellType.Blank),
            row => row.CreateCell(2).SetCellValue(3));
        using var reader = Reader(stream, type);

        Assert.Equal(
            [new Report("Alice", null!, 0m), new Report(null!, null!, 3m)],
            reader.Read<Report>());
    }

    [Theory]
    [InlineData(ExcelType.xlsx)]
    [InlineData(ExcelType.xls)]
    public void Opening_a_reader_closes_the_source_stream(ExcelType type)
    {
        var stream = Sheet(type, ["Email"], row => row.CreateCell(0).SetCellValue("a@mail.com"));
        using var reader = Reader(stream, type);

        Assert.False(stream.CanRead);
        Assert.Equal("a@mail.com", reader.Read<Report>().Single().Email);
    }

    private static void DateCell(IRow row, DateTime value, string format)
    {
        var cell = row.CreateCell(0);
        cell.SetCellValue(value);
        var workbook = row.Sheet.Workbook;
        var style = workbook.CreateCellStyle();
        style.DataFormat = workbook.CreateDataFormat().GetFormat(format);
        cell.CellStyle = style;
    }
}
