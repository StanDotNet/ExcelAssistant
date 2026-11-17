namespace ExcelAssistant.Tests;

public record Vehicle(
    string Model,
    byte Seats,
    short Doors,
    int Year,
    long Mileage,
    float Weight,
    double Rating,
    decimal Price,
    Guid Vin,
    TimeSpan Warranty,
    TimeOnly OpensAt,
    int? Owners,
    decimal? Discount,
    string? Notes);

public record Amount(decimal Value);

public record Flag(string Code, bool Enabled);

public record Stamp(DateTime At);

public record Day(DateOnly On);

public record Customer(Guid Id, string Name, string Email);

public record Report(string CustomerName, string Email, decimal Balance);

public class Settable
{
    public string? Title { get; set; }
    public int? Count { get; set; }
    public decimal Total { get; set; }
    public string Computed => $"{Title}:{Count}";
}
