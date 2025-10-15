


using System.Text.Json;
using ExcelAssistant;

using ExcelAssistant;




// Sample data
var reports = new List<Report>()
{
    new("Maki", "test@gmail.com", 100),
    new("Smith", "test1@gmail.com", 200),
    new("Lara", "test2@gmail.com", 450),
};

var customers = new List<Customer>()
{
    new(Guid.NewGuid(), "Alice", "alice@mail.com"),
    new(Guid.NewGuid(), "Bob", "bob@mail.com"),
    new(Guid.NewGuid(), "Charlie", "charlie@mail.com"),
};

using var writer = new ExcelWriter();

writer.AddMapping<Report>(r => r.Name, "Customer Name");
writer.AddMapping<Report>(r => r.Balance, "Account Balance");
writer.AddMapping<Customer>(c => c.Email, "Email Address");

writer.Exclude<Customer>(r => r.Id);

// Build multiple sheets
writer.BuildSheet(reports, "Reports");
writer.BuildSheet(customers, "Customers");

// Create and save a new Excel file with two sheets
using var stream = File.OpenWrite("records.xlsx");
writer.CopyTo(stream);


using var stream2 = File.OpenRead("records.xlsx");
using var reader = new ExcelReader(stream2);

reader.AddMapping<Report>(r => r.Name, "Customer Name");
reader.AddMapping<Report>(r => r.Balance, "Account Balance");
reader.AddMapping<Customer>(c => c.Email, "Email Address");

var reportRecords = reader.Read<Report>();
var customerRecords = reader.Read<Customer>("Customers");


Console.WriteLine();

// Example models
public record Report(string Name, string Email, decimal Balance);
public record Customer(Guid Id, string Name, string Email);

// await using var file = File.OpenRead("/Users/stan/Code/md/temp/scott/Specials2.xlsx");
// using var reader = new ExcelReader(file);
// reader.AddMapping<SpecialReport>(r => r.UID, "AppUserId");
// var records = reader.Read<SpecialReport>();
// var records1 = reader.Read<SpecialReport>("Scott Subaru");
// var records2 = reader.Read<SpecialReport>("Kipp Scott Cadillac");
// var records3 = reader.Read<SpecialReport>("Kipp Scott GMC Buick");
// var records4 = reader.Read<SpecialReport>("Subaru of Cranbrook");




public record SpecialReport(
    string UID,
    string Email,
    string Phone,
    string Type,
    string Title,
    string Description,
    string Details,
    string From,
    string To);


// var reportStr = await File.ReadAllTextAsync("report.json");
//
// var report = JsonSerializer.Deserialize<List<TransactionReport>>(reportStr, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
//
// using var writer = new ExcelWriter();
// writer.AddMapping<TransactionReport>(r => r.CustomerName, "Customer Name");
// writer.AddMapping<TransactionReport>(r => r.SourceName, "Source Name");
// writer.BuildSheet(report, "Rewards Transactions");
//
//
// await using var stream = File.OpenWrite("file.xlsx");
//
// writer.CopyTo(stream);
//
// Console.WriteLine("Done!");
//
//
//
// public record TransactionReport(
//     string CustomerName,
//     decimal Amount,
//     string Source,
//     string SourceName,
//     string Description,
//     string Created)
// {
// };