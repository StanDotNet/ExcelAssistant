


using System.Text.Json;
using ExcelAssistant;

var reportStr = await File.ReadAllTextAsync("report.json");

var report = JsonSerializer.Deserialize<List<TransactionReport>>(reportStr, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

using var writer = new ExcelWriter();
writer.AddMapping<TransactionReport>(r => r.CustomerName, "Customer Name");
writer.AddMapping<TransactionReport>(r => r.SourceName, "Source Name");
writer.BuildSheet(report, "Rewards Transactions");


await using var stream = File.OpenWrite("file.xlsx");

writer.CopyTo(stream);

Console.WriteLine("Done!");



public record TransactionReport(
    string CustomerName,
    decimal Amount,
    string Source,
    string SourceName,
    string Description,
    string Created)
{
};