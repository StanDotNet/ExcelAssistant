# ExcelAssistant

[![NuGet](https://img.shields.io/nuget/v/ExcelAssistant.svg)](https://www.nuget.org/packages/ExcelAssistant/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/ExcelAssistant.svg)](https://www.nuget.org/packages/ExcelAssistant/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/StanDotNet/ExcelAssistant/blob/main/LICENSE.md)

A lightweight .NET library for reading and writing Excel files (.xls / .xlsx) with minimal boilerplate. Map Excel columns to C# models using reflection and fuzzy header matching — no XML config or attributes required.

## Features

- Read Excel files into typed C# models or plain dictionaries
- Write any `List<T>` to Excel with styled headers
- Multi-sheet support — read and write multiple sheets in one file
- Fuzzy column header matching — tolerates typos and renamed columns
- Custom column name mappings via fluent API
- Exclude properties from output
- Supports both `.xls` and `.xlsx` formats
- Configurable column widths (min, max, coefficient)

## Installation

Package Manager:
```
Install-Package ExcelAssistant
```

.NET CLI:
```
dotnet add package ExcelAssistant
```

## Quick Start

### Writing

```csharp
using ExcelAssistant;

var reports = new List<Report>
{
    new("Alice", "alice@mail.com", 1200.50m),
    new("Bob",   "bob@mail.com",   850.00m),
};

using var writer = new ExcelWriter();

// Map C# property names to human-readable Excel column headers
writer.AddMapping<Report>(r => r.Name,    "Customer Name");
writer.AddMapping<Report>(r => r.Balance, "Account Balance");

writer.BuildSheet(reports, "Reports");

using var stream = File.OpenWrite("output.xlsx");
writer.CopyTo(stream);

public record Report(string Name, string Email, decimal Balance);
```

### Reading

```csharp
using var stream = File.OpenRead("output.xlsx");
using var reader = new ExcelReader(stream);

// Same mappings tell the reader which Excel column maps to which property
reader.AddMapping<Report>(r => r.Name,    "Customer Name");
reader.AddMapping<Report>(r => r.Balance, "Account Balance");

// Reads first sheet by default
var reports = reader.Read<Report>();

// Read a specific sheet by name
var customers = reader.Read<Customer>("Customers");
```

## Multi-Sheet Example

```csharp
using var writer = new ExcelWriter();

writer.AddMapping<Report>(r => r.Name,      "Customer Name");
writer.AddMapping<Customer>(c => c.Email,   "Email Address");
writer.Exclude<Customer>(c => c.Id);          // omit from output

writer.BuildSheet(reports,   "Reports");
writer.BuildSheet(customers, "Customers");

using var stream = File.OpenWrite("records.xlsx");
writer.CopyTo(stream);

// ---

using var stream2 = File.OpenRead("records.xlsx");
using var reader  = new ExcelReader(stream2);

reader.AddMapping<Report>(r => r.Name,    "Customer Name");
reader.AddMapping<Customer>(c => c.Email, "Email Address");

var reportRecords   = reader.Read<Report>();
var customerRecords = reader.Read<Customer>("Customers");
```

## Reading Without a Model

If you don't have a typed model, read each row as a plain dictionary:

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

foreach (var row in reader.Read())
{
    Console.WriteLine(row["Customer Name"]);
}
```

## Fuzzy Header Matching

By default the reader tolerates up to 10% difference between Excel column names and C# property names. This means minor variations like `"CustomerName"` vs `"Customer Name"` or `"customer_name"` are matched automatically.

Control the threshold via `MatchingPercentage` in configuration (0–100, default 90):

```csharp
var config = new ExcelConfiguration { MatchingPercentage = 80 };
using var reader = new ExcelReader(stream, config);
```

## Excluding Properties

```csharp
// Using typed expression
writer.Exclude<Customer>(c => c.InternalId);

// Using property name as string
writer.Exclude("InternalId");
```

## Configuration Reference

All options are set via `ExcelConfiguration`:

```csharp
var config = new ExcelConfiguration
{
    ExcelType           = ExcelType.xlsx,  // xlsx (default) or xls
    SheetName           = "Sheet1",        // target sheet; null = first sheet
    MatchingPercentage  = 90,              // fuzzy match threshold (0–100)
    ColumnSizeCoefficient = 300,           // width units per character
    MinColumnSize       = 10,              // minimum column width in characters
    MaxColumnSize       = 60,              // maximum column width in characters
};

using var writer = new ExcelWriter(config);
using var reader = new ExcelReader(stream, config);
```

| Property | Default | Description |
|---|---|---|
| `ExcelType` | `xlsx` | File format: `xlsx` or `xls` |
| `SheetName` | `null` | Sheet to read/write; `null` uses the first/active sheet |
| `MatchingPercentage` | `90` | Minimum fuzzy match score for header detection (0–100) |
| `ColumnSizeCoefficient` | `300` | Width units per character (affects all columns) |
| `MinColumnSize` | `10` | Minimum column width in characters |
| `MaxColumnSize` | `60` | Maximum column width in characters |
| `HumanReadableHeaders` | `{}` | Manual column name mappings (set via `AddMapping`) |
| `Exclusions` | `[]` | Properties to exclude from output (set via `Exclude`) |

## Supported Property Types

`string`, `byte`, `short`, `int`, `long`, `float`, `double`, `decimal`, `bool`, `Guid`, `DateTime`, `DateOnly`, `TimeSpan`, `TimeOnly` — and their nullable equivalents.

## Content Type

Use `ContentType` to set the correct HTTP response header when serving files from an API:

```csharp
Response.ContentType = writer.ContentType;
// "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" for .xlsx
// "application/vnd.ms-excel" for .xls
```

## License

MIT — see [LICENSE.md](https://github.com/StanDotNet/ExcelAssistant/blob/main/LICENSE.md).
