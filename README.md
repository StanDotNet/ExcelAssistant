# ExcelAssistant

ExcelAssistant is a lightweight and efficient .NET library designed to simplify working with Microsoft Excel files. It allows developers to easily read, write, and manipulate Excel workbooks and worksheets directly within their .NET applications.

## Features

- Read data from Excel files
- Write data to Excel files
- Manipulate worksheets: add, delete, rename, set tab colors
- Customize cells: set values, formats, column widths, and row heights
- Apply styles to cells and worksheets

## Installation

You can install ExcelAssistant using [NuGet](https://www.nuget.org/packages/ExcelAssistant/):

```bash
Install-Package ExcelAssistant
```

### .NET CLI Console

```
dotnet add package ExcelAssistant
```


## Getting Started

### Writing a Excel File

The following example demonstrates how to create an Excel file with multiple worksheets and custom headers:

```csharp
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
    new("Alice", "alice@mail.com"),
    new("Bob", "bob@mail.com"),
    new("Charlie", "charlie@mail.com"),
};

using var writer = new ExcelWriter();

// Define column mappings
writer.AddMapping<Report>(r => r.Name, "Customer Name");
writer.AddMapping<Report>(r => r.Balance, "Account Balance");
writer.AddMapping<Customer>(c => c.Email, "Email Address");

// Build multiple sheets
writer.BuildSheet(reports, "Reports");
writer.BuildSheet(customers, "Customers");

// Save the Excel file
using var stream = File.OpenWrite("records.xlsx");
writer.CopyTo(stream);


// Example models
public record Report(string Name, string Email, decimal Balance);
public record Customer(string Name, string Email);

```

### Reading from an Excel File

You can also read data back into C# models. Column headers in the Excel file must match the defined mappings.

```csharp
using var stream = File.OpenRead("records.xlsx");
using var reader = new ExcelReader(stream);

reader.AddMapping<Report>(r => r.Name, "Customer Name");
reader.AddMapping<Report>(r => r.Balance, "Account Balance");
reader.AddMapping<Customer>(c => c.Email, "Email Address");

// if sheet nane is not set it will read firs sheet
var reportRecords = reader.Read<Report>();
var customerRecords = reader.Read<Customer>("Customers");
```

## License

ExcelAssistant is released under the [MIT License](https://github.com/StanDotNet/ExcelAssistant/blob/main/LICENSE.md).