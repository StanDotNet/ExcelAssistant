# ExcelAssistant

ExcelAssistant is a .NET library that provides a simple and efficient way to work with Microsoft Excel files. With ExcelAssistant, you can easily read, write, and manipulate Excel files in your .NET applications.

## Features

- Read data from Excel files
- Write data to Excel files
- Manipulate Excel files (add, delete, rename sheets, set sheet color, etc.)
- Set cell values and formats
- Set column width and row height
- Apply styles to cells and sheets

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

Let's look at how we can create and write excel files. In our example, we will use the simple "Report" record model, but it can be any other c# class.

```csharp
using ExcelAssistant;

//Create a list of test data
var records = new List<Report>()
{
    new("Maki", "test@gmail.com", 100),
    new("Smith", "test1@gmail.com", 200),
    new("Lara", "test2@gmail.com", 450),
};

//Open or create the file for reading (the file will be in the project output directory)
using var stream = File.OpenWrite("records.xls");

//Create an instance of ExcelWriter.
using var writer = new ExcelWriter();
//Write records into the file
writer.WriteRecords(stream, records);

//"Report" model 
public record Report(string Name, string Email, decimal Balance);
```

Also, we can change our output file header by adding configuration:

```csharp
var config = new ExcelConfiguration
{
    HumanReadableHeaders = new Dictionary<string, string>()
    {
        {nameof(Report.Name), "Customer Name"},
        {nameof(Report.Email), "Customer Email"},
    }
};

using var stream = File.OpenWrite("records.xls");
using var writer = new ExcelWriter(config);
writer.WriteRecords(stream, records);
```

Reading files where table headers are the same as c# model properties:


```csharp
using var stream = File.OpenRead("records.xls");
using var reader = new ExcelReader(stream);
var records = reader.Read<Report>();
```

Now reading files using the configuration mentioned above

```csharp
using var stream = File.OpenRead("records.xls");
using var reader = new ExcelReader(stream, config);
var records = reader.Read<Report>();
```

The library provides the opportunity for manual reading. Method Read return IEnumerable<Dictionary<string,string>> where key is the column name and value is the data for the specific row.

```csharp
using var stream = File.OpenRead("records.xls");
using var reader = new ExcelReader(stream, config);
var records = reader
    .Read()
    .Select(rowData => new Report(
        rowData[nameof(Report.Name)],
        rowData[nameof(Report.Email)],
        decimal.Parse(rowData[nameof(Report.Balance)])
    ))
    .ToList();
```

## License

ExcelAssistant is released under the [MIT License](https://github.com/StanDotNet/ExcelAssistant/blob/main/LICENSE.md).