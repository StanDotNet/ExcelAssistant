using System.Reflection;
using FuzzySharp;
using FuzzySharp.PreProcess;
using NPOI.SS.UserModel;

namespace ExcelAssistant;

public class ExcelReader: ExcelHelper
{
    public ExcelReader(Stream stream) : this(stream, new())
    {
    }

    public ExcelReader(Stream stream, ExcelConfiguration configuration) 
        : base(configuration)
    {
        OpenWorkbook(stream);
        OpenSheet();
    }
    
    public List<TObject> Read<TObject>(CancellationToken cancellationToken = new()) where TObject : class
    {
        SetHeaders<TObject>(sheet.GetRow(sheet.FirstRowNum));
        var records = ReadRecords(cancellationToken);
        var instances = records.Select(CreateInstance<TObject>).ToList();

        return instances;
    }
    
    public IEnumerable<Dictionary<string,string>> Read(CancellationToken cancellationToken = new())
    {
        SetHeaders(sheet.GetRow(sheet.FirstRowNum));
        var records = ReadRecords(cancellationToken);
        foreach (var record in records)
        {
            yield return record;
        }
    }

    private void SetHeaders<TObject>(IRow row)
    {
        var headersName = GetHeaders<TObject>();
        var cells = row.Cells
            .Where(c => !string.IsNullOrWhiteSpace(c.StringCellValue))
            .ToList();

        headers = cells
            .Select(c => KeyValuePair.Create(c.ColumnIndex,
                configuration.HumanReadableHeaders.FirstOrDefault(h => h.Value == c.StringCellValue).Key
                ?? headersName.FirstOrDefault(h => Fuzz.PartialRatio(c.StringCellValue, h, PreprocessMode.Full)
                                                   >= configuration.MatchingPercentage) ?? string.Empty))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
    }
    
    private void SetHeaders(IRow row)
    {
        headers = row.Cells
            .Where(c => !string.IsNullOrWhiteSpace(c.StringCellValue))
            .ToDictionary(c => c.ColumnIndex, c => c.StringCellValue);
    }
    
    private IEnumerable<Dictionary<string,string>> ReadRecords(CancellationToken cancellationToken = new())
    {
        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }

            IRow row = sheet.GetRow(i);
            if (row == null) continue;
            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

            yield return GetRowData(row);
        }
    }
    
    private Dictionary<string, string> GetRowData(IRow row)
    {
        var rowData = new Dictionary<string, string>();
        foreach (var kv in headers)
        {
            var cell = row.GetCell(kv.Key);
            rowData.Add(kv.Value, cell?.ToString()?.Trim());
        }

        return rowData;
    }
    
    private TObject CreateInstance<TObject>(Dictionary<string, string> rowData) where TObject : class
    {
        var type = typeof(TObject);
        var constructor = type.GetConstructors()
            .OrderByDescending(c => c.GetParameters().Length)
            .First();
        
        var parameters = constructor
            .GetParameters()
            .Select(p => GetConstructorParameter(p, rowData))
            .ToArray();
        
        var instance = (TObject)Activator.CreateInstance(typeof(TObject), parameters);
        
        if (instance != null)
        {
            var ctorParameterNames = constructor.GetParameters().Select(p => p.Name);

            instance.GetType()
                .GetProperties()
                .Where(p => !ctorParameterNames.Contains(p.Name))
                .Where(p => p.SetMethod != null)
                .ToList()
                .ForEach(property =>
                {
                    var parameter = GetParameter(property.PropertyType, rowData.GetValueOrDefault(property.Name));
                    property.SetValue(instance, parameter);
                });
        }

        return instance;
    }

    private object? GetConstructorParameter(ParameterInfo parameter, Dictionary<string, string> rowData)
    {
        var type = parameter.ParameterType;
        var value = rowData.GetValueOrDefault(parameter.Name);
        if (string.IsNullOrWhiteSpace(value))
        {
            return parameter.HasDefaultValue && parameter.DefaultValue != null
                ? parameter.DefaultValue
                : (type.IsValueType ? Activator.CreateInstance(type) : null);
        }

        return GetParameter(type, value);
    }

    private object? GetParameter(Type type, string value) => type switch
    {
        _ when string.IsNullOrWhiteSpace(value) => null,
        _ when type == typeof(string) => value,
        _ when type == typeof(byte) || type == typeof(byte?) => byte.Parse(value),
        _ when type == typeof(short) || type == typeof(short?) => short.Parse(value),
        _ when type == typeof(int) || type == typeof(int?) => int.Parse(value),
        _ when type == typeof(long) || type == typeof(long?) => long.Parse(value),
        _ when type == typeof(double) || type == typeof(double?) => double.Parse(value),
        _ when type == typeof(float) || type == typeof(float?) => float.Parse(value),
        _ when type == typeof(decimal) || type == typeof(decimal?) => decimal.Parse(value),
        _ when type == typeof(Guid) || type == typeof(Guid?) => Guid.Parse(value),
        _ when type == typeof(DateTime) || type == typeof(DateTime?) => DateTime.Parse(value),
        _ when type == typeof(TimeSpan) || type == typeof(TimeSpan?) => TimeSpan.Parse(value),
        _ when type == typeof(DateOnly) || type == typeof(DateOnly?) => DateOnly.Parse(value),
        _ when type == typeof(TimeOnly) || type == typeof(TimeOnly?) => TimeOnly.Parse(value),
        _ =>  throw new NotSupportedException($"The configuration property type {type.Name} is not supported")
    };
}