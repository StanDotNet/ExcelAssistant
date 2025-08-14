using NPOI.SS.UserModel;

namespace ExcelAssistant;

public class ExcelWriter : ExcelHelper
{
    public ExcelWriter() : this(new ExcelConfiguration() {ExcelType = ExcelType.xlsx})
    {
    }

    public ExcelWriter(ExcelConfiguration configuration) : base(new ExcelConfiguration
    {
        ExcelType = configuration.ExcelType,
        SheetName = configuration.SheetName,
        HumanReadableHeaders = configuration.HumanReadableHeaders,
        MainColumnName = configuration.MainColumnName
    })
    {
        CreateWorkbook();
    }

    public void BuildSheet<TObject>(List<TObject> records, string sheetName = null, CancellationToken cancellationToken = new())
        where TObject : class
    {
        CreateSheet(sheetName);
        SetHeaders(records);
        SetDefaultColumnStyle();

        if (records?.Count > 0)
        {
            Write(records, cancellationToken);
            SetTableStyle(records.Count);
        }
    }

    public Stream WriteRecords<TObject>(Stream stream, List<TObject> records, CancellationToken cancellationToken = new())
        where TObject : class
    {
        CreateSheet();
        SetHeaders(records);
        SetDefaultColumnStyle();

        if (records?.Count > 0)
        {
            Write(records, cancellationToken);
            SetTableStyle(records.Count);

            workbook.Write(stream, true);
            stream.Position = 0;
        }

        return stream;
    }

    public void CopyTo(Stream stream)
    {
        workbook.Write(stream, true);
        stream.Position = 0;
    }

    protected virtual void Write<TObject>(List<TObject> records, CancellationToken cancellationToken = new())
        where TObject : class
    {
        var rowIndex = 0;
        var columnSize = headers.ToDictionary(kv => kv.Key, kv => kv.Value.Length);

        records.ForEach(r =>
        {
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            rowIndex++;

            IRow row = sheet.CreateRow(rowIndex);

            var values = GetValues(r);

            foreach (var keyValuePair in headers)
            {
                var cellValue = values.First(v => v.Name == keyValuePair.Value);
                
                row.CreateCell(keyValuePair.Key).SetCellValue(cellValue.Value).SetCellType(cellValue.CellType);

                columnSize[keyValuePair.Key] = new List<int>
                {
                    columnSize[keyValuePair.Key],
                    keyValuePair.Value.Length,
                    cellValue.Value.Length,
                    configuration.HumanReadableHeaders.GetValueOrDefault(keyValuePair.Value)?.Length ?? 0
                }.Max();
            }
        });

        foreach (var keyValuePair in columnSize)
        {
            sheet.SetColumnWidth(keyValuePair.Key, GetColumnSize(keyValuePair.Value));
        }
    }

    protected virtual void SetHeaders<TObject>(List<TObject> records) where TObject : class
    {
        var key = 0;
        var headersName = GetHeaders(records);
        IRow headersRow = sheet.CreateRow(key);
        headers.Clear();

        headersName.ForEach(n =>
        {
            headers.Add(key, n);
            var cell = headersRow.CreateCell(key);
            cell.CellStyle = GetBoldCellStyle(IndexedColors.Grey80Percent.Index, first: key == 0, last: key == headersName.Count - 1);
            cell.SetCellValue(configuration.HumanReadableHeaders.GetValueOrDefault(n) ?? n);
            key++;
        });
    }

    protected virtual List<string> GetHeaders<TObject>(List<TObject> records) where TObject : class =>
        GetHeaders<TObject>();
    
    protected virtual List<CellValue> GetValues<TObject>(TObject record) where TObject : class =>
        record?
            .GetType()
            .GetProperties()
            .Select(p => 
                new CellValue(p.Name, p.GetValue(record)?.ToString() ?? string.Empty, GetCellType(p.GetValue(record))))
            .ToList()
        ?? new List<CellValue>();
    
    private static CellType GetCellType(object? value)
    {
        if (value == null) return CellType.Blank;

        return value switch
        {
            bool => CellType.Boolean,
            sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal
                => CellType.Numeric,
            string s when s.StartsWith("=") => CellType.Formula,
            string => CellType.String,
            _ => CellType.Unknown
        };
    }
}

public record CellValue(string Name, string Value, CellType CellType);