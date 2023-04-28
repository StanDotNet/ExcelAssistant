using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelAssistant;

public class ExcelWriter : ExcelHelper
{
    public ExcelWriter() : this(new ExcelConfiguration())
    {
    }

    public ExcelWriter(ExcelConfiguration configuration) : base(new ExcelConfiguration
    {
        //Currently only xls type works properly in NPOI library
        ExcelType = ExcelType.xls,
        SheetName = configuration.SheetName,
        HumanReadableHeaders = configuration.HumanReadableHeaders,
        MainColumnName = configuration.MainColumnName
        
    })
    {
        //Currently only xls type works properly in NPOI library
        workbook = new HSSFWorkbook();
        CreateSheet();
    }
    
    public Stream WriteRecords<TObject>(Stream stream, List<TObject> records, CancellationToken cancellationToken = new())
    {
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
    
    protected virtual void Write<TObject>(List<TObject> records, CancellationToken cancellationToken = new())
    {
        var rowIndex = 0;
        var columnSize = headers.ToDictionary(kv => kv.Key, kv => kv.Value.Length + 1);
        
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
                var value = values[keyValuePair.Value];
                row.CreateCell(keyValuePair.Key).SetCellValue(value);
                
                if (!string.IsNullOrEmpty(value) && columnSize[keyValuePair.Key] < value.Length)
                {
                    columnSize[keyValuePair.Key] = value.Length;
                }
            }
        });
        
        foreach (var keyValuePair in columnSize)
        {
            sheet.SetColumnWidth(keyValuePair.Key, GetColumnSize(keyValuePair.Value));
        }
    }
    
    protected virtual void SetHeaders<TObject>(List<TObject> records)
    {
        var key = 0;
        var headersName = GetHeaders(records);
        IRow headersRow = sheet.CreateRow(key);
        
        headersName.ForEach(n =>
        {
            headers.Add(key, n);
            var cell = headersRow.CreateCell(key);
            cell.CellStyle = GetBoldCellStyle(IndexedColors.Grey80Percent.Index, first: key == 0, last: key == headersName.Count - 1);
            cell.SetCellValue(configuration.HumanReadableHeaders.GetValueOrDefault(n) ?? n);
            key++;
        });
    }
    
    protected virtual List<string> GetHeaders<TObject>(List<TObject> records) =>
        GetHeaders<TObject>();
    
    protected virtual Dictionary<string, string> GetValues<TObject>(TObject record) =>
        record
            ?.GetType()
            .GetProperties()
            .ToDictionary(k => k.Name, v => v.GetValue(record)?.ToString() ?? string.Empty) 
        ?? new();
}