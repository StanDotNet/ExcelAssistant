using FuzzySharp;
using FuzzySharp.PreProcess;
using NPOI.SS.UserModel;

namespace ExcelAssistant;

public abstract class ExcelReader<TObject> : ExcelHelper
{
    ExcelReader(Stream stream) : this(stream, new())
    {
    }

    ExcelReader(Stream stream, ExcelConfiguration configuration) 
        : base(configuration)
    {
        OpenWorkbook(stream);
        OpenSheet();
    }
    
    public List<TObject> Read(CancellationToken cancellationToken = new())
    {
        SetHeaders(sheet.GetRow(sheet.FirstRowNum));
        var records = ReadRecords(sheet, cancellationToken);

        return records;
    }
    
    protected virtual void SetHeaders(IRow row)
    {
        var headersName = GetHeaders<TObject>();
        var cells = row.Cells
            .Where(c => !string.IsNullOrWhiteSpace(c.StringCellValue))
            .ToList();
        
        foreach (var headerName in headersName)
        {
            var cell = cells.FirstOrDefault(c => Fuzz.PartialRatio(c.StringCellValue, headerName, PreprocessMode.Full) >= configuration.MatchingPercentage);
            if (cell != null)
            {
                headers.Add(cell.ColumnIndex, cell.StringCellValue.Trim());
                cells.Remove(cell);
            }
        }
        
        cells.ForEach(c => headers.Add(c.ColumnIndex, c.StringCellValue.Trim()));
    }
    
    protected virtual List<TObject> ReadRecords(ISheet sheet, CancellationToken cancellationToken = new())
    {
        var records = new List<TObject>();
        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                break;
            }

            IRow row = sheet.GetRow(i);
            if (row == null) continue;
            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
            
            records.Add(Map(row));
        }

        return records;
    }

    protected Dictionary<string, string> GetRowData(IRow row)
    {
        var rowData = new Dictionary<string, string>();
        foreach (var kv in headers)
        {
            var cell = row.GetCell(kv.Key);
            rowData.Add(kv.Value, cell?.ToString()?.Trim());
        }

        return rowData;
    }

    protected abstract TObject CreateInstance(Dictionary<string, string> headers);
    
    private TObject Map(IRow row)
    {
        var rowData = GetRowData(row);

        return CreateInstance(rowData);
    }
}