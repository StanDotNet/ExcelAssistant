using System.Linq.Expressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelAssistant;

public abstract class ExcelHelper : IDisposable
{
    /// <summary>
    /// Position header in the file and c# field in the model. By default will be analyzed using reflection.
    /// </summary>
    protected Dictionary<int, string> headers = new();

    protected readonly ExcelConfiguration configuration;

    protected IWorkbook workbook;
    protected ISheet? sheet;
    
    protected ExcelHelper() : this(new ExcelConfiguration())
    {
    }
    
    protected ExcelHelper(ExcelConfiguration configuration)
    {
        this.configuration = configuration;
    }
    
    protected void OpenWorkbook(Stream stream)
    {
        stream.Position = 0;

        if (configuration.ExcelType == null && stream is FileStream fileStream)
        {
            configuration.ExcelType = fileStream.Name.EndsWith(".xlsx") 
                ? ExcelType.xlsx 
                : ExcelType.xls;
        }
        
        CreateWorkbook(stream);
    }
    
    protected void CreateWorkbook(Stream? stream = null, ExcelType excelType = ExcelType.xls)
    {
        configuration.ExcelType ??= excelType;
        
        switch (configuration.ExcelType)
        {
            case ExcelType.xls:
                workbook = stream?.Length > 0 ? new HSSFWorkbook(stream) : new HSSFWorkbook();
                break;
            case ExcelType.xlsx:
                workbook = stream?.Length > 0 ? new XSSFWorkbook(stream) : new XSSFWorkbook();
                break;
            default: throw new Exception("Unsupported excel type");
        }
    }

    public void OpenSheet(string? sheetName = null)
    {
        sheetName ??= configuration.SheetName;
        sheet = string.IsNullOrWhiteSpace(sheetName)
            ? workbook.GetSheetAt(workbook.ActiveSheetIndex)
            : workbook.GetSheet(sheetName);
    }

    public void CreateSheet(string? sheetName = null)
    {
        sheetName ??= configuration.SheetName;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            sheet = workbook.CreateSheet(sheetName);
        }
        else if (sheet == null)
        {
            sheet = workbook.CreateSheet();
        }
    }

    public void AddMapping<TObject>(Expression<Func<TObject, object>> propertyExpression, string columnName)
    {
        var propertyName = GetPropertyName(propertyExpression);

        if (!string.IsNullOrWhiteSpace(propertyName))
        {
            configuration.HumanReadableHeaders.Add(propertyName, columnName);
        }
    }
    
    private string? GetPropertyName<TObject>(Expression<Func<TObject, object>> propertyExpression) =>
        propertyExpression.Body switch
        {
            MemberExpression memberExpression => memberExpression.Member.Name,
            UnaryExpression unaryExpression => ((MemberExpression)unaryExpression.Operand).Member.Name,
            _ => null
        };
    
    protected int GetColumnSize(int maxContentLength) =>
        maxContentLength * configuration.ColumnSizeCoefficient;
    
    protected virtual ICellStyle GetBoldCellStyle(short? color = null, bool first = false, bool last = false)
    {
        var style = sheet.Workbook.CreateCellStyle();
        
        if (color.HasValue)
        {
            if (first)
            {
                style.BorderTop = BorderStyle.Thick;
            }

            style.BorderRight = BorderStyle.Thin;
            style.RightBorderColor = IndexedColors.Grey25Percent.Index;

            if (last)
            {
                style.RightBorderColor = IndexedColors.Black.Index;
                style.BorderRight = BorderStyle.Medium;
            }

            style.FillBackgroundColor = color.Value;
            style.FillForegroundColor = color.Value;

            style.FillPattern = FillPattern.SolidForeground;
        }
        
        var font = sheet.Workbook.CreateFont();
        font.IsBold = true;
        
        if (color.HasValue)
        {
            font.Color = IndexedColors.White.Index;
        }
        
        style.SetFont(font);

        return style;
    }
    
    protected virtual void SetDefaultColumnStyle()
    {
        //leave default style;
    }
    
    protected virtual void SetTableStyle( int recordsCount, int? columnIndex = null, bool leftBorder = true, bool rightBorder = true)
    {
        if (!columnIndex.HasValue)
        {
            //leave default style;
            return;
        }

        for (int i = sheet.FirstRowNum; i <= recordsCount; i++)
        {
            var row = sheet.GetRow(i) ?? sheet.CreateRow(i);

            var isHeader = i == sheet.FirstRowNum;
            var isLastRow = i == recordsCount;

            var cell = row.GetCell(columnIndex.Value) ?? row.CreateCell(columnIndex.Value);
            var style = isHeader ? cell.CellStyle : sheet.Workbook.CreateCellStyle();

            if (leftBorder)
            {
                style.BorderLeft = BorderStyle.Medium;
            }

            if (rightBorder)
            {
                style.BorderRight = BorderStyle.Medium;
            }

            if (isLastRow)
            {
                style.BorderBottom = BorderStyle.Medium;
            }

            cell.CellStyle = style;
        }
    }
    
    protected virtual List<string> GetHeaders<TObject>() =>
        typeof(TObject).GetProperties().Select(p => p.Name).ToList();

    public void Dispose()
    {
        workbook?.Close();
        workbook?.Dispose();
    }
}