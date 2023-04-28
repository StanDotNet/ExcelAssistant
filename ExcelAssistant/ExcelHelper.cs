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

        switch (configuration.ExcelType)
        {
            case ExcelType.xls:
                workbook = new HSSFWorkbook(stream);
                break;
            case ExcelType.xlsx:
                workbook = new XSSFWorkbook(stream);
                break;
            default: throw new Exception("Unsupported excel type");
        }
    }

    protected void OpenSheet()
    {
        sheet = string.IsNullOrWhiteSpace(configuration.SheetName)
            ? workbook.GetSheetAt(0)
            : workbook.GetSheet(configuration.SheetName);
    }

    protected void CreateSheet()
    {
        sheet = string.IsNullOrWhiteSpace(configuration.SheetName)
            ? workbook.CreateSheet()
            : workbook.CreateSheet(configuration.SheetName);
    }

    protected int GetColumnSize(int maxlengthContent) =>
        maxlengthContent * 270;
    
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