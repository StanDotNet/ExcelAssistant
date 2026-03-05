namespace ExcelAssistant;

public class ExcelConfiguration
{
    /// <summary>
    /// Type of Excel file
    /// </summary>
    public ExcelType? ExcelType { get; set; } = ExcelAssistant.ExcelType.xlsx;
    
    /// <summary>
    /// Sheet Name. By default will be reading first sheet.
    /// </summary>
    public string SheetName { get; set; }

    /// <summary>
    /// Set it if you want to read data started not from first column.
    /// </summary>
    public string MainColumnName { get; set; }
    
    /// <summary>
    /// Headers like they displayed in excel file and c# field equivalent
    /// </summary>
    public Dictionary<string, string> HumanReadableHeaders { get; set; } = new();

    public List<string> Exclusions { get; set; } = [];

    /// <summary>
    /// Percent Matching excel headers with c# model fields. 
    /// </summary>
    public int MatchingPercentage { get; set; } = 90;

    /// <summary>
    /// The coefficient shows how much space is needed for one symbol.
    /// </summary>
    public int ColumnSizeCoefficient { get; set; } = 300;

    /// <summary>
    /// Minimum column width in characters. Prevents columns from being too narrow.
    /// </summary>
    public int MinColumnSize { get; set; } = 10;

    /// <summary>
    /// Maximum column width in characters. Prevents very long content from squeezing other columns.
    /// </summary>
    public int MaxColumnSize { get; set; } = 60;
}