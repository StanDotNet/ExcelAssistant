namespace ExcelAssistant;

public class ExcelConfiguration
{
    /// <summary>
    /// Type of Excel file
    /// </summary>
    public ExcelType? ExcelType { get; set; }
    
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

    /// <summary>
    /// Percent Matching excel headers with c# model fields. 
    /// </summary>
    public int MatchingPercentage { get; set; } = 80;

    /// <summary>
    /// The coefficient shows how much space is needed for one symbol.
    /// </summary>
    public int ColumnSizeCoefficient { get; set; } = 300;
}