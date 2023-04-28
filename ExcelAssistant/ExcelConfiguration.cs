namespace ExcelAssistant;

public class ExcelConfiguration
{
    /// <summary>
    /// Type of Excel file
    /// </summary>
    public ExcelType ExcelType { get; init; }
    
    /// <summary>
    /// Sheet Name. By default will be reading first sheet.
    /// </summary>
    public string SheetName { get; init; }

    /// <summary>
    /// Set it if you want to read data started not from first column.
    /// </summary>
    public string MainColumnName { get; init; }
    
    /// <summary>
    /// Headers like they displayed in excel file and c# field equivalent
    /// </summary>
    public Dictionary<string, string> HumanReadableHeaders { get; init; } = new();

    /// <summary>
    /// Percent Matching excel headers with c# model fields. 
    /// </summary>
    public int MatchingPercentage { get; init; } = 80;
}