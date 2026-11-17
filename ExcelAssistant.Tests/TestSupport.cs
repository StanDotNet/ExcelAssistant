using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelAssistant;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelAssistant.Tests;

internal static class TestSupport
{
    [ModuleInitializer]
    internal static void UseInvariantCulture()
    {
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
    }

    public static MemoryStream Write<T>(ExcelType type, List<T> records, Action<ExcelWriter>? setup = null, string sheetName = "Sheet")
        where T : class
    {
        using var writer = new ExcelWriter(new ExcelConfiguration { ExcelType = type });
        setup?.Invoke(writer);
        writer.BuildSheet(records, sheetName);

        var stream = new MemoryStream();
        writer.CopyTo(stream);
        return stream;
    }

    public static ExcelReader Reader(Stream stream, ExcelType type, int matchingPercentage = 90) =>
        new(stream, new ExcelConfiguration { ExcelType = type, MatchingPercentage = matchingPercentage });

    public static MemoryStream Sheet(ExcelType type, string[] headers, params Action<IRow>[] rows)
    {
        IWorkbook workbook = type == ExcelType.xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet");

        var headerRow = sheet.CreateRow(0);
        for (var i = 0; i < headers.Length; i++)
        {
            headerRow.CreateCell(i).SetCellValue(headers[i]);
        }

        for (var i = 0; i < rows.Length; i++)
        {
            rows[i](sheet.CreateRow(i + 1));
        }

        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    public static IDisposable CommaDecimalCulture()
    {
        var culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        culture.NumberFormat.NumberDecimalSeparator = ",";
        culture.NumberFormat.NumberGroupSeparator = " ";
        return new CultureScope(culture);
    }

    private sealed class CultureScope : IDisposable
    {
        private readonly CultureInfo previous = CultureInfo.CurrentCulture;

        public CultureScope(CultureInfo culture) => CultureInfo.CurrentCulture = culture;

        public void Dispose() => CultureInfo.CurrentCulture = previous;
    }
}
