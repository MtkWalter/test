using System.Globalization;
using System.Text;

namespace ExcelInsights.Core;

public static class CsvExporter
{
    public static void Write(string path, IEnumerable<Summary> summaries, string file, string sheet, string column, string filter)
    {
        using var writer = new StreamWriter(path, false, new UTF8Encoding(true));
        writer.WriteLine("檔案,工作表,數值欄位,篩選條件,群組,資料筆數,有效數值,空白,非數值,總和,平均,最小值,最大值");
        foreach (var s in summaries)
        {
            object?[] values = [file, sheet, column, filter, s.Group, s.Rows, s.NumericCount, s.BlankCount,
                s.InvalidCount, s.Sum, s.Average, s.Minimum, s.Maximum];
            writer.WriteLine(string.Join(",", values.Select(Escape)));
        }
    }

    private static string Escape(object? value)
    {
        var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        if (value is string && text.TrimStart().FirstOrDefault() is '=' or '+' or '-' or '@') text = "'" + text;
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }
}
