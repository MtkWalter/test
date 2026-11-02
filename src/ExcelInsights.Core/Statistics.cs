using System.Data;
using System.Globalization;

namespace ExcelInsights.Core;

public sealed record Summary(string Group, int Rows, int NumericCount, int BlankCount, int InvalidCount,
    double? Sum, double? Average, double? Minimum, double? Maximum);

public static class Statistics
{
    public static bool IsBlank(object? value) => value is null or DBNull || value is string s && string.IsNullOrWhiteSpace(s);

    public static bool TryNumber(object? value, out double number)
    {
        number = 0;
        if (value is DateTime or bool || IsBlank(value)) return false;
        if (value is string text)
        {
            // A fixed culture avoids interpreting the same workbook differently on another PC.
            if (!double.TryParse(text.Trim(), NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out number)) return false;
        }
        else if (value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal)
            number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
        else return false;
        return double.IsFinite(number);
    }

    public static IReadOnlyList<DataRow> Filter(DataTable table, string? column, string query)
    {
        if (string.IsNullOrWhiteSpace(query)) return table.Rows.Cast<DataRow>().ToArray();
        if (column is null || !table.Columns.Contains(column)) throw new ArgumentException("請選擇有效的篩選欄位。");
        return table.Rows.Cast<DataRow>().Where(row =>
            (Convert.ToString(row[column], CultureInfo.InvariantCulture) ?? "")
            .Contains(query, StringComparison.OrdinalIgnoreCase)).ToArray();
    }

    public static IReadOnlyList<Summary> Analyze(IEnumerable<DataRow> rows, string numericColumn, string? groupColumn = null)
    {
        return rows.GroupBy(row => groupColumn is null ? "全部" :
                IsBlank(row[groupColumn]) ? "（空白）" : Convert.ToString(row[groupColumn], CultureInfo.InvariantCulture)!)
            .Select(group =>
            {
                var count = 0;
                var blank = 0;
                var numeric = new List<double>();
                foreach (var row in group)
                {
                    count++;
                    var value = row[numericColumn];
                    if (IsBlank(value)) blank++;
                    else if (TryNumber(value, out var number)) numeric.Add(number);
                }
                double? sum = numeric.Count > 0 ? numeric.Sum() : null;
                // Prevent overflow from being exported as a misleading Infinity value.
                if (sum.HasValue && !double.IsFinite(sum.Value)) throw new OverflowException("數值總和超出可計算範圍。");
                return new Summary(group.Key, count, numeric.Count, blank, count - blank - numeric.Count,
                    sum, numeric.Count > 0 ? sum / numeric.Count : null,
                    numeric.Count > 0 ? numeric.Min() : null, numeric.Count > 0 ? numeric.Max() : null);
            }).ToArray();
    }
}
