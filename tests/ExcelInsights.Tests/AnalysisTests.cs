using System.Data;
using System.Globalization;
using System.Text;
using ExcelInsights.Core;
using Xunit;

namespace ExcelInsights.Tests;

public class AnalysisTests
{
    [Fact]
    public void FilterThenGroupCountsOnlyValidNumbers()
    {
        var table = new DataTable();
        table.Columns.Add("部門", typeof(string));
        table.Columns.Add("金額", typeof(object));
        table.Rows.Add("Taipei", 10d);
        table.Rows.Add("Taipei", "1,200.5");
        table.Rows.Add("Taipei", DBNull.Value);
        table.Rows.Add("Taipei", "錯誤");
        table.Rows.Add("Taipei", new DateTime(2026, 1, 1));
        table.Rows.Add("Tainan", -5d);
        table.Rows.Add("Kaohsiung", 500d);
        var rows = Statistics.Filter(table, "部門", "TAI");
        var summary = Statistics.Analyze(rows, "金額", "部門");
        Assert.Equal(2, summary.Count);
        var taipei = summary.Single(s => s.Group == "Taipei");
        Assert.Equal(5, taipei.Rows);
        Assert.Equal(2, taipei.NumericCount);
        Assert.Equal(1, taipei.BlankCount);
        Assert.Equal(2, taipei.InvalidCount);
        Assert.Equal(1210.5, taipei.Sum);
        Assert.Equal(605.25, taipei.Average);
        Assert.Equal(10, taipei.Minimum);
        Assert.Equal(1200.5, taipei.Maximum);
        Assert.Equal(-5, summary.Single(s => s.Group == "Tainan").Sum);
    }

    [Fact]
    public void NoNumericDataProducesEmptyStatisticsRatherThanZero()
    {
        var table = new DataTable();
        table.Columns.Add("value", typeof(object));
        table.Rows.Add("n/a");
        table.Rows.Add(true);
        var result = Assert.Single(Statistics.Analyze(table.Rows.Cast<DataRow>(), "value"));
        Assert.Equal(2, result.InvalidCount);
        Assert.Null(result.Sum);
        Assert.Null(result.Average);
        Assert.Empty(Statistics.Analyze(Array.Empty<DataRow>(), "value"));
    }

    [Theory]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("50%")]
    [InlineData("$100")]
    public void RejectsInvalidNumericText(string text) => Assert.False(Statistics.TryNumber(text, out _));

    [Fact]
    public void NumericTextUsesStableCulture()
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            Assert.True(Statistics.TryNumber("1,234.50", out var number));
            Assert.Equal(1234.5, number);
        }
        finally { CultureInfo.CurrentCulture = original; }
    }

    [Fact]
    public void CsvPreservesChineseQuotesNewlinesAndNeutralizesFormulas()
    {
        var path = Path.GetTempFileName();
        try
        {
            CsvExporter.Write(path, [new Summary("=1+1,\"測試\"\n下一行", 1, 1, 0, 0, -2, -2, -2, -2)],
                "範例.xlsx", "銷售", "金額", "無");
            var bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Take(3).SequenceEqual(new byte[] { 0xEF, 0xBB, 0xBF }));
            var text = File.ReadAllText(path, Encoding.UTF8);
            Assert.Contains("\"'=1+1,\"\"測試\"\"\n下一行\"", text);
            Assert.Contains("\"-2\"", text);
            Assert.Contains("範例.xlsx", text);
        }
        finally { File.Delete(path); }
    }
}
