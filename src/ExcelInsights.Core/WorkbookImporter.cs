using System.Data;
using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace ExcelInsights.Core;

public sealed record ImportedSheet(string FilePath, string Name, DataTable Table)
{
    public string DisplayName => $"{Path.GetFileName(FilePath)} / {Name}（{Table.Rows.Count:N0} 列）";
    public long CellCount => (long)Table.Rows.Count * Table.Columns.Count;
}

public static class WorkbookImporter
{
    public const int MaxRows = 100_000;
    public const int MaxColumns = 512;
    public const long MaxCells = 1_000_000;

    public static IReadOnlyList<ImportedSheet> Read(string path)
    {
        if (!new[] { ".xls", ".xlsx", ".xlsb" }.Contains(Path.GetExtension(path).ToLowerInvariant()))
            throw new InvalidDataException("只支援 .xlsx、.xls、.xlsb 檔案。");
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var sheets = new List<ImportedSheet>();
        long cells = 0;
        do
        {
            if (reader.FieldCount == 0 || !reader.Read()) continue;
            if (reader.FieldCount > MaxColumns) throw new InvalidDataException($"工作表超過 {MaxColumns} 欄限制。");
            var table = new DataTable { Locale = CultureInfo.InvariantCulture, CaseSensitive = true };
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var original = Convert.ToString(reader.GetValue(i), CultureInfo.InvariantCulture)?.Trim();
                if (string.IsNullOrWhiteSpace(original)) original = $"欄位 {i + 1}";
                var name = original;
                for (var suffix = 2; !names.Add(name); suffix++) name = $"{original} ({suffix})";
                table.Columns.Add(name, typeof(object));
            }
            while (reader.Read())
            {
                cells += reader.FieldCount;
                if (cells > MaxCells || table.Rows.Count >= MaxRows)
                    throw new InvalidDataException($"檔案超過上限：每表 {MaxRows:N0} 列、每檔 {MaxCells:N0} 儲存格。");
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
                if (values.All(Statistics.IsBlank)) continue;
                table.Rows.Add(values);
            }
            sheets.Add(new ImportedSheet(path, reader.Name, table));
        } while (reader.NextResult());
        return sheets;
    }
}
