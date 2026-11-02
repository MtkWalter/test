using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using ExcelInsights.Core;
using Microsoft.Win32;

namespace ExcelInsights.App;

public partial class MainWindow : Window
{
    private const long TotalCellLimit = 2_000_000;
    private readonly ObservableCollection<ImportedSheet> _sheets = [];
    private readonly DispatcherTimer _refreshTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private IReadOnlyList<Summary> _summaries = [];
    private bool _updating = true;
    private bool _busy;
    private ImportedSheet? Selected => SheetsList.SelectedItem as ImportedSheet;

    public MainWindow()
    {
        InitializeComponent();
        SheetsList.ItemsSource = _sheets;
        _refreshTimer.Tick += (_, _) => { _refreshTimer.Stop(); Refresh(); };
        _updating = false;
    }

    private async void Import_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "匯入 Excel（可多選）", Multiselect = true,
            Filter = "Excel 檔案|*.xlsx;*.xls;*.xlsb"
        };
        if (dialog.ShowDialog(this) == true) await ImportFiles(dialog.FileNames);
    }

    private async Task ImportFiles(IEnumerable<string> paths)
    {
        if (_busy) return;
        var files = paths.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (files.Length > 20)
        {
            MessageBox.Show(this, "每次最多選取 20 個檔案。", "匯入限制");
            return;
        }
        _busy = true;
        ImportButton.IsEnabled = false;
        var errors = new List<string>();
        var imported = 0;
        try
        {
            foreach (var file in files)
            {
                if (_sheets.Any(s => string.Equals(s.FilePath, file, StringComparison.OrdinalIgnoreCase))) continue;
                StatusText.Text = $"正在匯入 {Path.GetFileName(file)}…";
                try
                {
                    var sheets = await Task.Run(() => WorkbookImporter.Read(file));
                    if (_sheets.Sum(s => s.CellCount) + sheets.Sum(s => s.CellCount) > TotalCellLimit)
                        throw new InvalidDataException("已載入資料超過 2,000,000 儲存格限制；請先移除部分工作表。");
                    foreach (var sheet in sheets) _sheets.Add(sheet);
                    if (sheets.Count == 0) errors.Add($"{Path.GetFileName(file)}：沒有可讀取的工作表。");
                    else imported++;
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    errors.Add($"{Path.GetFileName(file)}：{ex.Message}");
                }
            }
            if (SheetsList.SelectedIndex < 0 && _sheets.Count > 0) SheetsList.SelectedIndex = 0;
            Refresh();
            StatusText.Text += $" 本次新增 {imported} 個檔案。";
            if (errors.Count > 0) MessageBox.Show(this, string.Join("\n\n", errors), "部分檔案未匯入", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            _busy = false;
            ImportButton.IsEnabled = true;
        }
    }

    private void Sheet_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_updating) return;
        _updating = true;
        var columns = Selected?.Table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray() ?? [];
        NumericColumn.ItemsSource = columns;
        FilterColumn.ItemsSource = columns;
        GroupColumn.ItemsSource = new[] { "（不分組）" }.Concat(columns).ToArray();
        GroupColumn.SelectedIndex = 0;
        FilterColumn.SelectedIndex = columns.Length > 0 ? 0 : -1;
        FilterText.Text = "";
        var numeric = Selected is { } sheet ? columns.FirstOrDefault(c => sheet.Table.Rows.Cast<DataRow>()
            .Any(row => Statistics.TryNumber(row[c], out _))) : null;
        NumericColumn.SelectedItem = numeric ?? columns.FirstOrDefault();
        _updating = false;
        Refresh();
    }

    private void Options_Changed(object sender, SelectionChangedEventArgs e) => ScheduleRefresh();
    private void Filter_Changed(object sender, TextChangedEventArgs e) => ScheduleRefresh();
    private void ScheduleRefresh()
    {
        if (_updating) return;
        ExportButton.IsEnabled = false;
        _refreshTimer.Stop();
        _refreshTimer.Start();
    }

    private void Refresh()
    {
        _refreshTimer.Stop();
        _summaries = [];
        ExportButton.IsEnabled = false;
        SummaryGrid.ItemsSource = null;
        if (Selected is not { } sheet)
        {
            PreviewGrid.ItemsSource = null;
            SummaryHeading.Text = "統計結果";
            StatusText.Text = "尚未匯入資料。請選取 Excel 檔案開始。";
            return;
        }
        try
        {
            var rows = Statistics.Filter(sheet.Table, FilterColumn.SelectedItem as string, FilterText.Text);
            var preview = sheet.Table.Clone();
            foreach (var row in rows.Take(5000)) preview.ImportRow(row);
            PreviewGrid.ItemsSource = preview.DefaultView;
            if (NumericColumn.SelectedItem is string column)
            {
                _summaries = Statistics.Analyze(rows, column,
                    GroupColumn.SelectedIndex > 0 ? GroupColumn.SelectedItem as string : null);
                SummaryGrid.ItemsSource = _summaries;
                SummaryHeading.Text = $"統計結果 · {column}";
                ExportButton.IsEnabled = _summaries.Count > 0;
            }
            StatusText.Text = $"{sheet.DisplayName} ｜ 篩選後 {rows.Count:N0} 列，預覽前 {Math.Min(rows.Count, 5000):N0} 列；統計包含全部篩選結果。";
        }
        catch (Exception ex) when (ex is ArgumentException or OverflowException)
        {
            StatusText.Text = ex.Message;
        }
    }

    private void Preview_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        // Index binding also handles headers containing dots, brackets, or slashes.
        if (e.Column is DataGridTextColumn column && sender is DataGrid grid && grid.ItemsSource is DataView view)
        {
            var index = view.Table!.Columns.IndexOf(e.PropertyName);
            column.Binding = new Binding($"[{index}]") { Mode = BindingMode.OneWay };
            column.MaxWidth = 320;
        }
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (Selected is not { } sheet || NumericColumn.SelectedItem is not string column) return;
        Refresh();
        if (_summaries.Count == 0) return;
        var dialog = new SaveFileDialog { Filter = "CSV 檔案|*.csv", FileName = "Excel統計.csv", AddExtension = true };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            CsvExporter.Write(dialog.FileName, _summaries, Path.GetFileName(sheet.FilePath), sheet.Name, column,
                string.IsNullOrWhiteSpace(FilterText.Text) ? "無" : $"{FilterColumn.SelectedItem} 包含 {FilterText.Text}");
            StatusText.Text = $"已匯出 {_summaries.Count:N0} 組統計：{dialog.FileName}";
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            MessageBox.Show(this, ex.Message, "無法匯出", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ResetFilter_Click(object sender, RoutedEventArgs e) => FilterText.Clear();
    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || Selected is not { } sheet) return;
        _sheets.Remove(sheet);
        if (_sheets.Count > 0 && SheetsList.SelectedIndex < 0) SheetsList.SelectedIndex = 0;
    }
    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        _sheets.Clear();
        Refresh();
    }
    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = !_busy && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }
    private async void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files) await ImportFiles(files);
    }
}
