using System.IO.Compression;
using ExcelInsights.Core;
using Xunit;

namespace ExcelInsights.Tests;

public class ImporterTests
{
    [Fact]
    public void DemoWorkbooksMatchDocumentedTotalsAndFilter()
    {
        var january = WorkbookImporter.Read(Path.Combine(AppContext.BaseDirectory, "samples", "銷售範例-一月.xlsx"));
        var february = WorkbookImporter.Read(Path.Combine(AppContext.BaseDirectory, "samples", "銷售範例-二月.xlsx"));
        Assert.Equal(2, january.Count);
        Assert.Single(february);
        var all = Statistics.Analyze(january[0].Table.Rows.Cast<System.Data.DataRow>(), "金額");
        var total = Assert.Single(all);
        Assert.Equal((6, 4, 1, 1), (total.Rows, total.NumericCount, total.BlankCount, total.InvalidCount));
        Assert.Equal(11100, total.Sum);
        Assert.Equal(2775, total.Average);
        var filtered = Statistics.Filter(january[0].Table, "地區", "台北");
        var group = Assert.Single(Statistics.Analyze(filtered, "金額", "地區"));
        Assert.Equal(3, group.Rows);
        Assert.Equal(8400, group.Sum);
        Assert.Equal(4200, group.Average);
        Assert.Equal(17600, Assert.Single(Statistics.Analyze(february[0].Table.Rows.Cast<System.Data.DataRow>(), "金額")).Sum);
    }

    [Fact]
    public void ReadsActualXlsxMultipleSheetsAndUniqueHeadersThenAnalyzes()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                Add(zip, "[Content_Types].xml", """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                      <Default Extension="xml" ContentType="application/xml"/>
                      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                      <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                    </Types>
                    """);
                Add(zip, "_rels/.rels", """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                    </Relationships>
                    """);
                Add(zip, "xl/workbook.xml", """
                    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <sheets><sheet name="銷售" sheetId="1" r:id="rId1"/><sheet name="第二表" sheetId="2" r:id="rId2"/></sheets>
                    </workbook>
                    """);
                Add(zip, "xl/_rels/workbook.xml.rels", """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
                    </Relationships>
                    """);
                Add(zip, "xl/worksheets/sheet1.xml", """
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C4"/><sheetData>
                      <row r="1"><c r="A1" t="inlineStr"><is><t>金額</t></is></c><c r="B1" t="inlineStr"><is><t>金額</t></is></c><c r="C1"/></row>
                      <row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>20</v></c><c r="C2"><v>1</v></c></row>
                      <row r="3"/>
                      <row r="4"><c r="A4"><v>30</v></c></row>
                    </sheetData></worksheet>
                    """);
                Add(zip, "xl/worksheets/sheet2.xml", """
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                      <row r="1"><c r="A1" t="inlineStr"><is><t>數量</t></is></c></row>
                      <row r="2"><c r="A2"><v>7</v></c></row>
                    </sheetData></worksheet>
                    """);
            }
            var sheets = WorkbookImporter.Read(path);
            Assert.Equal(2, sheets.Count);
            Assert.Equal("銷售", sheets[0].Name);
            Assert.Equal(2, sheets[0].Table.Rows.Count);
            Assert.Equal("金額 (2)", sheets[0].Table.Columns[1].ColumnName);
            Assert.Equal("欄位 3", sheets[0].Table.Columns[2].ColumnName);
            Assert.Equal(40, Assert.Single(Statistics.Analyze(sheets[0].Table.Rows.Cast<System.Data.DataRow>(), "金額")).Sum);
            Assert.Equal(7d, sheets[1].Table.Rows[0][0]);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DamagedWorkbookFailsWithoutChangingSource()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            File.WriteAllText(path, "not a workbook");
            Assert.ThrowsAny<Exception>(() => WorkbookImporter.Read(path));
            Assert.Equal("not a workbook", File.ReadAllText(path));
        }
        finally { File.Delete(path); }
    }

    private static void Add(ZipArchive zip, string name, string xml)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(xml);
    }
}
