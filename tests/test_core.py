import csv
import hashlib
from datetime import datetime

import numpy as np
import pandas as pd
import pytest
from openpyxl import Workbook

from excel_insights import core
from excel_insights.core import analyze, export_csv, filter_rows, number, read_workbook, unique_headers


def write_xlsx(path, rows):
    book = Workbook()
    for row in rows:
        book.active.append(row)
    book.save(path)
    book.close()
    return path


def test_original_samples_keep_same_totals(samples):
    january = read_workbook(samples / "銷售範例-一月.xlsx")
    february = read_workbook(samples / "銷售範例-二月.xlsx")
    assert len(january) == 2
    assert len(february) == 1
    result = analyze(january[0].data, "金額").iloc[0]
    assert result.tolist() == ["全部", 6, 4, 1, 1, 11100.0, 2775.0, 1200.0, 6000.0]
    filtered = filter_rows(january[0].data, "地區", "台北")
    assert analyze(filtered, "金額", "地區").iloc[0].tolist() == ["台北", 3, 2, 1, 0, 8400, 4200, 2400, 6000]
    assert analyze(february[0].data, "金額").iloc[0]["總和"] == 17600


def test_preserves_types_headers_and_literal_na(tmp_path):
    path = write_xlsx(tmp_path / "types.xlsx", [
        ["a", "a", None, "文字"],
        [10, True, datetime(2026, 1, 2), "NA"],
        [None, None, None, None],
        ["1,200.50", False, None, "None"],
    ])
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    sheet = read_workbook(path)[0]
    assert list(sheet.data.columns) == ["a", "a (2)", "欄位 3", "文字"]
    assert len(sheet.data) == 2
    assert sheet.data.iat[0, 1] is True
    assert isinstance(sheet.data.iat[0, 2], datetime)
    assert sheet.data.iat[0, 3] == "NA"
    assert analyze(sheet.data, "a").iloc[0]["總和"] == 1210.5
    assert analyze(sheet.data, "a (2)").iloc[0]["非數值"] == 2
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before


def test_legacy_xls_dates_and_booleans(tmp_path):
    import xlwt
    path = tmp_path / "legacy.xls"
    book = xlwt.Workbook()
    sheet = book.add_sheet("舊格式")
    sheet.write(0, 0, "金額")
    sheet.write(0, 1, "日期")
    sheet.write(0, 2, "是否")
    sheet.write(1, 0, -12.5)
    sheet.write(1, 1, datetime(2026, 1, 2), xlwt.easyxf(num_format_str="YYYY-MM-DD"))
    sheet.write(1, 2, True)
    book.save(str(path))
    data = read_workbook(path)[0].data
    assert data.iat[0, 0] == -12.5
    assert data.iat[0, 1] == datetime(2026, 1, 2)
    assert data.iat[0, 2] is True
    assert analyze(data, "金額").iloc[0]["總和"] == -12.5


@pytest.mark.parametrize("value,expected", [
    ("1,234.50", 1234.5), ("-5", -5), (".5", .5), ("2e3", 2000), (0, 0),
    (np.int64(7), 7), (True, None), (np.bool_(True), None), ("NaN", None),
    ("Infinity", None), (float("inf"), None), ("1,2", None), ("50%", None),
    ("$100", None), (" ", None), (datetime(2026, 1, 1), None),
])
def test_numeric_conversion(value, expected):
    assert number(value) == expected


def test_invalid_and_empty_values_are_not_zero():
    data = pd.DataFrame({"value": [None, "", "n/a", True]}, dtype=object)
    summary = analyze(data, "value").iloc[0]
    assert summary["空白"] == 2
    assert summary["非數值"] == 2
    assert summary["總和"] is None
    assert summary["平均"] is None
    assert analyze(data.iloc[:0], "value").empty


def test_blank_groups_do_not_merge_with_literal_blank_label():
    data = pd.DataFrame({"group": [None, "（空白）", "Taipei", "Taipei"], "v": [1, 2, 3, 4]})
    summary = analyze(data, "v", "group")
    assert summary["總和"].tolist() == [1, 2, 7]
    assert len(set(summary["群組"])) == 3
    assert len(filter_rows(data, "group", "TAI")) == 2
    assert filter_rows(data, "group", "not found").empty


def test_duplicate_headers_always_unique_even_if_suffix_already_exists():
    names = unique_headers(["A", "A", "a (2)", None, "欄位 4"])
    assert len({name.casefold() for name in names}) == 5


def test_csv_roundtrip_and_formula_protection(tmp_path):
    path = tmp_path / "result.csv"
    summary = analyze(pd.DataFrame({"g": ['=1+1,"測試"\n下一行'], "v": [-2]}), "v", "g")
    export_csv(path, summary, file="範例.xlsx", sheet="銷售", column="v", group="g", query="無")
    assert path.read_bytes().startswith(b"\xef\xbb\xbf")
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.reader(stream))
    assert rows[1][0] == "範例.xlsx"
    assert rows[1][5] == "'=1+1,\"測試\"\n下一行"
    assert rows[1][-4] == "-2.0"
    assert len(rows[0]) == len(rows[1])


def test_corrupt_or_unsupported_files_fail_without_modification(tmp_path):
    path = tmp_path / "bad.xlsx"
    path.write_text("not an Excel file")
    with pytest.raises(Exception):
        read_workbook(path)
    assert path.read_text() == "not an Excel file"
    with pytest.raises(ValueError, match="只支援"):
        read_workbook(tmp_path / "file.csv")


def test_bounds_reject_file_instead_of_truncating(tmp_path, monkeypatch):
    path = write_xlsx(tmp_path / "large.xlsx", [["v"], [1], [2], [3]])
    monkeypatch.setattr(core, "MAX_ROWS", 2)
    with pytest.raises(ValueError, match="列限制"):
        read_workbook(path)


def test_late_wide_row_cannot_expand_past_cell_limit(tmp_path, monkeypatch):
    path = write_xlsx(tmp_path / "wide.xlsx", [["v"], [1], [2], list(range(10))])
    monkeypatch.setattr(core, "MAX_CELLS", 20)
    with pytest.raises(ValueError, match="儲存格"):
        read_workbook(path)


def test_styled_last_excel_cell_does_not_inflate_actual_range(tmp_path):
    from openpyxl.styles import PatternFill
    path = tmp_path / "monthly.xlsx"
    book = Workbook(); sheet = book.active; sheet.title = "monthly pivot(CS)"
    sheet.append(["Tech Node", "Amount"])
    sheet.append(["3nm", 120])
    sheet["XFD1048576"].fill = PatternFill("solid", fgColor="FFFFFF00")
    book.save(path); book.close()
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    result = read_workbook(path)[0]
    assert result.name == "monthly pivot(CS)"
    assert result.data.shape == (1, 2)
    assert result.data.iloc[0].tolist() == ["3nm", 120]
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before


def test_format_padding_preserves_fourth_row_header_and_skips_empty_sheet(tmp_path):
    from openpyxl.styles import PatternFill
    path = tmp_path / "mixed.xlsx"
    book = Workbook(); sheet = book.active; sheet.title = "input"
    sheet.cell(2, 2, "FTA 5"); sheet.cell(3, 3, "V")
    for col, value in enumerate(["Tech Node", "Amount", "scenario"], 1): sheet.cell(4, col, value)
    for col, value in enumerate(["3nm", 120, "act"], 1): sheet.cell(5, col, value)
    sheet.cell(6, 600).number_format = "#,##0"
    blank = book.create_sheet("format only")
    blank["XFD1048576"].fill = PatternFill("solid", fgColor="FFFF0000")
    book.save(path); book.close()
    results = read_workbook(path)
    assert len(results) == 1
    assert results[0].header_row == 4
    assert results[0].data.shape == (1, 3)
    assert read_workbook(path, header_row=4)[0].data.equals(results[0].data)


def test_real_columns_past_512_are_imported_without_dropping_last_value(tmp_path):
    path = tmp_path / "wide-pivot.xlsx"
    book = Workbook(); sheet = book.active
    sheet.cell(1, 1, "Tech Node"); sheet.cell(1, 600, "Amount")
    sheet.cell(2, 1, "3nm"); sheet.cell(2, 600, 125)
    book.save(path); book.close()
    result = read_workbook(path)[0]
    assert result.data.shape == (1, 600)
    assert result.data.iloc[0]["Amount"] == 125
    assert analyze(result.data, "Amount").iloc[0]["總和"] == 125


def test_sparse_data_rows_are_counted_not_blank_physical_gaps(tmp_path, monkeypatch):
    path = tmp_path / "sparse.xlsx"
    book = Workbook(); sheet = book.active
    sheet.append(["Tech Node", "Amount"])
    sheet.append(["3nm", 120])
    sheet.cell(1_048_576, 1, "5nm"); sheet.cell(1_048_576, 2, 80)
    book.save(path); book.close()
    monkeypatch.setattr(core, "MAX_ROWS", 2)
    monkeypatch.setattr(core, "MAX_CELLS", 10)
    result = read_workbook(path)[0]
    assert len(result.data) == 2
    assert result.data["Amount"].tolist() == [120, 80]


def test_actual_column_limit_still_rejects_data_instead_of_truncating(tmp_path, monkeypatch):
    path = write_xlsx(tmp_path / "too-wide.xlsx", [["Tech Node", "Amount", "extra"], ["3nm", 120, 1]])
    monkeypatch.setattr(core, "MAX_COLUMNS", 2)
    with pytest.raises(ValueError, match="有效資料"):
        read_workbook(path)


def test_sparse_reader_preserves_cached_formulas_and_1904_dates(tmp_path):
    from zipfile import ZipFile, ZIP_DEFLATED
    import xml.etree.ElementTree as ET
    from openpyxl.utils.datetime import CALENDAR_MAC_1904
    source = tmp_path / "cached-source.xlsx"
    book = Workbook(); book.epoch = CALENDAR_MAC_1904; sheet = book.active
    sheet.append(["Tech Node", "Amount", "日期", "是否"])
    sheet.append(["3nm", "=2+3", datetime(2026, 9, 1), False])
    sheet.cell(1048576, 16384).number_format = "0.00"
    book.save(source); book.close()
    path = tmp_path / "cached.xlsx"
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(source) as original, ZipFile(path, "w", ZIP_DEFLATED) as patched:
        for entry in original.infolist():
            data = original.read(entry.filename)
            if entry.filename == "xl/worksheets/sheet1.xml":
                xml = ET.fromstring(data)
                xml.find("s:dimension", ns).set("ref", "A1:A1")
                xml.find(".//s:c[@r='B2']/s:v", ns).text = "5"
                data = ET.tostring(xml, encoding="utf-8")
            patched.writestr(entry, data)
    result = read_workbook(path)[0]
    assert result.data.shape == (1, 4)
    assert result.data.iloc[0].tolist() == ["3nm", 5, datetime(2026, 9, 1), False]


def test_overflow_is_reported():
    with pytest.raises(OverflowError):
        analyze(pd.DataFrame({"v": [1e308, 1e308]}), "v")


def test_xlsb_checks_dimensions_before_engine_allocates_row_arrays(tmp_path, monkeypatch):
    from contextlib import contextmanager
    from types import SimpleNamespace
    import pyxlsb

    class OversizedSheet:
        dimension = SimpleNamespace(r=0, c=0, h=2, w=16385)

        def rows(self, **kwargs):
            pytest.fail("Oversized binary worksheet should be rejected before allocating rows")

    @contextmanager
    def sheet_context(*args):
        yield OversizedSheet()

    @contextmanager
    def workbook_context(*args):
        yield SimpleNamespace(sheets=["large"], get_sheet=sheet_context)

    monkeypatch.setattr(pyxlsb, "open_workbook", workbook_context)
    with pytest.raises(ValueError, match="限制"):
        read_workbook(tmp_path / "large.xlsb")
