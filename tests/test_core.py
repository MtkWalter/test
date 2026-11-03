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


def test_overflow_is_reported():
    with pytest.raises(OverflowError):
        analyze(pd.DataFrame({"v": [1e308, 1e308]}), "v")
