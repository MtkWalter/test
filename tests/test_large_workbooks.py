"""Regression cases for the user's former 100k-row / 1M-cell / 512-col errors."""
import hashlib

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

from excel_insights.core import analyze, read_workbook
from excel_insights.monthly_report import generate_report, read_report
from excel_insights.report_sample import create_sample


@pytest.fixture(scope="module")
def large_workbook(tmp_path_factory):
    path = tmp_path_factory.mktemp("large") / "large.xlsx"
    book = Workbook(write_only=True)
    sheet = book.create_sheet("來源")
    headers = ["sales-NTD", "budget / act", "TSMC/non TSMC", "Tech Node", "END USER", "P/N",
               "Year(calender)", "Month(calender)", "YM(calender)", "supplier", "客戶簡稱", "CUSTOMER"]
    sheet.append(headers)
    for scenario, value in (("act", 115), ("ori", 100)):
        for _ in range(50_001):
            sheet.append([value, f"FY27 {scenario}", "tsmc", "N20", "Large Customer", "PN",
                          2026, 9, "202609", "FT products", "NOT_THE_CUSTOMER", "NOT_THE_CUSTOMER"])
    pivot = book.create_sheet("monthly pivot(CS)")
    # A format at Excel's last column must not enlarge the data rectangle.
    far = WriteOnlyCell(pivot)
    far.font = Font(bold=True)
    pivot.append(["Tech Node", "金額"] + [None] * (16_384 - 3) + [far])
    pivot.append(["N20", 123])
    book.save(path)
    book.close()
    return path


def test_generic_reader_over_100k_rows_and_one_million_cells(large_workbook):
    before = hashlib.sha256(large_workbook.read_bytes()).digest()
    sheets = read_workbook(large_workbook)
    assert sheets[0].data.shape == (100_002, 12)
    assert sheets[0].cells == 1_200_024
    assert analyze(sheets[0].data, "sales-NTD").iloc[0]["總和"] == 10_750_215
    assert sheets[1].name == "monthly pivot(CS)"
    assert sheets[1].data.shape == (1, 2)
    assert sheets[1].data.iat[0, 1] == 123
    assert hashlib.sha256(large_workbook.read_bytes()).digest() == before


def test_report_streams_the_same_large_file_without_loading_pivot(large_workbook, tmp_path):
    output = tmp_path / "large-report.xlsx"
    result = generate_report(large_workbook, output, 2026, 9)
    assert result["act_rows"] == result["ori_rows"] == 50_001
    book = load_workbook(output, data_only=True)
    row = next(row for row in book.active if row[0].value == "N20")
    assert row[1].value == pytest.approx(5.750115)
    assert row[2].value == pytest.approx(5.0001)
    assert row[8].value == 1.15
    assert "Large Customer +0.8M" in row[11].value
    assert len(list(book["差異明細"].iter_rows(min_row=2))) == 1
    book.close()


def test_actual_data_after_512_columns_is_retained(tmp_path):
    path = tmp_path / "wide.xlsx"
    book = Workbook()
    sheet = book.active
    sheet["A1"] = "first"
    sheet.cell(1, 600, "last")
    sheet["A2"] = 1
    sheet.cell(2, 600, 42)
    sheet["XFD2"].font = Font(bold=True)
    book.save(path)
    book.close()
    result = read_workbook(path)[0].data
    assert result.shape == (1, 600)
    assert result["last"].iloc[0] == 42


def test_report_finds_required_header_after_column_512(tmp_path):
    source = create_sample(tmp_path / "moved.xlsx")
    book = load_workbook(source)
    sheet = book.active
    for row in range(1, sheet.max_row + 1):
        sheet.cell(row, 703, sheet[f"BI{row}"].value)
        sheet[f"BI{row}"] = None
    book.save(source)
    book.close()
    data = read_report(source, 2026, 9)
    assert ("N20", "TSMC", "END_A") in data.customers
    assert data.counts["act"] == 20


def test_format_only_rows_do_not_count_as_data_rows(tmp_path, monkeypatch):
    from excel_insights import core
    source = tmp_path / "formatted.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.append(["value"])
    sheet.append([123])
    sheet["XFD100010"].font = Font(bold=True)
    book.save(source)
    book.close()
    monkeypatch.setattr(core, "MAX_ROWS", 2)
    data = read_workbook(source)[0].data
    assert data.shape == (1, 1)
    assert data.iat[0, 0] == 123
