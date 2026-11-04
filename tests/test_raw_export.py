"""Preserve the complete source package, even when strict reporting rejects it."""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from excel_insights.monthly_report import ReportValidationError, generate_report, read_report
from excel_insights.raw_export import generate_report_with_raw, inspect_raw_workbook
from excel_insights.report_preview import preview_workbook
from excel_insights.report_sample import create_sample


def assert_raw_preserved(source, output):
    with ZipFile(source) as before, ZipFile(output) as after:
        changed = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "[Content_Types].xml"}
        assert set(before.namelist()) <= set(after.namelist())
        for name in before.namelist():
            if name not in changed:
                assert before.read(name) == after.read(name), name
        embedded = next(n for n in after.namelist() if n.startswith("raw-input-") and "/original." in n)
        assert after.read(embedded) == source.read_bytes()
        assert after.testzip() is None


def test_raw_export_preserves_raw_formulas_styles_hidden_sheets_and_report(tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    raw = book.active
    raw["A100"] = "=SUM(1,2)"
    raw["A100"].font = Font(color="FF123456", bold=True)
    raw["B100"] = "#VALUE!"
    raw["C100"] = "=this is literal"
    raw["C100"].data_type = "s"
    raw.merge_cells("A102:E102")
    raw.column_dimensions["AB"].hidden = True
    raw["AAA105"] = "beyond 512 columns"
    book.create_sheet("隱藏原始資料").sheet_state = "hidden"
    book.save(source)
    book.close()
    original = source.read_bytes()
    output = tmp_path / "output.xlsx"
    result = generate_report_with_raw(source, output, 2026, 9)
    assert result["report_generated"] is True and result["act_rows"] == 20
    assert result["raw_sheet_count"] == 3
    assert_raw_preserved(source, output)
    book = load_workbook(output)
    assert book["銷售明細"]["A100"].value == "=SUM(1,2)"
    assert book["銷售明細"]["A100"].font.color.rgb == "FF123456"
    assert book["銷售明細"]["C100"].data_type == "s"
    assert book["銷售明細"]["AAA105"].value == "beyond 512 columns"
    assert book["隱藏原始資料"].sheet_state == "hidden"
    row = next(row for row in book[result["preview_sheet"]] if row[0].value == "N20")
    assert row[1].value == 120 and row[8].value == 1.2 and row[10].value == .8
    assert row[8].fill.fgColor.rgb == "00FFF0AF"
    assert row[10].fill.fgColor.rgb == "00C9EFF3"
    assert row[8].number_format == "0%"
    book.close()
    preview = preview_workbook(output, result["preview_sheet"])
    assert any(r["cells"][0] == "N20" and r["cells"][8] == "120%" for r in preview["rows"])
    assert source.read_bytes() == original


def test_all_errors_and_multiple_fields_per_row_are_retained(tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    for number in range(3, 43):
        row = book.active
        row[f"BD{number}"] = "FY27 act"
        row[f"BP{number}"] = "202609"
        row[f"BO{number}"] = 2026
        row[f"BE{number}"] = 9
        row[f"AS{number}"] = "not money: " + "x" * 500
        row[f"BF{number}"] = None
        row[f"BM{number}"] = None
        row[f"BB{number}"] = "ordinary"
    book.save(source)
    book.close()
    with pytest.raises(ReportValidationError) as caught:
        read_report(source, 2026, 9)
    issues = caught.value.issues
    assert len(issues) == 120
    assert issues[0]["row"] == 3 and issues[-1]["row"] == 42
    assert len(issues[0]["value"]) > 500
    assert {i["column"] for i in issues[:3]} == {"AS", "BF", "BM, BB"}
    strict = tmp_path / "strict.xlsx"
    with pytest.raises(ReportValidationError):
        generate_report(source, strict, 2026, 9)
    assert not strict.exists()
    output = tmp_path / "raw.xlsx"
    result = generate_report_with_raw(source, output, 2026, 9)
    assert result["report_generated"] is False and result["act_rows"] is None
    assert result["issues"] == issues
    assert_raw_preserved(source, output)
    book = load_workbook(output)
    assert "monthly pivot" not in book.sheetnames
    rows = list(book["完整錯誤清單"].values)
    assert len(rows) == 121 and rows[-1][2] == 42
    assert rows[1][5] == issues[0]["value"]
    assert "未產生統計" in book["匯出說明"]["B2"].value
    book.close()


def test_missing_headers_and_invalid_settings_still_export_all_sheets(tmp_path):
    source = tmp_path / "input.xlsx"
    book = Workbook()
    book.active.title = "No headers"
    book.active.append(["unrecognized", "original value"])
    book.create_sheet("Other").append([1, 2, "=1+1"])
    book.save(source)
    book.close()
    catalog = inspect_raw_workbook(source)
    assert catalog["report_sheets"] == [] and "必要欄名" in catalog["format_error"]
    for index, kwargs in enumerate(({}, {"rules": {"tsmc_values": []}}, {"settings_error": "自訂設定錯誤"})):
        output = tmp_path / f"output{index}.xlsx"
        result = generate_report_with_raw(source, output, 2026, 9, **kwargs)
        assert not result["report_generated"] and result["issues"]
        assert_raw_preserved(source, output)


def test_colliding_sheet_names_and_styles_do_not_replace_source(tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    for name in ("monthly pivot", "差異明細", "報表資訊", "匯出說明", "完整錯誤清單"):
        sheet = book.create_sheet(name)
        sheet["A1"] = "Original " + name
        sheet["A1"].fill = PatternFill("solid", fgColor="FF010203")
        sheet["A1"].number_format = '"Original "0.0000'
    book.save(source)
    book.close()
    output = tmp_path / "output.xlsx"
    result = generate_report_with_raw(source, output, 2026, 9, sheet="銷售明細")
    assert result["preview_sheet"] != "monthly pivot"
    assert result["info_sheet"] != "匯出說明"
    assert_raw_preserved(source, output)
    book = load_workbook(output)
    assert book["monthly pivot"]["A1"].value == "Original monthly pivot"
    assert book["monthly pivot"]["A1"].fill.fgColor.rgb == "FF010203"
    assert any(row[0].value == "N20" for row in book[result["preview_sheet"]])
    assert "原始檔名" == book[result["info_sheet"]]["A3"].value
    book.close()


def test_macro_payload_and_correct_extension_are_preserved(tmp_path):
    ordinary = create_sample(tmp_path / "ordinary.xlsx")
    source = tmp_path / "macros.xlsm"
    with ZipFile(ordinary) as original, ZipFile(source, "w", ZIP_DEFLATED) as target:
        for entry in original.infolist():
            data = original.read(entry)
            if entry.filename == "[Content_Types].xml":
                data = data.replace(b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", b"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            target.writestr(entry, data)
        target.writestr("xl/vbaProject.bin", b"macro-test-payload-preserve-without-execution")
    result = generate_report_with_raw(source, tmp_path / "out.xlsx", 2026, 9)
    assert result["output_extension"] == ".xlsm"
    assert result["output"].endswith("out.xlsm")
    assert_raw_preserved(source, result["output"])
    with ZipFile(result["output"]) as output:
        assert output.read("xl/vbaProject.bin") == b"macro-test-payload-preserve-without-execution"


def test_raw_export_refuses_source_overwrite_and_invalid_package(tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    original = source.read_bytes()
    with pytest.raises(ValueError, match="不可覆寫輸入"):
        generate_report_with_raw(source, source, 2026, 9, overwrite=True)
    assert source.read_bytes() == original
    existing = tmp_path / "exists.xlsx"
    existing.write_bytes(b"keep existing")
    with pytest.raises(FileExistsError):
        generate_report_with_raw(source, existing, 2026, 9)
    assert existing.read_bytes() == b"keep existing"
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a ZIP workbook")
    with pytest.raises(ValueError, match="損毀、加密"):
        inspect_raw_workbook(broken)


def test_long_error_values_use_continuation_rows_without_formula_injection(tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    original_value = "=" + "X" * 32766
    book.active["AS3"] = original_value
    book.active["AS3"].data_type = "s"
    book.save(source)
    book.close()
    output = tmp_path / "out.xlsx"
    result = generate_report_with_raw(source, output, 2026, 9)
    assert not result["report_generated"] and len(result["issues"]) == 1
    book = load_workbook(output)
    errors = list(book["完整錯誤清單"].iter_rows(min_row=2))
    assert len(errors) == 2
    assert "".join(row[5].value or "" for row in errors) == "AS=" + original_value
    assert all(row[5].data_type == "s" for row in errors)
    book.close()
    assert_raw_preserved(source, output)
