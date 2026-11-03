from copy import deepcopy
from dataclasses import asdict
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from excel_insights.core import Sheet, read_workbook
from excel_insights.reporting import (
    Rule, build_report, demo_sheet, export_report, formatted_number, load_config, save_config, suggest_config,
    load_remark_file, remark_context, save_remark_file,
)
from excel_insights.schema import schema_key


def test_detects_realistic_fourth_row_header_and_manual_override(tmp_path):
    path = tmp_path / "input.xlsx"
    book = Workbook(); sheet = book.active
    sheet.append(["", "1000", "FTA 5"])
    sheet.append(["單價", "V", "公式"])
    sheet.append([None])
    sheet.append(["Tech Node", "sales-NTD", "budget / act", "TSMC/non TSMC"])
    sheet.append(["3nm", 100, "act", "TSMC"])
    book.save(path)
    actual = read_workbook(path)[0]
    assert actual.header_row == 4
    assert len(actual.data) == 1
    assert actual.data.iloc[0]["Tech Node"] == "3nm"
    assert read_workbook(path, header_row=1)[0].header_row == 1


def test_display_rounding_matches_excel_at_half_values():
    assert formatted_number(12_500_000, divisor=1_000_000) == "13"
    assert formatted_number(-12_500_000, divisor=1_000_000) == "-13"
    assert formatted_number(1.145, multiplier=100) == "115"


def test_manual_remark_file_can_be_reopened_without_transaction_rows(tmp_path):
    source = demo_sheet(); config = suggest_config([source])
    remarks = {("node", "3nm"): "客戶要求追蹤\n下次更新", ("total", "總計（報表）"): "總覽"}
    path = tmp_path / "remarks.json"
    save_remark_file(path, remarks, config)
    context, restored = load_remark_file(path)
    assert context == remark_context(config)
    assert restored == remarks
    assert "sales-NTD" not in path.read_text()


def test_demo_aggregates_service_and_compares_without_sending_data():
    source = demo_sheet(); config = suggest_config([source])
    result = build_report([source], config)
    row = result.frame.iloc[0]
    assert row["TSMC act"] == 120_000_000
    assert row["TSMC ori"] == 100_000_000
    assert row["TSMC 差額"] == 20_000_000
    assert row["TSMC 比較 %"] == 1.2
    assert row["Total act"] == 191_000_000
    config.total_mode = "classified"
    assert build_report([source], config).frame.iloc[0]["Total act"] == 181_000_000


def test_auto_remark_follows_node_customer_part_hierarchy_and_all_triggered_nodes():
    source = demo_sheet(); config = suggest_config([source])
    result = build_report([source], config)
    assert result.counts["remark_nodes"] == 6
    assert len(result.remark_details) == 12
    assert result.frame["Remark"].str.contains("觸發：").all()
    text = result.frame.iloc[0]["Remark"]
    assert "TSMC 120.0%、Non-T 80.0%" in text
    assert "客戶 Demo Customer A：實績 141、預算 120、差額 +21" in text
    assert "品號 DEMO-PART：實績 141、預算 120、差額 +21" in text
    assert "客戶 Demo Customer B：實績 50、預算 60、差額 -10" in text
    assert result.remark_details.iloc[0]["差額"] == 21_000_000
    # A class can breach a threshold even while the node's Total stays in range.
    config.remark_basis = "total"
    assert build_report([source], config).counts.get("remark_nodes", 0) == 0


@pytest.mark.parametrize("actual,expected", [(115, True), (85, True), (114.99, False), (85.01, False)])
def test_remark_inclusive_ratio_thresholds(actual, expected):
    source = Sheet(Path("boundary.xlsx"), "data", pd.DataFrame({
        "Tech Node": ["3nm"] * 2, "Amount": [actual, 100], "scenario": ["act", "ori"],
        "Category": ["TSMC"] * 2, "CUSTOMER": ["Customer A"] * 2, "P/N": ["P1"] * 2,
    }, dtype=object))
    config = suggest_config([source])
    config.percentage = "change"; config.upper = 15; config.lower = -15
    result = build_report([source], config)
    assert bool(result.counts.get("remark_nodes")) is expected


def test_fi_alone_can_trigger_and_part_sums_combine_multiple_files():
    data = pd.DataFrame({
        "Tech Node": ["3nm"] * 8,
        "Amount": [100, 100, 100, 100, 30, 20, 10, 10],
        "scenario": ["act", "ori"] * 4,
        "Category": ["TSMC", "TSMC", "Non-T", "Non-T", "FI", "FI", "Service", "Service"],
        "CUSTOMER": ["Customer A"] * 6 + ["Customer B"] * 2,
        "P/N": ["P1"] * 4 + ["P2"] * 4,
    }, dtype=object)
    sources = [Sheet(Path("first.xlsx"), "data", data.iloc[::2]), Sheet(Path("second.xlsx"), "data", data.iloc[1::2])]
    config = suggest_config(sources); result = build_report(sources, config)
    assert result.frame.iloc[0]["Remark"].splitlines()[0] == "觸發：FI 150.0%"
    details = result.remark_details
    assert details.loc[(details["客戶"] == "Customer A") & (details["品號"] == "P1"), "實績"].item() == 200
    assert details.loc[(details["客戶"] == "Customer A") & (details["品號"] == "P2"), "差額"].item() == 10


def test_long_remarks_keep_every_customer_part_in_detail_sheet(tmp_path):
    rows = []
    for i in range(500):
        for version, amount in (("act", 2), ("ori", 1)):
            rows.append(["3nm", amount, version, "TSMC", "Customer A", f"P{i:04d}-" + "x" * 80])
    source = Sheet(Path("long.xlsx"), "data", pd.DataFrame(rows, columns=["Tech Node", "Amount", "scenario", "Category", "CUSTOMER", "P/N"], dtype=object))
    config = suggest_config([source]); config.display_unit = 1
    result = build_report([source], config)
    assert len(result.remark_details) == 500
    assert len(result.frame.iloc[0]["Remark"]) <= 32767
    assert "完整" not in result.frame.iloc[0]["Remark"] or "明細" in result.frame.iloc[0]["Remark"]
    assert "其餘內容詳見" in result.frame.iloc[0]["Remark"]
    path = tmp_path / "detail.xlsx"; export_report(path, result, config, [source])
    book = load_workbook(path)
    assert book["Remark 明細"].max_row == 501
    book.close()


def test_multi_file_mappings_and_overlapping_rules_do_not_double_total():
    first = demo_sheet()
    second = Sheet(Path("other.xlsx"), "other", first.data.rename(columns={"sales-NTD": "Amount"}).copy())
    config = suggest_config([first, second])
    config.rules["FI"] = deepcopy(config.rules["TSMC"])
    result = build_report([first, second], config)
    assert result.counts["overlap"] == 24
    assert result.frame.iloc[0]["Total act"] == 382_000_000
    assert result.frame.iloc[0]["FI act"] == 240_000_000
    assert any("重複" not in warning and "只加一次" in warning for warning in result.warnings)


def test_period_filters_and_fiscal_end_year():
    source = demo_sheet(); config = suggest_config([source])
    config.year_mode = "fiscal"; config.year = 2027; config.fiscal_start = 4; config.month = 9
    assert build_report([source], config).counts["used"] == 48
    config.year = 2026
    assert build_report([source], config).frame.empty
    config.year_mode = "calendar"; config.year = 2026
    assert len(build_report([source], config).frame) == 6
    config.year_mode = "fiscal"; config.fiscal_start = 1
    assert len(build_report([source], config).frame) == 6


def test_date_fallback_and_unknown_dates_are_reported():
    source = Sheet(Path("date.xlsx"), "date", pd.DataFrame({
        "Tech Node": ["3nm", "3nm", "3nm"], "Amount": [100, 80, 50],
        "scenario": ["act", "ori", "act"], "Date": ["2026-09-01", "2026-09-02", "invalid"],
    }, dtype=object))
    config = suggest_config([source]); config.month = 9
    result = build_report([source], config)
    assert result.counts["used"] == 2
    assert result.counts["invalid_period"] == 1
    assert result.frame.iloc[0]["Total act"] == 100


def test_bad_rows_and_zero_baseline_stay_visible_in_checks():
    source = Sheet(Path("bad.xlsx"), "bad", pd.DataFrame({
        "Tech Node": ["3nm", "3nm", "3nm", "3nm", None, "3nm"],
        "Amount": [100, 0, "not number", None, 10, 10],
        "scenario": ["act", "ori", "act", "act", "act", "forecast"],
        "Category": ["TSMC"] * 6,
    }, dtype=object))
    config = suggest_config([source]); result = build_report([source], config)
    row = result.frame.iloc[0]
    assert row["TSMC 差額"] == 100
    assert row["TSMC 比較 %"] is None
    assert "ori=0" in row["Remark"]
    assert result.counts["invalid_amount"] == 2
    assert result.counts["blank_node"] == 1
    assert result.counts["unknown_scenario"] == 1


def test_changed_filters_units_formula_and_total_ratio():
    source = demo_sheet(); config = suggest_config([source])
    config.filters = {"product": "Product"}
    config.multiplier = .001
    config.difference = "ori-act"; config.percentage = "change"
    config.upper = 15; config.lower = -15; config.total_row = True
    result = build_report([source], config)
    assert result.counts["filtered"] == 12
    row = result.frame.iloc[0]
    assert row["TSMC 差額"] == -20000
    assert row["TSMC 比較 %"] == -.2
    total = result.frame.iloc[-1]
    assert total["Tech Node"] == "總計（報表）"
    assert total["TSMC 比較 %"] == pytest.approx((total["TSMC ori"] - total["TSMC act"]) / total["TSMC ori"])


def test_presets_roundtrip_without_transaction_rows(tmp_path):
    source = demo_sheet(); config = suggest_config([source])
    config.year_mode = "fiscal"; config.filters = {"part": "DEMO-PART"}
    path = tmp_path / "settings.json"
    save_config(path, config)
    restored = load_config(path)
    assert asdict(restored) == asdict(config)
    assert "dataframe" not in path.read_text()
    assert "120000000" not in path.read_text()
    assert build_report([source], restored).frame.equals(build_report([source], config).frame)
    path.write_text('{"version":2,"settings":{}}')
    with pytest.raises(ValueError, match="版本"): load_config(path)


def test_missing_required_columns_and_invalid_threshold_block_output():
    source = demo_sheet(); config = suggest_config([source])
    config.mappings[schema_key(source.data.columns)]["amount"] = None
    with pytest.raises(ValueError, match="amount"): build_report([source], config)
    config = suggest_config([source]); config.upper = 50
    with pytest.raises(ValueError, match="門檻"): build_report([source], config)


def test_ambiguous_budget_and_ori_requires_local_selection_instead_of_guessing():
    source = demo_sheet()
    extra = source.data.iloc[[0]].copy()
    extra["budget / act"] = "budget"
    source.data = pd.concat([source.data, extra], ignore_index=True)
    config = suggest_config([source])
    assert config.actual == "act"
    assert config.original == ""
    with pytest.raises(ValueError, match="ori"): build_report([source], config)


def test_template_export_reopens_with_values_styles_rules_and_no_injected_formula(tmp_path):
    source = demo_sheet(); source.data.at[0, "Tech Node"] = "=1+1"
    source.data.at[0, "Remark"] = "=HYPERLINK(\"https://example.com\")"
    config = suggest_config([source]); result = build_report([source], config)
    path = tmp_path / "output.xlsx"
    export_report(path, result, config, [source])
    book = load_workbook(path, data_only=False); sheet = book.active
    assert sheet["D11"].value == "Tech Node"
    assert sheet["D12"].value == "=1+1"
    assert sheet["D12"].data_type == "s"
    assert sheet["R12"].data_type == "s"
    assert sheet["F12"].value == 100_000_000
    assert sheet["D11"].fill.fgColor.rgb == "FF002060"
    assert ",," in sheet["E12"].number_format
    assert "M10:N10" in [str(value) for value in sheet.merged_cells.ranges]
    assert sheet["N12"].value is None  # No ori for this distinct source group.
    assert "設定與檢查" in book.sheetnames
    assert sheet["K10"].value == "Total (service incl.)"
    rules = [rule.formula[0] for rules in sheet.conditional_formatting._cf_rules.values() for rule in rules]
    assert any("$N12>=1.15" in formula for formula in rules)
    assert not any("$L12>=1.15" in formula for formula in rules)
    book.close()
    config.total_mode = "classified"
    classified_path = tmp_path / "classified.xlsx"
    export_report(classified_path, build_report([source], config), config, [source])
    book = load_workbook(classified_path)
    assert book.active["K10"].value == "Total (classified)"
    book.close()
    with pytest.raises(ValueError, match="覆寫"): export_report(source.path, result, config, [source])


def test_report_grows_past_ten_rows_without_losing_formats(tmp_path):
    source = demo_sheet()
    source.data["Tech Node"] = [f"node-{i // 2}" for i in range(len(source.data))]
    config = suggest_config([source]); result = build_report([source], config)
    path = tmp_path / "long.xlsx"; export_report(path, result, config, [source])
    book = load_workbook(path)
    assert book.active["D35"].value == "node-23"
    assert book.active["E35"].number_format == book.active["E12"].number_format
    book.close()
