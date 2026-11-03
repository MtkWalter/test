from copy import deepcopy
from dataclasses import asdict
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from excel_insights.core import Sheet, read_workbook
from excel_insights.reporting import (
    Rule, build_report, demo_sheet, export_report, load_config, save_config, suggest_config,
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
    rules = [rule.formula[0] for rules in sheet.conditional_formatting._cf_rules.values() for rule in rules]
    assert any("$N12>=1.15" in formula for formula in rules)
    assert not any("$L12>=1.15" in formula for formula in rules)
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
