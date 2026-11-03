import hashlib
import json
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from excel_insights.monthly_report import (
    ReportRules, _chunks, build_workbook, contribution_order, generate_report,
    inspect_workbook, main, money, read_report,
)
from excel_insights.report_sample import create_sample


@pytest.fixture
def report_input(tmp_path):
    return create_sample(tmp_path / "input.xlsx")


def get_row(sheet, label):
    return next(row for row in sheet.iter_rows(min_row=4) if row[0].value == label)


def test_monthly_report_roundtrip_filters_and_both_remarks(report_input, tmp_path):
    original = hashlib.sha256(report_input.read_bytes()).hexdigest()
    output = tmp_path / "報表.xlsx"
    result = generate_report(report_input, output, 2026, 9)
    assert result["act_rows"] == result["ori_rows"] == 20
    assert result["fiscal_year"] == "FY27"
    assert hashlib.sha256(report_input.read_bytes()).hexdigest() == original
    book = load_workbook(output, data_only=True)
    assert book.sheetnames == ["monthly pivot", "差異明細", "報表資訊"]
    row = get_row(book.active, "N20")
    assert [c.value for c in row[:11]] == ["N20", 120, 100, 80, 100, 200, 200, 20, 1.2, -20, 0.8]
    remark = row[11].value
    assert "TSMC\n" in remark and "Non-T\n" in remark
    assert remark.index("END_A +40.0M") < remark.index("END_C +10.0M") < remark.index("END_B -30.0M")
    assert remark.index("NON_T_DOWN -50.0M") < remark.index("NON_T_UP +30.0M")
    assert remark.index("PL_A1 +35.0M") < remark.index("PL_A2 +5.0M")
    assert remark.index("PL_B_DOWN -50.0M") < remark.index("PL_B_UP +20.0M")
    assert "EXCLUDED" not in remark
    assert "CUSTOMER並非" not in remark and "簡稱並非" not in remark
    assert row[7].fill.fgColor.rgb == row[8].fill.fgColor.rgb == "00FFF0AF"
    assert row[9].fill.fgColor.rgb == row[10].fill.fgColor.rgb == "00C9EFF3"
    assert row[8].number_format == row[10].number_format == "0%"
    assert all(row[i].number_format == '#,##0;[Red]-#,##0;0' for i in [1, 2, 3, 4, 5, 6, 7, 9])
    assert len(list(book["差異明細"].iter_rows(min_row=2))) == 20
    book.close()


def test_boundaries_merges_and_use_raw_ratios(report_input, tmp_path):
    generate_report(report_input, tmp_path / "out.xlsx", 2026, 9)
    book = load_workbook(tmp_path / "out.xlsx")
    for node, value, color in [("N2/A16", 1.15, "00FFF0AF"), ("N3/N3E", .85, "00C9EFF3")]:
        row = get_row(book.active, node)
        assert row[8].value == value
        assert row[11].value
        assert row[8].fill.fgColor.rgb == color
    assert "Non-T\n" not in get_row(book.active, "N2/A16")[11].value
    for node in ["N5", "N40", "N28"]:
        assert get_row(book.active, node)[11].value is None
    # Visually rounded budget 0 must not suppress a valid 150% ratio.
    row = get_row(book.active, "OT")
    assert row[4].value == .4
    assert row[10].value == 1.5
    assert "Non-T\n" in row[11].value and "TSMC\n" not in row[11].value
    book.close()


def test_zero_missing_negative_budget_and_service_adj_use_same_totals(report_input, tmp_path):
    generate_report(report_input, tmp_path / "out.xlsx", 2026, 9)
    book = load_workbook(tmp_path / "out.xlsx")
    for node, ratio_col in [("N7", 8), ("N14", 10), ("Reclaim", 8)]:
        row = get_row(book.active, node)
        assert row[ratio_col].value is None
        assert row[11].value is None
        assert row[ratio_col].fill.fgColor.rgb not in {"00FFF0AF", "00C9EFF3"}
    assert get_row(book.active, "Reclaim")[2].value is None
    for node in ["Service", "ADJ"]:
        row = get_row(book.active, node)
        assert row[5].value == (row[1].value or 0) + (row[3].value or 0)
        assert row[6].value == (row[2].value or 0) + (row[4].value or 0)
    book.close()


def test_missing_actual_or_formula_cache_fails_without_partial_output(report_input, tmp_path):
    book = load_workbook(report_input)
    book.active["AS3"] = None
    book.save(report_input)
    book.close()
    output = tmp_path / "out.xlsx"
    with pytest.raises(ValueError, match="實績金額空白"):
        generate_report(report_input, output, 2026, 9)
    assert not output.exists()
    book = load_workbook(report_input)
    book.active["AS3"] = "=1+2"
    book.save(report_input)
    book.close()
    with pytest.raises(ValueError, match="公式沒有已儲存"):
        generate_report(report_input, output, 2026, 9)
    assert not output.exists()


def test_explicit_rules_reject_unknown_segment_and_preserve_unknown_nodes(report_input, tmp_path):
    with pytest.raises(ValueError, match="尚未納入分類"):
        read_report(report_input, 2026, 9, rules=ReportRules(non_t_values=["non-t"]))
    rules = ReportRules(non_t_values=["non-t", "FI"], node_aliases={"N20": "CUSTOM_NODE"})
    data = read_report(report_input, 2026, 9, rules=rules)
    assert "CUSTOM_NODE" in data.nodes
    assert data.totals[("CUSTOM_NODE", "TSMC")].actual == Decimal("120000000")
    rules_path = tmp_path / "rules.json"
    rules_path.write_text(json.dumps(rules.to_dict()), encoding="utf-8")
    assert ReportRules.load(rules_path).to_dict() == rules.to_dict()


def test_single_month_year_and_ambiguous_sheets(report_input, tmp_path):
    with pytest.raises(ValueError, match="找不到 2026/10"):
        read_report(report_input, 2026, 10)
    book = load_workbook(report_input)
    book.copy_worksheet(book.active).title = "第二張"
    book.save(report_input)
    book.close()
    assert len(inspect_workbook(report_input)) == 2
    with pytest.raises(ValueError, match="多個可用工作表"):
        read_report(report_input, 2026, 9)
    assert read_report(report_input, 2026, 9, sheet="第二張").counts["act"] == 20


def test_metadata_conflicts_and_unknown_bf_or_node_fail(report_input):
    for column, value, message, original in [("BP", "202610", "不一致", "202609"),
                                             ("BF", None, "客群分類空白", "tsmc"),
                                             ("BM", None, "Tech Node 空白", "N20")]:
        book = load_workbook(report_input)
        book.active[f"{column}3"] = value
        book.save(report_input)
        book.close()
        with pytest.raises(ValueError, match=message):
            read_report(report_input, 2026, 9)
        book = load_workbook(report_input)
        book.active[f"{column}3"] = original
        book.save(report_input)
        book.close()


def test_output_cannot_overwrite_input_or_existing_file(report_input, tmp_path):
    original = report_input.read_bytes()
    with pytest.raises(ValueError, match="不可覆寫輸入"):
        generate_report(report_input, report_input, 2026, 9, overwrite=True)
    assert report_input.read_bytes() == original
    output = tmp_path / "out.xlsx"
    output.write_bytes(b"user data")
    with pytest.raises(FileExistsError):
        generate_report(report_input, output, 2026, 9)
    assert output.read_bytes() == b"user data"
    generate_report(report_input, output, 2026, 9, overwrite=True)
    assert output.read_bytes().startswith(b"PK")


def test_long_remarks_preserve_every_customer_and_literal_labels(report_input):
    data = read_report(report_input, 2026, 9)
    # A customer label beginning with '=' must remain plain text in all worksheets.
    bucket = data.customers.pop(("N20", "TSMC", "END_A"))
    data.customers[("N20", "TSMC", "=SUM(1,2)")] = bucket
    for product_key in list(data.products):
        if product_key[:3] == ("N20", "TSMC", "END_A"):
            item = data.products.pop(product_key)
            data.products[("N20", "TSMC", "=SUM(1,2)", product_key[3])] = item
    for i in range(550):
        name = f"Customer_{i:04d}_" + "X" * 60
        data.customers[("N20", "TSMC", name)] = bucket
        data.products[("N20", "TSMC", name, "PN")] = bucket
    book = build_workbook(data)
    rows = [r for r in book.active.iter_rows(min_row=4) if r[0].value in {"N20", "N20（續）"}]
    assert len(rows) > 1
    whole = "".join(r[11].value for r in rows)
    assert "Customer_0000_" in whole and "Customer_0549_" in whole and "=SUM(1,2)" in whole
    assert all(len(r[11].value) <= 30_000 for r in rows)
    literal = next(row[2] for row in book["差異明細"] if row[2].value == "=SUM(1,2)")
    assert literal.data_type == "s"
    book.close()


def test_sorting_magnitude_zero_and_money_rounding():
    values = [("small_positive", Decimal(10)), ("large_negative", Decimal(-80)),
              ("zero", Decimal(0)), ("large_positive", Decimal(30)), ("small_negative", Decimal(-40))]
    order = lambda direction: [name for name, value in sorted(values, key=lambda pair: contribution_order(pair[0], pair[1], Decimal(direction)))]
    assert order(1) == ["large_positive", "small_positive", "large_negative", "small_negative", "zero"]
    assert order(-1) == ["large_negative", "small_negative", "large_positive", "small_positive", "zero"]
    assert money(Decimal("1250000")) == "+1.3M"
    assert money(Decimal("-1250000")) == "-1.3M"
    assert money(Decimal(0)) == "0.0M"
    long = "ABC\n" * 12_000
    assert "".join(_chunks(long)) == long


def test_cli_writes_current_summary_and_error(report_input, tmp_path):
    output, summary = tmp_path / "cli.xlsx", tmp_path / "summary.json"
    arguments = ["--input", str(report_input), "--output", str(output), "--year", "2026", "--month", "9", "--summary-json", str(summary)]
    assert main(arguments) == 0
    assert json.loads(summary.read_text(encoding="utf-8"))["act_rows"] == 20
    assert main(arguments) == 1  # Refuses to overwrite the first output.
    assert "error" in json.loads(summary.read_text(encoding="utf-8"))


def test_all_customers_include_actual_only_budget_only_and_zero_change(report_input):
    book = load_workbook(report_input)
    sheet = book.active
    entries = [("ACT_ONLY", "act", 10_000_000), ("BUDGET_ONLY", "ori", 4_000_000),
               ("ZERO_CHANGE", "act", 1_000_000), ("ZERO_CHANGE", "ori", 1_000_000)]
    for customer, scenario, value in entries:
        row = sheet.max_row + 1
        for column, cell_value in {"AS": value, "BD": f"FY27 {scenario}", "BF": "tsmc", "BM": "N20",
                                   "BI": customer, "BC": customer, "BP": "202609", "BE": 9, "BO": 2026}.items():
            sheet[f"{column}{row}"] = cell_value
    book.save(report_input)
    book.close()
    data = read_report(report_input, 2026, 9)
    output = build_workbook(data)
    remark = get_row(output.active, "N20")[11].value
    assert "ACT_ONLY +10.0M" in remark
    assert "BUDGET_ONLY -4.0M" in remark
    assert "ZERO_CHANGE 0.0M" in remark
    assert remark.index("END_B -30.0M") < remark.index("BUDGET_ONLY -4.0M") < remark.index("ZERO_CHANGE 0.0M")
    output.close()
