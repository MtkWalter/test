"""Single-month sales report engine and CLI; independent of Qt and pandas."""
from __future__ import annotations

import argparse
import calendar
import json
import os
import re
import tempfile
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import cached_property
from pathlib import Path
from typing import Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MILLION = Decimal("1000000")
LOW, HIGH = Decimal("0.85"), Decimal("1.15")
NODE_ORDER = (
    "Reclaim", "N800", "N40", "N30-50", "N28", "N20", "N16", "N14", "N7",
    "N5", "N3/N3E", "N2/A16", "AP-SoIC", "AP-CoWo", "Service", "OT", "ADJ",
)
DEFAULT_ALIASES = {
    "N2": "N2/A16", "A16": "N2/A16", "N3": "N3/N3E", "N3E": "N3/N3E",
}
HEADERS = {
    "amount": ("sales-NTD",), "scenario": ("budget / act",),
    "segment": ("TSMC/non TSMC",), "node": ("Tech Node",),
    "customer": ("END USER",), "product": ("P/N",), "supplier": ("supplier",),
    "ym": ("YM(calender)", "YM(calendar)"),
    "year": ("Year(calender)", "Year(calendar)"),
    "month": ("Month(calender)", "Month(calendar)"), "date": ("日期",),
}
REQUIRED = {"amount", "scenario", "segment", "node", "customer", "product"}
NUMERIC = re.compile(r"^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")
SCENARIO = re.compile(r"^FY(\d{2}|\d{4})\s*(act|ori|rev)$", re.I)


class ReportValidationError(ValueError):
    """All source diagnostics, without a first-N-rows limit."""

    def __init__(self, issues: list[dict]):
        self.issues = issues
        super().__init__(f"來源資料需要修正，尚未產生報表（共 {len(issues):,} 項錯誤）。")

    def __str__(self):
        return super().__str__() + "\n" + "\n".join(
            f"工作表「{i['sheet']}」第 {i['row']} 列 [{i['column']}]：{i['message']}；原始值：{i['value']}"
            for i in self.issues)


def diagnostic_value(value: object) -> str:
    if value is None:
        return "（空白）"
    return str(value)


def text(value: object) -> str:
    return "" if value is None else unicodedata.normalize("NFKC", str(value)).strip()


def key(value: object) -> str:
    return re.sub(r"\s+", "", text(value)).casefold()


@dataclass
class ReportRules:
    """Business defaults, editable without changing the application."""

    tsmc_values: list[str] = field(default_factory=lambda: ["tsmc"])
    # None means all other nonempty BF values, including FI, are Non-T.
    non_t_values: list[str] | None = None
    node_aliases: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_ALIASES))
    fiscal_start_month: int = 4

    def __post_init__(self):
        if not self.tsmc_values or any(not text(v) for v in self.tsmc_values):
            raise ValueError("TSMC 分類不得空白。")
        if self.non_t_values is not None:
            if any(not text(v) for v in self.non_t_values):
                raise ValueError("Non-T 分類不得包含空白。")
            if {key(v) for v in self.tsmc_values} & {key(v) for v in self.non_t_values}:
                raise ValueError("同一個 BF 值不能同時屬於 TSMC 與 Non-T。")
        if any(not text(a) or not text(b) for a, b in self.node_aliases.items()):
            raise ValueError("Tech Node 對照的原始名稱與報表名稱都不得空白。")
        if len({key(a) for a in self.node_aliases}) != len(self.node_aliases):
            raise ValueError("Tech Node 對照不能有重複的原始名稱。")
        if not isinstance(self.fiscal_start_month, int) or not 1 <= self.fiscal_start_month <= 12:
            raise ValueError("會計年度起始月份必須是 1–12。")

    def to_dict(self) -> dict:
        return {"tsmc_values": self.tsmc_values, "non_t_values": self.non_t_values,
                "node_aliases": self.node_aliases, "fiscal_start_month": self.fiscal_start_month}

    @classmethod
    def load(cls, path: str | Path) -> ReportRules:
        return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8-sig")))

    @classmethod
    def from_dict(cls, data: object) -> ReportRules:
        allowed = {"tsmc_values", "non_t_values", "node_aliases", "fiscal_start_month"}
        if not isinstance(data, dict) or set(data) - allowed:
            raise ValueError("分類設定格式不正確，或包含未知的設定名稱。")
        if "node_aliases" in data and (not isinstance(data["node_aliases"], dict) or
                                      any(not isinstance(v, str) for v in data["node_aliases"].values())):
            raise ValueError("node_aliases 必須是名稱對照表。")
        for name in ("tsmc_values", "non_t_values"):
            value = data.get(name)
            if value is not None and (not isinstance(value, list) or any(not isinstance(v, str) for v in value)):
                raise ValueError(f"{name} 必須是文字清單。")
        if data.get("tsmc_values", []) is None:
            raise ValueError("tsmc_values 不能為 null。")
        if "fiscal_start_month" in data and (isinstance(data["fiscal_start_month"], bool) or
                                             not isinstance(data["fiscal_start_month"], int)):
            raise ValueError("會計年度起始月份必須是整數。")
        return cls(**data)

    def segment(self, value: object) -> str:
        if not key(value):
            raise ValueError("BF 客群分類空白，請補上 TSMC／Non-T 分類。")
        if key(value) in {key(v) for v in self.tsmc_values}:
            return "TSMC"
        if self.non_t_values is None or key(value) in {key(v) for v in self.non_t_values}:
            return "Non-T"
        raise ValueError(f"BF 分類「{text(value)}」尚未納入分類設定。")

    def node(self, value: object, supplier: object) -> str:
        if not text(value):
            if key(supplier) == "service":
                return "Service"
            raise ValueError("BM Tech Node 空白，請補上分類或提供對照來源。")
        aliases = {key(a): text(b) for a, b in self.node_aliases.items()}
        mapped = aliases.get(key(value), text(value))
        canonical = {key(v): v for v in (*NODE_ORDER, *self.node_aliases.values())}
        return canonical.get(key(mapped), mapped)


def fiscal_label(year: int, month: int, start: int = 4) -> str:
    ending = year if start == 1 else year + (month >= start)
    return f"FY{ending % 100:02d}"


def amount(value: object) -> Decimal | None:
    if value is None or isinstance(value, str) and not value.strip():
        return None
    if isinstance(value, (bool, date)):
        raise ValueError("AS 必須是數字，不能是布林值或日期。")
    raw = str(value).strip()
    if raw.startswith("(") and raw.endswith(")"):
        raw = "-" + raw[1:-1]
    if not NUMERIC.fullmatch(raw):
        raise ValueError(f"AS 金額不是有效數字：{raw[:60]}")
    try:
        number = Decimal(raw.replace(",", ""))
    except InvalidOperation as exc:
        raise ValueError("AS 金額不是有效數字。") from exc
    if not number.is_finite():
        raise ValueError("AS 金額必須是有限數字。")
    return number


def integer(value: object) -> int:
    if isinstance(value, bool):
        raise ValueError("年份／月份必須是整數。")
    try:
        number = Decimal(str(value))
        if not number.is_finite() or number != number.to_integral_value():
            raise ValueError
        return int(number)
    except (ValueError, InvalidOperation, TypeError, OverflowError) as exc:
        raise ValueError("年份／月份必須是整數。") from exc


def period(values: dict) -> tuple[int, int]:
    candidates = []
    ym = values.get("ym")
    if text(ym):
        if isinstance(ym, (date, datetime)):
            candidates.append((ym.year, ym.month))
        else:
            raw = text(ym)
            if re.fullmatch(r"\d{6}(?:\.0+)?", raw):
                raw = raw[:6]
                candidates.append((int(raw[:4]), int(raw[4:])))
            else:
                match = re.fullmatch(r"(\d{4})[-/](\d{1,2})", raw)
                if not match:
                    raise ValueError(f"BP 年月格式無法辨識：{raw}")
                candidates.append((int(match[1]), int(match[2])))
    year, month = values.get("year"), values.get("month")
    if text(year) or text(month):
        if not text(year) or not text(month):
            raise ValueError("BO 年份、BE 月份必須成對提供。")
        candidates.append((integer(year), integer(month)))
    if not candidates:
        raw_date = values.get("date")
        if isinstance(raw_date, (date, datetime)):
            candidates.append((raw_date.year, raw_date.month))
        elif match := re.match(r"^(\d{4})[-/](\d{1,2})[-/]\d{1,2}(?:\s|$)", text(raw_date)):
            candidates.append((int(match[1]), int(match[2])))
    if not candidates:
        raise ValueError("缺少年月：請提供 BP，或 BO＋BE。")
    if len(set(candidates)) != 1:
        raise ValueError("BP 年月與 BO／BE 不一致。")
    y, m = candidates[0]
    if not 1900 <= y <= 9999 or not 1 <= m <= 12:
        raise ValueError("年份或月份超出有效範圍。")
    return y, m


@dataclass
class SheetLayout:
    name: str
    header_row: int
    columns: dict[str, int]


def _layout(sheet) -> SheetLayout | None:
    matches = {key(alias): name for name, aliases in HEADERS.items() for alias in aliases}
    sheet.reset_dimensions()
    for row_index, row in enumerate(sheet.iter_rows(max_row=20, values_only=True), 1):
        columns = {}
        duplicate = set()
        for index, value in enumerate(row):
            if value is None:
                continue
            name = matches.get(key(value))
            if name:
                if name in columns:
                    duplicate.add(name)
                columns[name] = index
        if REQUIRED <= columns.keys():
            if duplicate:
                raise ValueError(f"工作表「{sheet.title}」有重複欄名：{', '.join(sorted(duplicate))}")
            if not ("ym" in columns or {"year", "month"} <= columns.keys() or "date" in columns):
                continue
            return SheetLayout(sheet.title, row_index, columns)
    return None


def inspect_workbook(path: str | Path) -> list[SheetLayout]:
    path = Path(path)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("請使用 .xlsx 或 .xlsm 輸入檔案。")
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        layouts = [layout for sheet in book if (layout := _layout(sheet))]
        if not layouts:
            raise ValueError("找不到報表來源工作表。必要欄名：sales-NTD、budget / act、TSMC/non TSMC、Tech Node、END USER、P/N，以及年月欄。")
        return layouts
    finally:
        book.close()


@dataclass
class Bucket:
    actual: Decimal = Decimal(0)
    budget: Decimal = Decimal(0)
    actual_rows: int = 0
    budget_rows: int = 0
    missing_budget: bool = False

    def add(self, scenario: str, value: Decimal | None):
        if scenario == "act":
            self.actual_rows += 1
            self.actual += value if value is not None else Decimal(0)
        else:
            self.budget_rows += 1
            if value is None:
                self.missing_budget = True
            else:
                self.budget += value

    @property
    def budget_value(self) -> Decimal | None:
        return None if not self.budget_rows or self.missing_budget else self.budget

    @property
    def delta(self) -> Decimal:
        return self.actual - self.budget


@dataclass
class ReportData:
    source: Path
    sheet: str
    year: int
    month: int
    fiscal_year: str
    rules: ReportRules
    totals: dict[tuple[str, str], Bucket] = field(default_factory=lambda: defaultdict(Bucket))
    customers: dict[tuple[str, str, str], Bucket] = field(default_factory=lambda: defaultdict(Bucket))
    products: dict[tuple[str, str, str, str], Bucket] = field(default_factory=lambda: defaultdict(Bucket))
    counts: dict[str, int] = field(default_factory=lambda: {"act": 0, "ori": 0})
    warnings: list[str] = field(default_factory=list)

    @cached_property
    def customer_groups(self):
        groups = defaultdict(list)
        for (node, segment, name), bucket in self.customers.items():
            groups[(node, segment)].append((name, bucket))
        return groups

    @cached_property
    def product_groups(self):
        groups = defaultdict(list)
        for (node, segment, customer, pn), bucket in self.products.items():
            groups[(node, segment, customer)].append((pn, bucket))
        return groups

    @property
    def nodes(self) -> list[str]:
        seen = {node for node, _ in self.totals}
        # Keep the sample's rows, but add Service only when present.
        base = [n for n in NODE_ORDER if n != "Service" or n in seen]
        extras = sorted(seen - set(NODE_ORDER), key=str.casefold)
        position = base.index("OT")
        return base[:position] + extras + base[position:]


def read_report(path: str | Path, year: int, month: int, *, sheet: str | None = None,
                fiscal_year: str | None = None, rules: ReportRules | None = None,
                progress: Callable[[str], None] | None = None) -> ReportData:
    if not 1900 <= year <= 9999 or not 1 <= month <= 12:
        raise ValueError("請選擇有效的年份與月份。")
    path = Path(path).resolve()
    layouts = inspect_workbook(path)
    if sheet:
        layout = next((s for s in layouts if s.name == sheet), None)
        if layout is None:
            raise ValueError(f"工作表「{sheet}」不存在或缺少必要欄位。")
    elif len(layouts) == 1:
        layout = layouts[0]
    else:
        raise ValueError("有多個可用工作表，請明確選擇工作表：" + "、".join(s.name for s in layouts))
    rules = rules or ReportRules()
    fy = text(fiscal_year or fiscal_label(year, month, rules.fiscal_start_month)).upper()
    if not re.fullmatch(r"FY\d{2}", fy):
        raise ValueError("年度標籤請使用 FY27 這類格式。")
    result = ReportData(path, layout.name, year, month, fy, rules)
    errors: list[dict] = []
    valid_amounts = 0
    warning_counts: dict[str, int] = defaultdict(int)
    book = load_workbook(path, read_only=True, data_only=True)
    # Read formula cells in parallel so missing cached results are never treated as
    # genuine empty budgets. Both iterators stream; no dataframe copy is retained.
    formula_book = load_workbook(path, read_only=True, data_only=False)
    try:
        width = max(layout.columns.values()) + 1
        book[layout.name].reset_dimensions()
        formula_book[layout.name].reset_dimensions()
        cached_rows = book[layout.name].iter_rows(min_row=layout.header_row + 1, max_col=width, values_only=True)
        formula_rows = formula_book[layout.name].iter_rows(min_row=layout.header_row + 1, max_col=width)
        for row_number, (row, formula_row) in enumerate(zip(cached_rows, formula_rows), layout.header_row + 1):
            if progress and row_number % 2000 == 0:
                progress(f"正在讀取第 {row_number:,} 列…")
            raw_scenario = text(row[layout.columns["scenario"]])
            match = SCENARIO.fullmatch(raw_scenario)
            # Also skips the screenshot reconstruction's column-letter row.
            if not match or int(match[1]) % 100 != int(fy[2:]) or match[2].lower() == "rev":
                continue
            values = {name: row[index] for name, index in layout.columns.items()}

            def record_error(fields, exc):
                fields = [f for f in fields if f in layout.columns]
                columns = ", ".join(get_column_letter(layout.columns[f] + 1) for f in fields)
                raw = "; ".join(f"{get_column_letter(layout.columns[f] + 1)}={diagnostic_value(formula_row[layout.columns[f]].value)}"
                                for f in fields)
                errors.append({"sheet": layout.name, "row": row_number, "column": columns,
                               "field": " / ".join(HEADERS[f][0] for f in fields),
                               "value": raw, "message": str(exc)})

            try:
                if period(values) != (year, month):
                    continue
            except ValueError as exc:
                record_error(("ym", "year", "month", "date"), exc)
                continue
            scenario = match[2].lower()
            before = len(errors)
            try:
                index = layout.columns["amount"]
                if formula_row[index].data_type == "f" and values["amount"] is None:
                    raise ValueError("AS 公式沒有已儲存的計算結果，請先以 Excel／LibreOffice 重新計算並儲存。")
                value = amount(values["amount"])
                if value is None and scenario == "act":
                    raise ValueError("AS 實績金額空白；請補齊來源金額，真正的零值請填 0。")
            except ValueError as exc:
                record_error(("amount",), exc)
            try:
                segment = rules.segment(values["segment"])
            except ValueError as exc:
                record_error(("segment",), exc)
            try:
                node = rules.node(values["node"], values.get("supplier"))
            except ValueError as exc:
                record_error(("node", "supplier"), exc)
            if len(errors) != before:
                continue
            if value is None:
                warning_counts["預算金額缺值，所屬分類的預算與達成率顯示空白"] += 1
            else:
                valid_amounts += 1
            if not text(values["node"]):
                warning_counts["Service 的 BM 空白，獨立列為 Service，未分攤到其他節點"] += 1
            customer = text(values["customer"]) or "（未填 END USER）"
            product = text(values["product"]) or "（未填 P/N）"
            if not text(values["customer"]):
                warning_counts["BI END USER 空白，使用「未填 END USER」分組，請補齊來源以正確辨識客戶"] += 1
            if not text(values["product"]):
                warning_counts["BC P/N 空白，使用「未填 P/N」分組"] += 1
            for target in (result.totals[(node, segment)], result.customers[(node, segment, customer)],
                           result.products[(node, segment, customer, product)]):
                target.add(scenario, value)
            result.counts[scenario] += 1
        if errors:
            raise ReportValidationError(errors)
        if not sum(result.counts.values()):
            raise ValueError(f"找不到 {year}/{month:02d} 的 {fy} act／ori 資料，請確認月份、年度標籤與工作表。")
        if not valid_amounts:
            raise ValueError("所選資料的 AS 金額全部缺漏，請補齊來源金額後再產生報表。")
        for scenario, label in (("act", "實績"), ("ori", "原始預算")):
            if not result.counts[scenario]:
                result.warnings.append(f"所選月份沒有{label}明細；請確認來源資料是否完整。")
        result.warnings.extend(f"{message}（{count} 列）" for message, count in sorted(warning_counts.items()))
        return result
    finally:
        book.close()
        formula_book.close()


def ratio(bucket: Bucket | None) -> Decimal | None:
    if bucket is None or bucket.budget_value is None or bucket.budget_value <= 0:
        return None
    return bucket.actual / bucket.budget_value


def triggered(value: Decimal | None) -> bool:
    return value is not None and (value >= HIGH or value <= LOW)


def contribution_order(name: str, delta: Decimal, direction: Decimal):
    if delta == 0:
        priority = 2
    elif direction == 0 or delta * direction > 0:
        priority = 0
    else:
        priority = 1
    return priority, -abs(delta), name.casefold(), name


def money(value: Decimal) -> str:
    rounded = (value / MILLION).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    if rounded == 0:
        rounded = abs(rounded)  # Apply the original direction explicitly below.
    return ("+" if value > 0 else "-" if value < 0 else "") + f"{abs(rounded):,.1f}M"


def remark_lines(data: ReportData, node: str) -> list[str]:
    lines = []
    for segment in ("TSMC", "Non-T"):
        total = data.totals.get((node, segment))
        if not triggered(ratio(total)):
            continue
        lines.append(segment)
        customers = sorted(data.customer_groups[(node, segment)],
                           key=lambda pair: contribution_order(pair[0], pair[1].delta, total.delta))
        for customer, bucket in customers:
            products = sorted(data.product_groups[(node, segment, customer)],
                              key=lambda pair: contribution_order(pair[0], pair[1].delta, bucket.delta))
            details = "；".join(f"{pn} {money(b.delta)}" if not b.missing_budget else f"{pn} 預算缺值"
                               for pn, b in products)
            change = "預算缺值" if bucket.missing_budget else money(bucket.delta)
            lines.append(f"{customer} {change}" + (f"：{details}" if details else ""))
    return lines


def _chunks(value: str, limit: int = 30_000) -> list[str]:
    """Use continuation rows rather than silently truncating Excel's cell limit."""
    chunks = []
    while len(value) > limit:
        boundary = max(value.rfind("\n", 0, limit), value.rfind("；", 0, limit))
        size = boundary + 1 if boundary >= limit // 2 else limit
        chunks.append(value[:size])
        value = value[size:]
    return chunks + [value]


def _literal(cell, value):
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"  # Input labels are text, never Excel formulas.


def _million(value: Decimal | None):
    return None if value is None else float(value / MILLION)


def _delta(bucket: Bucket | None):
    return None if bucket is None or bucket.budget_value is None else bucket.actual - bucket.budget_value


def build_workbook(data: ReportData) -> Workbook:
    book = Workbook()
    sheet = book.active
    sheet.title = "monthly pivot"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A4"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:3"
    sheet.sheet_properties.outlinePr.summaryBelow = False
    for span in ("A1:G1", "H1:I1", "J1:K1", "B2:C2", "D2:E2", "F2:G2", "H2:I2", "J2:K2"):
        sheet.merge_cells(span)
    sheet["A1"] = f"{calendar.month_abbr[data.month]}  |  {data.fiscal_year} Sales & Costs – Monthly Pivot"
    sheet["H1"], sheet["J1"] = "≥ 115% → Yellow", "≤ 85% → Blue"
    sheet["L1"] = "金額：百萬元｜單月"
    top = {"A2": "Tech Node", "B2": "tsmc", "D2": "Non-T", "F2": "Total (service incl.)",
           "H2": "TSMC", "J2": "Non-T", "L2": "Remark"}
    for cell, label in top.items():
        sheet[cell] = label
    labels = ["Tech Node", f"{data.fiscal_year} act", f"{data.fiscal_year} ori",
              f"{data.fiscal_year} act", f"{data.fiscal_year} ori", f"{data.fiscal_year} act",
              f"{data.fiscal_year} ori", f"vs {data.fiscal_year} ori", f"vs {data.fiscal_year} ori %",
              f"vs {data.fiscal_year} ori", f"vs {data.fiscal_year} ori %", "Remark"]
    for col, label in enumerate(labels, 1):
        sheet.cell(3, col, label)
    dark, yellow, blue = "253859", "FFF0AF", "C9EFF3"
    border = Border(bottom=Side(style="hair", color="CDD6E3"))
    for row in sheet.iter_rows(min_row=1, max_row=3, max_col=12):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=dark if cell.row != 2 else "E7EAF2")
            cell.font = Font(name="Calibri", size=11, bold=True,
                             color="FFFFFF" if cell.row != 2 else "263448")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet["H1"].fill = PatternFill("solid", fgColor=yellow)
    sheet["J1"].fill = PatternFill("solid", fgColor=blue)
    sheet["H1"].font = sheet["J1"].font = Font(name="Calibri", size=11, color="263448", bold=True)
    sheet["A1"].fill = PatternFill("solid", fgColor="1C2845")
    sheet["A1"].font = Font(name="Calibri", size=14, color="FFFFFF", bold=True)
    sheet["A1"].alignment = Alignment(horizontal="left", vertical="center")
    sheet["L1"].fill = PatternFill("solid", fgColor="E8EEF5")
    sheet["L1"].font = Font(name="Calibri", size=11, color="53657C")
    sheet["F2"].fill = PatternFill("solid", fgColor="D5EFD5")
    for coordinate in ("H2", "J2"):
        sheet[coordinate].font = Font(name="Calibri", size=11, color="A54141", bold=True)
    row_number = 4
    for node in data.nodes:
        t = data.totals.get((node, "TSMC"))
        n = data.totals.get((node, "Non-T"))
        actual_t = t.actual if t else None
        actual_n = n.actual if n else None
        budget_t = t.budget_value if t else None
        budget_n = n.budget_value if n else None
        total_actual = (actual_t or Decimal(0)) + (actual_n or Decimal(0)) if t or n else None
        total_budget = ((budget_t or Decimal(0)) + (budget_n or Decimal(0))
                        if budget_t is not None or budget_n is not None else None)
        rt, rn = ratio(t), ratio(n)
        values = [node, _million(actual_t), _million(budget_t), _million(actual_n), _million(budget_n),
                  _million(total_actual), _million(total_budget), _million(_delta(t)),
                  float(rt) if rt is not None else None, _million(_delta(n)),
                  float(rn) if rn is not None else None]
        chunks = _chunks("\n".join(remark_lines(data, node)))
        for part, chunk in enumerate(chunks):
            for col in range(1, 13):
                cell = sheet.cell(row_number, col)
                value = values[col - 1] if part == 0 and col <= 11 else chunk if col == 12 else f"{node}（續）" if col == 1 else None
                _literal(cell, value)
                cell.font = Font(name="Calibri", size=11, color="263448")
                cell.fill = PatternFill("solid", fgColor="EAF3F7" if data.nodes.index(node) % 2 == 0 else "EDE2D1")
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=col in (1, 12), horizontal="left" if col in (1, 12) else "right")
                if 2 <= col <= 11:
                    cell.number_format = "0%" if col in (9, 11) else '#,##0;[Red]-#,##0;0'
            if not part:
                for rate, columns in ((rt, (8, 9)), (rn, (10, 11))):
                    if triggered(rate):
                        for col in columns:
                            sheet.cell(row_number, col).fill = PatternFill("solid", fgColor=yellow if rate >= HIGH else blue)
            sheet.row_dimensions[row_number].height = min(409, max(29, (chunk.count("\n") + 1 + len(chunk) // 65) * 15))
            row_number += 1
    row_number += 1
    sheet.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=12)
    sheet.cell(row_number, 1, "金額：百萬元；主表顯示整數、Remark 一位小數。使用原始精度計算；預算≤0或缺值時達成率空白。過長備註接續列示，完整內容可見差異明細。")
    sheet.cell(row_number, 1).alignment = Alignment(wrap_text=True, vertical="top")
    sheet.row_dimensions[row_number].height = 30
    for col, width in enumerate((19, 13, 13, 13, 13, 14, 14, 16, 16, 16, 16, 83), 1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    sheet.row_dimensions[1].height = 34
    sheet.row_dimensions[2].height = 30
    sheet.row_dimensions[3].height = 31
    sheet.print_area = f"A1:L{row_number}"
    details = book.create_sheet("差異明細")
    details.append(["Tech Node", "客群", "END USER", "P/N", "ACT（元）", "BUDGET（元）", "ACT-BUDGET（元）", "需填 Remark"])
    for node in data.nodes:
        for segment in ("TSMC", "Non-T"):
            total = data.totals.get((node, segment))
            customers = sorted(data.customer_groups[(node, segment)],
                               key=lambda pair: contribution_order(pair[0], pair[1].delta, total.delta))
            for customer, bucket in customers:
                products = sorted(data.product_groups[(node, segment, customer)],
                                  key=lambda pair: contribution_order(pair[0], pair[1].delta, bucket.delta))
                for pn, item in products:
                    details.append([node, segment, customer, pn, float(item.actual),
                                    float(item.budget_value) if item.budget_value is not None else None,
                                    float(item.delta) if not item.missing_budget else None,
                                    "是" if triggered(ratio(total)) else "否"])
                    for cell in details[details.max_row]:
                        if isinstance(cell.value, str):
                            cell.data_type = "s"
    details.freeze_panes = "E2"
    details.auto_filter.ref = details.dimensions
    for col, width in zip("ABCDEFGH", (20, 12, 30, 28, 22, 22, 22, 18)):
        details.column_dimensions[col].width = width
    info = book.create_sheet("報表資訊")
    info_rows = [
        ("來源檔案", data.source.name), ("來源工作表", data.sheet),
        ("報表月份", f"{data.year}/{data.month:02d}（單月）"), ("年度標籤", data.fiscal_year),
        ("實績明細列數", data.counts["act"]), ("原始預算明細列數", data.counts["ori"]),
        ("金額來源", "AS / sales-NTD；主表百萬元整數，Remark 百萬元一位小數"),
        ("客戶來源", "僅 BI / END USER；不以客戶簡稱或 CUSTOMER 代替"),
        ("達成率", "ACT/BUDGET；預算≤0或缺值則空白；按未四捨五入比率判斷門檻"),
        ("Remark 門檻", "≥115% 或 ≤85%；TSMC / Non-T 各自判斷"),
        ("客戶排序", "全部列出；與客群總差額同方向者優先，各方向依差額絕對值遞減；零差額最後、同額依名稱排序"),
        ("品號預設", "全部列出；與所屬客戶差額同方向者優先，再依差額絕對值遞減"),
        ("Non-T 分類", "所有非空且非 TSMC 的 BF 值（含 FI）" if data.rules.non_t_values is None else "、".join(data.rules.non_t_values)),
        ("TSMC 分類", "、".join(data.rules.tsmc_values)),
        ("Tech Node 對照", json.dumps(data.rules.node_aliases, ensure_ascii=False)),
        ("其他 Tech Node", "保留原名；Service 的 BM 空白時獨立列為 Service"),
        ("Total", "F=B+D、G=C+E；Service、ADJ 相同；以原始金額相加後顯示"),
    ]
    for row in info_rows:
        info.append(row)
        for cell in info[info.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = "s"
    for warning in data.warnings:
        info.append(["資料提醒", warning])
    info.column_dimensions["A"].width = 24
    info.column_dimensions["B"].width = 115
    for row in info:
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    return book


def generate_report(source: str | Path, output: str | Path, year: int, month: int, *,
                    sheet: str | None = None, fiscal_year: str | None = None,
                    rules: ReportRules | None = None, overwrite: bool = False,
                    progress: Callable[[str], None] | None = None) -> dict:
    source, output = Path(source).resolve(), Path(output).resolve()
    if output.suffix.lower() != ".xlsx":
        raise ValueError("輸出檔案必須是 .xlsx。")
    if source == output or output.exists() and source.exists() and os.path.samefile(source, output):
        raise ValueError("輸出不可覆寫輸入檔案，請選擇另一個檔名。")
    if output.exists() and not overwrite:
        raise FileExistsError("輸出檔案已存在，請選擇新檔名或確認覆寫。")
    data = read_report(source, year, month, sheet=sheet, fiscal_year=fiscal_year, rules=rules, progress=progress)
    if progress:
        progress("正在產生 Excel 與完整差異明細…")
    output.parent.mkdir(parents=True, exist_ok=True)
    book = build_workbook(data)
    try:
        with tempfile.NamedTemporaryFile(prefix=".monthly-", suffix=".xlsx", dir=output.parent, delete=False) as tmp:
            temp = Path(tmp.name)
        try:
            book.save(temp)
            if output.exists() and not overwrite:
                raise FileExistsError("輸出檔案已存在，請選擇新檔名。")
            os.replace(temp, output)
        finally:
            temp.unlink(missing_ok=True)
    finally:
        book.close()
    required = sum(any(triggered(ratio(data.totals.get((node, segment)))) for segment in ("TSMC", "Non-T")) for node in data.nodes)
    return {"output": str(output), "sheet": data.sheet, "year": year, "month": month,
            "fiscal_year": data.fiscal_year, "act_rows": data.counts["act"], "ori_rows": data.counts["ori"],
            "remark_nodes": required, "warnings": data.warnings}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Excel 單月報表產生器：INPUT → monthly pivot OUTPUT")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--year", type=int, required=True, help="西元年，例如 2026")
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--sheet", help="來源工作表；僅一張符合時可省略")
    parser.add_argument("--fiscal-year", help="例如 FY27；預設依 4 月起算")
    parser.add_argument("--rules", type=Path, help="選用的分類設定 JSON")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--summary-json", type=Path)
    args = parser.parse_args(argv)
    try:
        rules = ReportRules.load(args.rules) if args.rules else ReportRules()
        result = generate_report(args.input, args.output, args.year, args.month, sheet=args.sheet,
                                 fiscal_year=args.fiscal_year, rules=rules, overwrite=args.overwrite)
    except Exception as exc:
        result = {"error": str(exc)}
        if args.summary_json:
            args.summary_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(str(exc))
        return 1
    if args.summary_json:
        args.summary_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
