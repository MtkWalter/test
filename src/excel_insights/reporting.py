"""Editable local report rules, aggregation, settings and template output."""
from __future__ import annotations

import calendar
import json
import math
from copy import copy
from dataclasses import asdict, dataclass, field as dc_field
from datetime import date, datetime
from importlib.resources import files
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

from .core import Sheet, display, is_blank, number
from .schema import FIELDS, normalize, schema_key, suggest_mapping

CATEGORIES = ("TSMC", "Non-T", "FI")
REPORT_COLUMNS = ["Tech Node", "TSMC act", "TSMC ori", "Non-T act", "Non-T ori",
                  "FI act", "FI ori", "Total act", "Total ori", "TSMC 差額", "TSMC 比較 %",
                  "Non-T 差額", "Non-T 比較 %", "FI 差額", "Remark"]
YELLOW = "FFFFFFCC"
BLUE = "FFCCFFFF"
MAX_REPORT_ROWS = 5000


@dataclass
class Rule:
    field: str = "category"
    operator: str = "equals"
    values: list[str] = dc_field(default_factory=list)


@dataclass
class ReportConfig:
    mappings: dict[str, dict[str, str | None]] = dc_field(default_factory=dict)
    actual: str = ""
    original: str = ""
    rules: dict[str, Rule] = dc_field(default_factory=lambda: {c: Rule() for c in CATEGORIES})
    filters: dict[str, str] = dc_field(default_factory=dict)
    year_mode: str = "all"
    year: int = 2027
    fiscal_start: int = 4
    month: int = 0
    difference: str = "act-ori"
    percentage: str = "ratio"
    upper: float = 115.0
    lower: float = 85.0
    multiplier: float = 1.0
    display_unit: int = 1_000_000
    decimals: int = 0
    total_mode: str = "all"
    total_row: bool = False
    label: str = ""

    def validate(self):
        if self.year_mode not in {"all", "calendar", "fiscal"} or self.difference not in {"act-ori", "ori-act"}:
            raise ValueError("年度或差額設定無效。")
        if self.percentage not in {"ratio", "change"} or self.total_mode not in {"all", "classified"}:
            raise ValueError("比率或 Total 設定無效。")
        if not 1900 <= self.year <= 9999 or not 1 <= self.fiscal_start <= 12 or not 0 <= self.month <= 12:
            raise ValueError("年度或月份超出範圍。")
        if not all(math.isfinite(v) for v in (self.upper, self.lower, self.multiplier)):
            raise ValueError("數值設定必須是有限數。")
        if self.upper <= self.lower or self.multiplier <= 0 or self.display_unit not in {1, 1000, 1_000_000}:
            raise ValueError("黃色門檻須大於藍色門檻；倍率須大於 0。")
        if not isinstance(self.decimals, int) or not 0 <= self.decimals <= 6:
            raise ValueError("金額小數位必須介於 0 與 6。")
        if set(self.rules) != set(CATEGORIES):
            raise ValueError("分類設定不完整。")
        for rule in self.rules.values():
            if rule.field not in {"category", "enduser", "product", "supplier", "part"} or rule.operator not in {"equals", "contains"}:
                raise ValueError("分類條件無效。")
            if not isinstance(rule.values, list) or not all(isinstance(v, str) for v in rule.values):
                raise ValueError("分類值必須是文字清單。")
        if any(key not in {"supplier", "part", "product"} or not isinstance(v, str) for key, v in self.filters.items()):
            raise ValueError("篩選設定無效。")
        for mapping in self.mappings.values():
            if not isinstance(mapping, dict) or any(k not in FIELDS or v is not None and not isinstance(v, str) for k, v in mapping.items()):
                raise ValueError("欄位對應設定無效。")

    def mapping(self, sheet: Sheet) -> dict[str, str | None]:
        return self.mappings.get(schema_key(sheet.data.columns), suggest_mapping(sheet.data.columns))

    @property
    def period_label(self):
        if self.label.strip():
            return self.label.strip()
        if self.year_mode == "fiscal":
            return f"FY{self.year % 100:02d}"
        return f"CY{self.year}" if self.year_mode == "calendar" else "全部年度"


def save_config(path: str | Path, config: ReportConfig):
    config.validate()
    Path(path).write_text(json.dumps({"version": 1, "settings": asdict(config)}, ensure_ascii=False, indent=2), encoding="utf-8")


def load_config(path: str | Path) -> ReportConfig:
    path = Path(path)
    if path.stat().st_size > 1_000_000:
        raise ValueError("設定檔過大。")
    try:
        document = json.loads(path.read_text(encoding="utf-8-sig"))
        if document["version"] != 1:
            raise ValueError("不支援的設定版本。")
        values = document["settings"]
        values["rules"] = {key: Rule(**rule) for key, rule in values["rules"].items()}
        config = ReportConfig(**values)
        config.validate()
        return config
    except (KeyError, TypeError, AttributeError) as exc:
        raise ValueError("設定檔格式不正確。") from exc


def observed_values(sheets: list[Sheet], config: ReportConfig, field_name: str) -> list[str]:
    values = set()
    for sheet in sheets:
        column = config.mapping(sheet).get(field_name)
        if column in sheet.data.columns:
            values.update(display(value).strip() for value in sheet.data[column] if not is_blank(value))
    return sorted(values, key=str.casefold)


def suggest_config(sheets: list[Sheet]) -> ReportConfig:
    config = ReportConfig()
    for sheet in sheets:
        config.mappings[schema_key(sheet.data.columns)] = suggest_mapping(sheet.data.columns)
    scenarios = observed_values(sheets, config, "scenario")
    config.actual = next((v for v in scenarios if normalize(v) in {"act", "actual", "實際"}), "")
    config.original = next((v for v in scenarios if normalize(v) in {"ori", "original", "budget", "預算"}), "")
    categories = observed_values(sheets, config, "category")
    aliases = {"TSMC": {"tsmc", "台積電"}, "Non-T": {"nontsmc", "nont", "非台積電"}, "FI": {"fi"}}
    for name in CATEGORIES:
        config.rules[name].values = [v for v in categories if normalize(v) in aliases[name]]
    return config


def _month(value):
    numeric = number(value)
    if numeric is not None and numeric.is_integer() and 1 <= numeric <= 12:
        return int(numeric)
    text = display(value).strip().casefold().removesuffix("月")
    if text.isdigit() and 1 <= int(text) <= 12:
        return int(text)
    for m in range(1, 13):
        if text in {calendar.month_abbr[m].casefold(), calendar.month_name[m].casefold()}:
            return m
    return None


def _year(value):
    numeric = number(value)
    return int(numeric) if numeric is not None and numeric.is_integer() and 1900 <= numeric <= 9999 else None


def _period(row):
    year, month = _year(row.get("year")), _month(row.get("month"))
    value = row.get("date")
    if year is None or month is None:
        if isinstance(value, (datetime, date)):
            year, month = year or value.year, month or value.month
        elif isinstance(value, str) and value.strip():
            try:
                parsed = pd.Timestamp(value)
                if not pd.isna(parsed):
                    year, month = year or parsed.year, month or parsed.month
            except (ValueError, TypeError):
                pass
    return year, month


def _matches(row, rule):
    text = display(row.get(rule.field)).strip().casefold()
    wanted = [v.strip().casefold() for v in rule.values if v.strip()]
    if not text or not wanted:
        return False
    return text in wanted if rule.operator == "equals" else any(value in text for value in wanted)


def comparison(actual, original, config):
    if actual is None or original is None:
        return None, None
    delta = actual - original if config.difference == "act-ori" else original - actual
    rate = None if original == 0 else actual / original if config.percentage == "ratio" else delta / original
    if not math.isfinite(delta) or rate is not None and not math.isfinite(rate):
        raise OverflowError("比較數值超出可計算範圍。")
    return delta, rate


@dataclass
class ReportResult:
    frame: pd.DataFrame
    warnings: list[str]
    counts: dict[str, int]
    sources: list[str]


def build_report(sheets: list[Sheet], config: ReportConfig) -> ReportResult:
    config.validate()
    if not config.actual or not config.original or config.actual.strip().casefold() == config.original.strip().casefold():
        raise ValueError("請選擇不同的 act 與 ori 來源值。")
    counts = dict(input=0, filtered=0, invalid_period=0, invalid_amount=0, blank_node=0, unknown_scenario=0, overlap=0, used=0)
    groups = {}
    category_counts = {name: 0 for name in CATEGORIES}
    sources = []
    warnings = []
    for name, rule in config.rules.items():
        if not rule.values:
            warnings.append(f"{name} 尚未指定分類值，該分類輸出留白。")
    for sheet in sheets:
        mapping = config.mapping(sheet)
        for required in ("node", "amount", "scenario"):
            if mapping.get(required) not in sheet.data.columns:
                raise ValueError(f"{sheet.label}：尚未對應 {required} 欄位。")
        for field_name, value in config.filters.items():
            if value and mapping.get(field_name) not in sheet.data.columns:
                raise ValueError(f"{sheet.label}：無法套用 {field_name} 篩選，請先對應欄位。")
        for name, rule in config.rules.items():
            if rule.values and mapping.get(rule.field) not in sheet.data.columns:
                raise ValueError(f"{sheet.label}：{name} 使用的分類欄位尚未對應。")
        sources.append(f"{sheet.path.name} / {sheet.name}（標題列 {sheet.header_row}）")
        positions = {field_name: sheet.data.columns.get_loc(column) for field_name, column in mapping.items() if column in sheet.data.columns}
        for raw in sheet.data.itertuples(index=False, name=None):
            counts["input"] += 1
            row = {key: raw[index] for key, index in positions.items()}
            if any(value and display(row.get(key)).strip().casefold() != value.strip().casefold() for key, value in config.filters.items()):
                counts["filtered"] += 1
                continue
            if config.year_mode != "all" or config.month:
                year, month = _period(row)
                if config.month and month is None or config.year_mode != "all" and year is None or config.year_mode == "fiscal" and month is None:
                    counts["invalid_period"] += 1
                    continue
                fiscal = year + (1 if config.fiscal_start != 1 and month is not None and month >= config.fiscal_start else 0) if year else None
                test_year = fiscal if config.year_mode == "fiscal" else year
                if config.month and month != config.month or config.year_mode != "all" and test_year != config.year:
                    counts["filtered"] += 1
                    continue
            scenario = display(row.get("scenario")).strip().casefold()
            if scenario not in {config.actual.strip().casefold(), config.original.strip().casefold()}:
                counts["unknown_scenario"] += 1
                continue
            node = display(row.get("node")).strip()
            if not node:
                counts["blank_node"] += 1
                continue
            amount = number(row.get("amount"))
            if amount is None:
                counts["invalid_amount"] += 1
                continue
            amount *= config.multiplier
            if not math.isfinite(amount):
                raise OverflowError("來源金額乘上倍率後超出範圍。")
            counts["used"] += 1
            bucket = groups.setdefault(node, {"values": [[] for _ in range(8)], "remarks": []})
            version = 0 if scenario == config.actual.strip().casefold() else 1
            matches = [_matches(row, config.rules[name]) for name in CATEGORIES]
            if sum(matches) > 1:
                counts["overlap"] += 1
            for i, matched in enumerate(matches):
                if matched:
                    bucket["values"][i * 2 + version].append(amount)
                    category_counts[CATEGORIES[i]] += 1
            if config.total_mode == "all" or any(matches):
                bucket["values"][6 + version].append(amount)
            remark = display(row.get("remark")).strip()
            if remark and remark not in bucket["remarks"]:
                bucket["remarks"].append(remark)
    def total(values):
        result = math.fsum(values) if values else None
        if result is not None and not math.isfinite(result):
            raise OverflowError("加總超出可計算範圍。")
        return result

    def make_row(node, values, remarks):
        t, n, f = (comparison(values[i], values[i + 1], config) for i in (0, 2, 4))
        zero_notes = [f"{name} ori=0，比率留白" for i, name in enumerate(CATEGORIES) if values[i * 2] is not None and values[i * 2 + 1] == 0]
        return [node, *values, t[0], t[1], n[0], n[1], f[0], "；".join([*remarks, *zero_notes])[:32000]]
    rows = [make_row(node, [total(v) for v in bucket["values"]], bucket["remarks"]) for node, bucket in groups.items()]
    if config.total_row and rows:
        rows.append(make_row("總計（報表）", [total([row[i + 1] for row in rows if row[i + 1] is not None]) for i in range(8)], []))
    if len(rows) > MAX_REPORT_ROWS:
        raise ValueError(f"報表超過 {MAX_REPORT_ROWS:,} 列，請縮小篩選範圍。")
    messages = {"invalid_period": "年度或月份無法判讀", "invalid_amount": "金額空白或不是有效數值", "blank_node": "Tech Node 空白", "unknown_scenario": "不是所選 act／ori 版本"}
    for key, message in messages.items():
        if counts[key]:
            warnings.append(f"{counts[key]:,} 筆{message}，已略過。")
    if counts["overlap"]:
        warnings.append(f"{counts['overlap']:,} 筆同時符合多個分類；Total 每筆只加一次。")
    for name, rule in config.rules.items():
        if rule.values and not category_counts[name]:
            warnings.append(f"{name} 未匹配任何有效明細，請檢查分類欄位或選取值。")
    if not rows:
        warnings.append("沒有可輸出的資料。請檢查來源是否只有欄名、版本值與篩選設定。")
    return ReportResult(pd.DataFrame(rows, columns=REPORT_COLUMNS, dtype=object), warnings, counts, sources)


def export_report(path: str | Path, result: ReportResult, config: ReportConfig, sources: list[Sheet]):
    config.validate()
    if result.frame.empty:
        raise ValueError("沒有可匯出的資料。")
    path = Path(path).resolve()
    template = Path(str(files("excel_insights").joinpath("assets/output_template.xlsx")))
    if path == template.resolve() or any(path == source.path.resolve() for source in sources):
        raise ValueError("請選擇新檔名，不能覆寫來源檔案或內建範本。")
    book = load_workbook(template)
    sheet = book.active
    last = max(21, 11 + len(result.frame))
    for row in range(22, last + 1):
        for col in range(4, 19):
            original = sheet.cell(13, col)
            cell = sheet.cell(row, col)
            cell._style = copy(original._style)
        sheet.row_dimensions[row].height = sheet.row_dimensions[13].height
    def write(cell, value):
        cell.value = None if is_blank(value) else value
        if isinstance(value, str):
            cell.data_type = "s"  # Never execute source text as Excel formulas.
    write(sheet["E9"], calendar.month_abbr[config.month] if config.month else "All months")
    for col in range(5, 13):
        write(sheet.cell(11, col), f"{config.period_label} {'act' if col % 2 else 'ori'}")
    difference_label = "act − ori" if config.difference == "act-ori" else "ori − act"
    rate_label = "act / ori %" if config.percentage == "ratio" else "差額 / ori %"
    for col, value in {13: difference_label, 14: rate_label, 15: difference_label, 16: rate_label, 17: difference_label}.items():
        write(sheet.cell(11, col), value)
    write(sheet["M9"], f"≥ {config.upper:g}% → yellow")
    write(sheet["O9"], f"≤ {config.lower:g}% → blue")
    digits = "." + "0" * config.decimals if config.decimals else ""
    scaling = {1: "", 1000: ",", 1_000_000: ",,"}[config.display_unit]
    amount_format = f'#,##0{digits}{scaling};[Red]-#,##0{digits}{scaling}'
    for r, values in enumerate(result.frame.itertuples(index=False, name=None), 12):
        for c, value in enumerate(values, 4):
            write(sheet.cell(r, c), value)
            if c in {14, 16}:
                sheet.cell(r, c).number_format = '0%'
            elif 5 <= c <= 17:
                sheet.cell(r, c).number_format = amount_format
        if values[0] == "總計（報表）":
            for cell in sheet[r][3:18]:
                cell.fill = PatternFill("solid", fgColor="FFE2F0D9")
    end = 11 + len(result.frame)
    for start, rate in (("M", "N"), ("O", "P")):
        for operator, threshold, color in ((">=", config.upper / 100, YELLOW), ("<=", config.lower / 100, BLUE)):
            sheet.conditional_formatting.add(f"{start}12:{rate}{end}", FormulaRule(formula=[f'AND(ISNUMBER(${rate}12),${rate}12{operator}{threshold})'], fill=PatternFill("solid", fgColor=color)))
    fi_rate = "$I12/$J12" if config.percentage == "ratio" else "($I12-$J12)/$J12" if config.difference == "act-ori" else "($J12-$I12)/$J12"
    for operator, threshold, color in ((">=", config.upper / 100, YELLOW), ("<=", config.lower / 100, BLUE)):
        sheet.conditional_formatting.add(f"Q12:Q{end}", FormulaRule(formula=[f'IFERROR(AND(ISNUMBER($I12),ISNUMBER($J12),$J12<>0,{fi_rate}{operator}{threshold}),FALSE)'], fill=PatternFill("solid", fgColor=color)))
    sheet.freeze_panes = "E12"
    sheet.sheet_view.showGridLines = False
    sheet.print_area = f"D9:R{last}"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    notes = book.create_sheet("設定與檢查")
    lines = ["本報表為本機匯出快照；修改 Excel 數值不會重新計算，請在程式修改設定後重新匯出。",
             f"表頭：{config.period_label}；年度模式：{config.year_mode}；年度：{config.year}；FY 起始月：{config.fiscal_start}；月份：{config.month or '全部'}。",
             f"金額倍率：{config.multiplier}；顯示除數：{config.display_unit}（儲存格仍保留原始換算後數值）。",
             f"差額：{config.difference}；百分比：{config.percentage}；Total：{config.total_mode}。",
             "Total all 包含所有符合期間／篩選條件且版本為 act／ori 的有效明細（含 service／未分類）；classified 僅包含分類聯集。",
             *result.sources, *result.warnings, json.dumps(result.counts, ensure_ascii=False),
             "完整設定（無原始交易列）：", json.dumps(asdict(config), ensure_ascii=False)]
    note_row = 1
    for value in lines:
        for offset in range(0, max(1, len(value)), 30000):
            write(notes.cell(note_row, 1), value[offset:offset + 30000])
            note_row += 1
    notes.column_dimensions["A"].width = 110
    book.save(path)
    book.close()


def demo_sheet() -> Sheet:
    rows = []
    for i, node in enumerate(("2nm", "3nm", "5nm", "7nm", "16nm", "28nm"), 1):
        for category, base, ratio in (("TSMC", 100, 1.2 if i % 2 else .8), ("Non-T", 50, .8 if i % 2 else 1.2), ("FI", 20, 1.05), ("Service", 10, 1.0)):
            for scenario, value in (("ori", base), ("act", base * ratio)):
                rows.append([node, value * i * 1_000_000, scenario, category, "Demo End User", "Service" if category == "Service" else "Product", "Demo Supplier", "DEMO-PART", 2026, 9, "示範資料，非實際業績"])
    columns = ["Tech Node", "sales-NTD", "budget / act", "TSMC/non TSMC", "END USER", "Prod", "suppiler", "P/N", "Year(calender)", "Month(calender)", "Remark"]
    return Sheet(Path("示範資料.xlsx"), "示範（非實際業績）", pd.DataFrame(rows, columns=columns, dtype=object), 4)
