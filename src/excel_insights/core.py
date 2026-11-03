"""Workbook ingestion and analysis, independent of the desktop UI."""
from __future__ import annotations

import csv
import math
import numbers
import re
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Iterable

import pandas as pd

MAX_ROWS = 100_000
MAX_COLUMNS = 512
MAX_CELLS = 1_000_000
TOTAL_CELL_LIMIT = 2_000_000
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsb"}
NUMBER_PATTERN = re.compile(
    r"^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$"
)


def is_blank(value: object) -> bool:
    if value is None or isinstance(value, str) and not value.strip():
        return True
    return bool(pd.isna(value))


def display(value: object) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, (datetime, date, time)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def number(value: object) -> float | None:
    if is_blank(value) or isinstance(value, (bool, date, time)) or pd.api.types.is_bool(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        if not NUMBER_PATTERN.fullmatch(value):
            return None
        value = value.replace(",", "")
    elif not isinstance(value, (numbers.Real, Decimal)):
        return None
    try:
        parsed = float(value)
    except (ValueError, OverflowError):
        return None
    return parsed if math.isfinite(parsed) else None


@dataclass
class Sheet:
    path: Path
    name: str
    data: pd.DataFrame

    @property
    def label(self) -> str:
        return f"{self.path.name} / {self.name}（{len(self.data):,} 列）"

    @property
    def cells(self) -> int:
        return self.data.shape[0] * self.data.shape[1]


def unique_headers(values: Iterable[object]) -> list[str]:
    names: list[str] = []
    used: set[str] = set()
    for index, value in enumerate(values, 1):
        original = display(value).strip() or f"欄位 {index}"
        name = original
        suffix = 2
        while name.casefold() in used:
            name = f"{original} ({suffix})"
            suffix += 1
        used.add(name.casefold())
        names.append(name)
    return names


class _Collector:
    """Enforce file-wide limits while streaming, before building dataframes."""

    def __init__(self, path: Path):
        self.path = path
        self.scanned_cells = 0
        self.sheets: list[Sheet] = []

    def consume(self, name: str, rows: Iterable[Iterable[object]]) -> None:
        records: list[list[object]] = []
        width = 0
        for row_number, raw in enumerate(rows):
            row = list(raw)
            # Ignore unused trailing cells; retain later columns with actual data.
            while row and is_blank(row[-1]):
                row.pop()
            width = max(width, len(row))
            self.scanned_cells += max(width, 1)
            if width > MAX_COLUMNS:
                raise ValueError(f"工作表「{name}」超過 {MAX_COLUMNS} 欄限制。")
            if self.scanned_cells > MAX_CELLS or row_number > MAX_ROWS:
                raise ValueError(f"檔案超過上限：每表 {MAX_ROWS:,} 列、每檔 {MAX_CELLS:,} 儲存格。")
            retained_rows = max(0, len(records) - 1) + (1 if row_number > 0 else 0)
            if retained_rows * width + sum(sheet.cells for sheet in self.sheets) > MAX_CELLS:
                raise ValueError(f"檔案超過 {MAX_CELLS:,} 儲存格限制。")
            if row_number == 0 or any(not is_blank(value) for value in row):
                records.append(row)
        if not records or not width:
            return
        header = records[0] + [None] * (width - len(records[0]))
        body = [row + [None] * (width - len(row)) for row in records[1:]]
        # Explicit object dtype preserves booleans, dates, literal NA and numeric text.
        data = pd.DataFrame(body, columns=unique_headers(header), dtype=object)
        self.sheets.append(Sheet(self.path, name, data))


def read_workbook(path: str | Path) -> list[Sheet]:
    path = Path(path).resolve()
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError("只支援 .xlsx、.xls、.xlsb 檔案。")
    collector = _Collector(path)
    if suffix == ".xlsx":
        from openpyxl import load_workbook

        book = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in book.worksheets:
                if sheet.max_column and sheet.max_column > MAX_COLUMNS:
                    raise ValueError(f"工作表「{sheet.title}」超過 {MAX_COLUMNS} 欄限制（包含格式範圍）。")
                if sheet.max_row and sheet.max_row > MAX_ROWS + 1:
                    raise ValueError(f"工作表「{sheet.title}」超過 {MAX_ROWS:,} 列限制（包含空白範圍）。")
                collector.consume(sheet.title, sheet.iter_rows(values_only=True))
        finally:
            book.close()
    elif suffix == ".xls":
        import xlrd

        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            for index in range(book.nsheets):
                sheet = book.sheet_by_index(index)
                if sheet.ncols > MAX_COLUMNS or sheet.nrows > MAX_ROWS + 1:
                    raise ValueError("工作表超過列數或欄數限制。")

                def rows():
                    for r in range(sheet.nrows):
                        result = []
                        for cell in sheet.row(r):
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                value = xlrd.xldate_as_datetime(cell.value, book.datemode)
                            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                                value = bool(cell.value)
                            elif cell.ctype == xlrd.XL_CELL_ERROR:
                                value = xlrd.error_text_from_code.get(cell.value, "#ERROR")
                            else:
                                value = cell.value
                            result.append(value)
                        yield result

                collector.consume(sheet.name, rows())
                book.unload_sheet(index)
        finally:
            book.release_resources()
    else:
        from pyxlsb import open_workbook

        with open_workbook(str(path)) as book:
            for name in book.sheets:
                with book.get_sheet(name) as sheet:
                    dimension = sheet.dimension
                    if dimension.c + dimension.w > MAX_COLUMNS or dimension.r + dimension.h > MAX_ROWS + 1:
                        raise ValueError("工作表超過列數或欄數限制。")
                    # sparse=False preserves physical row positions, including the header.
                    def rows():
                        for raw in sheet.rows(sparse=False):
                            if raw and max(cell.c for cell in raw) >= MAX_COLUMNS:
                                raise ValueError(f"工作表超過 {MAX_COLUMNS} 欄限制。")
                            values = [None] * (max((cell.c for cell in raw), default=-1) + 1)
                            for cell in raw:
                                values[cell.c] = cell.v
                            yield values

                    collector.consume(name, rows())
    return collector.sheets


def filter_rows(data: pd.DataFrame, column: str | None, query: str) -> pd.DataFrame:
    if not query.strip():
        return data
    if column not in data.columns:
        raise ValueError("請選擇有效的篩選欄位。")
    needle = query.casefold()
    return data.loc[data[column].map(lambda value: needle in display(value).casefold())]


SUMMARY_COLUMNS = ["群組", "資料筆數", "有效數值", "空白", "非數值", "總和", "平均", "最小值", "最大值"]


def analyze(data: pd.DataFrame, numeric_column: str, group_column: str | None = None) -> pd.DataFrame:
    if numeric_column not in data.columns or group_column is not None and group_column not in data.columns:
        raise ValueError("請選擇有效的統計欄位。")
    groups: dict[tuple[bool, str], list[object]] = {}
    values = data[numeric_column].tolist()
    group_values = data[group_column].tolist() if group_column is not None else ["全部"] * len(data)
    for group, value in zip(group_values, values):
        key = (is_blank(group), display(group))
        groups.setdefault(key, []).append(value)
    summaries = []
    for (blank_group, group), members in groups.items():
        valid = [parsed for value in members if (parsed := number(value)) is not None]
        blanks = sum(is_blank(value) for value in members)
        total = math.fsum(valid) if valid else None
        if total is not None and not math.isfinite(total):
            raise OverflowError("數值總和超出可計算範圍。")
        label = "（空白）" if blank_group else "（文字：空白）" if group == "（空白）" else group
        summaries.append([label, len(members), len(valid), blanks, len(members) - len(valid) - blanks,
                          total, total / len(valid) if valid else None,
                          min(valid) if valid else None, max(valid) if valid else None])
    return pd.DataFrame(summaries, columns=SUMMARY_COLUMNS, dtype=object)


def export_csv(path: str | Path, summary: pd.DataFrame, *, file: str, sheet: str,
               column: str, group: str = "不分組", query: str = "無") -> None:
    def safe(value: object) -> object:
        if is_blank(value):
            return ""
        if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
            return "'" + value
        return value

    with Path(path).open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["檔案", "工作表", "數值欄位", "分組欄位", "篩選條件", *SUMMARY_COLUMNS])
        for row in summary.itertuples(index=False, name=None):
            writer.writerow([safe(value) for value in [file, sheet, column, group, query, *row]])
