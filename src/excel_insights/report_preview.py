"""Read only the report preview; shared by local HTTP and WebAssembly frontends."""
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook

PREVIEW_ROWS = 200


def _display(value, *, percent=False):
    if value is None:
        return ""
    amount = Decimal(str(value)) * (100 if percent else 1)
    return f"{amount.quantize(Decimal(1), rounding=ROUND_HALF_UP):,}" + ("%" if percent else "")


def preview_workbook(path: Path, sheet_name: str = "monthly pivot") -> dict:
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book[sheet_name]
        headers = [[str(cell.value or "") for cell in row] for row in sheet.iter_rows(min_row=1, max_row=3, max_col=12)]
        footnotes = []
        rows = []
        truncated = False
        for row in sheet.iter_rows(min_row=4, max_col=12):
            node = row[0].value
            if not node:
                continue
            if str(node).startswith(("金額：", "註：", "注：")):
                footnotes.append(str(node))
                continue
            if len(rows) >= PREVIEW_ROWS:
                truncated = True
                break
            cells = [str(node)]
            cells.extend(_display(c.value, percent=i in (8, 10)) for i, c in enumerate(row[1:11], 1))
            cells.append(str(row[11].value or ""))
            colors = []
            for i in (8, 10):
                color = row[i].fill.fgColor.rgb
                colors.append("high" if color in ("00FFF0AF", "FFFFF0AF") else
                              "low" if color in ("00C9EFF3", "FFC9EFF3") else "")
            background = row[0].fill.fgColor.rgb
            shade = "blue" if background in ("00EAF3F7", "FFEAF3F7") else "tan" if background in ("00EDE2D1", "FFEDE2D1") else ""
            rows.append({"cells": cells, "tsmc": colors[0], "non_t": colors[1], "shade": shade})
        return {"rows": rows, "truncated": truncated, "headers": headers, "footnotes": footnotes}
    finally:
        book.close()
