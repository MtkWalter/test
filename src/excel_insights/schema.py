"""Local, deterministic suggestions; no network or remote model is involved."""
import hashlib
import re

FIELDS = {
    "node": ("Tech Node", "tech node", "技術節點", "製程", "Node"),
    "amount": ("sales-NTD", "本幣銷貨金額", "銷售金額", "金額", "Amount"),
    "scenario": ("budget / act", "budget/act", "scenario", "預算/實際", "版本"),
    "category": ("TSMC/non TSMC", "客戶分類", "Category"),
    "enduser": ("END USER", "EndUser", "終端客戶"),
    "customer": ("CUSTOMER", "客戶簡稱", "客戶名稱", "客戶代號"),
    "product": ("Prod", "產品類別", "Product"),
    "supplier": ("suppiler", "supplier", "供應商"),
    "part": ("P/N", "料號", "品號"),
    "year": ("Year(calender)", "Year(calendar)", "年度", "Year"),
    "month": ("Month(calender)", "Month(calendar)", "月份", "Month"),
    "date": ("日期", "Date"),
    "remark": ("Remark", "備註"),
}
FIELD_LABELS = {
    "node": "分組／Tech Node", "amount": "加總金額", "scenario": "act／ori 版本",
    "category": "客戶分類", "enduser": "END USER", "customer": "客戶（Remark 彙總）", "product": "產品類別",
    "supplier": "供應商", "part": "料號", "year": "曆年", "month": "月份",
    "date": "日期（年度／月份備援）", "remark": "原始備註",
}


def normalize(value: str) -> str:
    return re.sub(r"[\W_]+", "", value.casefold())


def suggest_mapping(columns) -> dict[str, str | None]:
    result = {}
    for field, aliases in FIELDS.items():
        result[field] = None
        for alias in aliases:
            matches = [str(c) for c in columns if normalize(str(c)) == normalize(alias)]
            if len(matches) == 1:
                result[field] = matches[0]
                break
    return result


def suggest_header(rows) -> int:
    known = {normalize(alias) for aliases in FIELDS.values() for alias in aliases}
    scored = []
    for index, row in enumerate(rows):
        matches = {normalize(v) for v in row if isinstance(v, str)} & known
        scored.append((len(matches), index))
    # Keep the first row for unfamiliar layouts; suggestions need recognizable headers.
    if not scored or max(score for score, _ in scored) < 2:
        return 1
    return max(scored, key=lambda item: (item[0], -item[1]))[1] + 1


def schema_key(columns) -> str:
    return hashlib.sha256("\x1f".join(map(str, columns)).encode()).hexdigest()
