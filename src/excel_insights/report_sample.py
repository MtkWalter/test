"""Generate clearly synthetic, deterministic report data for demos and smoke tests."""
from pathlib import Path

from openpyxl import Workbook


SAMPLE_HEADERS = {
    "C": "客戶簡稱", "AS": "sales-NTD", "BB": "supplier", "BC": "P/N",
    "BD": "budget / act", "BE": "Month(calender)", "BF": "TSMC/non TSMC",
    "BI": "END USER", "BJ": "CUSTOMER", "BM": "Tech Node", "BO": "Year(calender)", "BP": "YM(calender)",
}


def create_sample(path: str | Path) -> Path:
    """Amounts in millions below; written to AS in original NTD units."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "銷售明細"
    for col, label in SAMPLE_HEADERS.items():
        sheet[f"{col}1"] = label
        sheet[f"{col}2"] = col  # Same extra header row as the uploaded input.

    def pair(node, segment, customer, pn, actual, budget, supplier="FT products"):
        for scenario, value in (("act", actual), ("ori", budget)):
            row = sheet.max_row + 1
            cells = {"C": "簡稱並非客戶來源", "BJ": "CUSTOMER並非客戶來源", "BB": supplier,
                     "BC": pn, "BD": f"FY27 {scenario}", "BF": segment, "BI": customer,
                     "BM": node, "BO": 2026, "BE": 9, "BP": "202609",
                     "AS": None if value is None else value * 1_000_000}
            for col, val in cells.items():
                sheet[f"{col}{row}"] = val

    pair("N20", "tsmc", "END_A", "PL_A1", 45, 10)
    pair("N20", "tsmc", "END_A", "PL_A2", 25, 20)
    pair("N20", "tsmc", "END_B", "PL_B_DOWN", -10, 40)
    pair("N20", "tsmc", "END_B", "PL_B_UP", 20, 0)
    pair("N20", "tsmc", "END_C", "PL_C", 40, 30)
    pair("N20", "non-t", "NON_T_DOWN", "PL_D", 30, 80)
    pair("N20", "non-t", "NON_T_UP", "PL_E", 50, 20)
    pair("N2", "tsmc", "END_N2", "PL_N2", 60, 50)
    pair("A16", "tsmc", "END_A16", "PL_A16", 55, 50)
    pair("N3", "tsmc", "END_N3", "PL_N3", 40, 50)
    pair("N3E", "tsmc", "END_N3E", "PL_N3E", 45, 50)
    pair("N5", "tsmc", "NORMAL", "PL_NORMAL", 105, 100)
    pair("N7", "tsmc", "ZERO_BUDGET", "PL_ZERO", 1.5, 0)
    pair("N14", "non-t", "NEGATIVE_BUDGET", "PL_NEG", 20, -10)
    pair("Reclaim", "tsmc", "MISSING_BUDGET", "PL_MISSING", 10, None)
    pair(None, "FI", "SERVICE_USER", "CMP", 5, 2, "Service")
    pair("ADJ", "non-t", "FX", "FX", -1, 2)
    pair("OT", "non-t", "SMALL_BUDGET", "PL_SMALL", 0.6, 0.4)
    pair("N40", "tsmc", "BELOW_HIGH", "PL_NEAR_HIGH", 114.9999, 100)
    pair("N28", "tsmc", "ABOVE_LOW", "PL_NEAR_LOW", 85.0001, 100)
    # Out-of-period and rev values must never leak into the selected report.
    for scenario, ym in (("FY27 act", "202608"), ("FY27 rev", "202609"), ("FY26 act", "202509")):
        row = sheet.max_row + 1
        for col, value in {"AS": 9_999_000_000, "BD": scenario, "BP": ym, "BO": int(ym[:4]),
                           "BE": int(ym[4:]), "BF": "tsmc", "BM": "N20", "BI": "EXCLUDED", "BC": "EXCLUDED"}.items():
            sheet[f"{col}{row}"] = value
    sheet.freeze_panes = "D3"
    sheet.column_dimensions["AS"].width = 20
    sheet.column_dimensions["BI"].width = 25
    note = book.create_sheet("範例說明")
    note.append(["這是合成測試資料，不是使用者的真實銷售數字。"])
    note.append(["選擇 2026 年 9 月、FY27、銷售明細。"])
    note.append(["N20：TSMC ACT=120M、BUDGET=100M；Non-T ACT=80M、BUDGET=100M。"])
    book.save(path)
    book.close()
    return path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="建立合成的單月報表範例 INPUT")
    parser.add_argument("output", type=Path)
    print(create_sample(parser.parse_args().output))
