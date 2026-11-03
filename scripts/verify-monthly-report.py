"""Independent assertions for the workbook written by the frozen executable."""
import sys
from openpyxl import load_workbook


def verify(path):
    book = load_workbook(path, data_only=True)
    try:
        sheet = book["monthly pivot"]
        rows = {row[0].value: row for row in sheet.iter_rows(min_row=4) if row[0].value}
        row = rows["N20"]
        expected = ["N20", 120, 100, 80, 100, 200, 200, 20, 1.2, -20, .8]
        if [cell.value for cell in row[:11]] != expected:
            raise ValueError("Frozen report N20 amounts or ratios do not match the sample.")
        remark = row[11].value
        if not (remark.index("END_A +40.0M") < remark.index("END_C +10.0M") < remark.index("END_B -30.0M")):
            raise ValueError("Frozen report customer ordering is incorrect.")
        if not (remark.index("NON_T_DOWN -50.0M") < remark.index("NON_T_UP +30.0M")):
            raise ValueError("Frozen report Non-T ordering is incorrect.")
        for label, expected_ratio in (("N2/A16", 1.15), ("N3/N3E", .85)):
            if rows[label][8].value != expected_ratio or not rows[label][11].value:
                raise ValueError("Boundary ratios or required remarks are incorrect.")
        if rows["N7"][8].value is not None or rows["N14"][10].value is not None:
            raise ValueError("Nonpositive budgets must produce blank ratios.")
        if rows["N5"][11].value is not None or rows["N40"][11].value is not None:
            raise ValueError("Normal / rounded-only ratios must not trigger remarks.")
        if len(list(book["差異明細"].iter_rows(min_row=2))) != 20:
            raise ValueError("Frozen report lost product details.")
        print("Validated report: amounts, both remark groups, boundaries, ordering and blanks.")
    finally:
        book.close()


if __name__ == "__main__":
    verify(sys.argv[1])
