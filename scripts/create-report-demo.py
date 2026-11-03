"""Generate synthetic input and an example report; no user data is embedded."""
from pathlib import Path

from openpyxl import Workbook

from excel_insights.reporting import build_report, demo_sheet, export_report, save_config, suggest_config

root = Path(__file__).resolve().parents[1] / "samples"
root.mkdir(exist_ok=True)
source = demo_sheet()
book = Workbook(); sheet = book.active; sheet.title = "Demo Input"
sheet.append(["示範資料，非實際業績。所有金額與分類均為虛構。"])
sheet.append(["用來測試第 4 列標題自動判讀與本機預覽。"])
sheet.append([None])
sheet.append(source.data.columns.tolist())
for row in source.data.itertuples(index=False, name=None): sheet.append(row)
book.save(root / "report-demo-input.xlsx")
source.path = root / "report-demo-input.xlsx"
source.name = "Demo Input"
config = suggest_config([source])
config.year_mode = "fiscal"; config.year = 2027; config.fiscal_start = 4; config.month = 9
export_report(root / "report-demo-output.xlsx", build_report([source], config), config, [source])
save_config(root / "report-demo-settings.json", config)
print("Created report-demo-input.xlsx, report-demo-output.xlsx and report-demo-settings.json")
