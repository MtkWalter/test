"""Create small, real .xlsx demo workbooks using only Python's standard library."""
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1] / 'samples'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def workbook(path, sheets):
    def column(index):
        name = ''
        while index:
            index, remainder = divmod(index - 1, 26)
            name = chr(65 + remainder) + name
        return name

    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        overrides = ''.join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, len(sheets) + 1))
        z.writestr('[Content_Types].xml',
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   + overrides + '</Types>')
        z.writestr('_rels/.rels', f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        entries = ''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(sheets, 1))
        z.writestr('xl/workbook.xml', f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{entries}</sheets></workbook>')
        links = ''.join(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
        z.writestr('xl/_rels/workbook.xml.rels', f'<Relationships xmlns="{PKG}">{links}</Relationships>')
        for i, rows in enumerate(sheets.values(), 1):
            xml_rows = []
            for r, values in enumerate(rows, 1):
                cells = []
                for c, value in enumerate(values, 1):
                    address = f'{column(c)}{r}'
                    if value is None:
                        continue
                    if isinstance(value, (int, float)):
                        cells.append(f'<c r="{address}"><v>{value}</v></c>')
                    else:
                        cells.append(f'<c r="{address}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
                xml_rows.append(f'<row r="{r}">{"".join(cells)}</row>')
            z.writestr(f'xl/worksheets/sheet{i}.xml', f'<worksheet xmlns="{MAIN}"><sheetData>{"".join(xml_rows)}</sheetData></worksheet>')


if __name__ == '__main__':
    ROOT.mkdir(exist_ok=True)
    header = ['日期', '地區', '商品', '數量', '金額']
    workbook(ROOT / '銷售範例-一月.xlsx', {
        '銷售明細': [header,
                 ['2026-01-02', '台北', '鍵盤', 2, 2400],
                 ['2026-01-03', '台中', '滑鼠', 3, 1500],
                 ['2026-01-04', '台北', '螢幕', 1, 6000],
                 ['2026-01-05', '高雄', '鍵盤', 1, 1200],
                 ['2026-01-06', '台北', '滑鼠', 1, None],
                 ['2026-01-07', '台中', '螢幕', 1, '待確認']],
        '庫存': [['商品', '庫存量'], ['鍵盤', 20], ['滑鼠', 35], ['螢幕', 12]]})
    workbook(ROOT / '銷售範例-二月.xlsx', {
        '銷售明細': [header,
                 ['2026-02-02', '台北', '鍵盤', 3, 3600],
                 ['2026-02-03', '高雄', '螢幕', 2, 12000],
                 ['2026-02-04', '台中', '滑鼠', 4, 2000]]})
    print(f'Created demo workbooks in {ROOT}')
