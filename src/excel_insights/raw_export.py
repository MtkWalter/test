"""Append reports to an unchanged raw workbook package, including all diagnostics.

Original worksheet XML and ancillary parts are copied, not decoded/re-encoded by
openpyxl. An exact copy of the input file is also retained inside the ZIP package.
"""
from __future__ import annotations

import copy
import os
import posixpath
import re
import shutil
import tempfile
import uuid
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill

from .monthly_report import ReportRules, ReportValidationError, generate_report, inspect_workbook

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
STYLE_ORDER = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst")


def _xml(archive, name):
    return archive.read(name).decode("utf-8-sig")


def _children(root, name):
    element = root.find(f"{{{MAIN}}}{name}")
    return list(element) if element is not None else []


def _fragment(element):
    return ET.tostring(element, encoding="unicode")


def _append(xml, name, children, *, count=None, order=()):
    """Keep original XML namespaces, extension data and unknown attributes intact."""
    if not children:
        return xml
    tag = rf"(?:[\w.-]+:)?{re.escape(name)}"
    pattern = rf"<(?P<tag>{tag})(?P<attrs>\s[^<>]*?)?(?:\s*/>|>(?P<body>[\s\S]*?)</(?P=tag)\s*>)"
    match = re.search(pattern, xml)
    addition = "".join(_fragment(c) for c in children)
    if match:
        attrs = match["attrs"] or ""
        # Remove the terminal '/' if it was consumed as an attribute in <tag />.
        attrs = attrs.rstrip().removesuffix("/").rstrip()
        if count is not None:
            attrs = re.sub(r'\s+count\s*=\s*[\'"][^\'"]*[\'"]', "", attrs) + f' count="{count}"'
        replacement = f"<{match['tag']}{attrs}>{match['body'] or ''}{addition}</{match['tag']}>"
        return xml[:match.start()] + replacement + xml[match.end():]
    element = ET.Element(f"{{{MAIN}}}{name}", {} if count is None else {"count": str(count)})
    element.extend(children)
    fragment = _fragment(element)
    if name in order:
        for later in order[order.index(name) + 1:]:
            match = re.search(rf"<(?:[\w.-]+:)?{later}(?:\s|/?>)", xml)
            if match:
                return xml[:match.start()] + fragment + xml[match.start():]
    closing = list(re.finditer(r"</[\w:.-]+\s*>", xml))[-1]
    return xml[:closing.start()] + fragment + xml[closing.start():]


def _merge_styles(original, generated):
    old, new = ET.fromstring(original), ET.fromstring(generated)
    mappings = {}
    old_formats = _children(old, "numFmts")
    used = {int(e.attrib["numFmtId"]) for e in old_formats}
    by_code = {e.attrib["formatCode"]: int(e.attrib["numFmtId"]) for e in old_formats}
    extra_formats = []
    formats = {}
    for element in _children(new, "numFmts"):
        code, previous = element.attrib["formatCode"], int(element.attrib["numFmtId"])
        if code not in by_code:
            next_id = max({163, *used}) + 1
            used.add(next_id)
            by_code[code] = next_id
            element.set("numFmtId", str(next_id))
            extra_formats.append(element)
        formats[previous] = by_code[code]
    original = _append(original, "numFmts", extra_formats,
                       count=len(old_formats) + len(extra_formats), order=STYLE_ORDER)
    for section in ("fonts", "fills", "borders", "cellStyleXfs", "cellXfs"):
        offset = len(_children(old, section))
        elements = _children(new, section)
        mappings[section] = {i: offset + i for i in range(len(elements))}
        if section in {"cellStyleXfs", "cellXfs"}:
            for element in elements:
                for attribute, target in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders"), ("xfId", "cellStyleXfs")):
                    if attribute in element.attrib:
                        element.set(attribute, str(mappings[target][int(element.attrib[attribute])]))
                fmt = int(element.get("numFmtId", "0"))
                if fmt in formats:
                    element.set("numFmtId", str(formats[fmt]))
        original = _append(original, section, elements, count=offset + len(elements), order=STYLE_ORDER)
    return original, mappings["cellXfs"]


def inspect_raw_workbook(source) -> dict:
    """Accept real workbook packages even if their report headers are malformed."""
    source = Path(source)
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("請選擇 .xlsx 或 .xlsm Excel 檔案。")
    try:
        with ZipFile(source) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets = [s.attrib["name"] for s in root.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")]
            if not sheets:
                raise ValueError("Excel 沒有工作表。")
            types = archive.read("[Content_Types].xml")
            macro = b"macroEnabled.main+xml" in types
    except (BadZipFile, KeyError, ET.ParseError, UnicodeError) as exc:
        raise ValueError("無法辨識來源 Excel；檔案損毀、加密或不是有效的 Excel 活頁簿。") from exc
    error = ""
    try:
        report_sheets = [layout.name for layout in inspect_workbook(source)]
    except Exception as exc:
        report_sheets = []
        error = str(exc)
    return {"sheets": report_sheets + [s for s in sheets if s not in report_sheets],
            "report_sheets": report_sheets, "format_error": error,
            "output_extension": ".xlsm" if macro else ".xlsx"}


def _unique_name(name, names):
    candidate, index = name[:31], 1
    while candidate.casefold() in names:
        index += 1
        candidate = name[:31 - len(str(index)) - 3] + f" ({index})"
    names.add(candidate.casefold())
    return candidate


def _put_row(sheet, values, *, header=False):
    cells = []
    for value in values:
        if isinstance(value, str):
            # Diagnostics remain literal text, including leading '=' and controls.
            value = ILLEGAL_CHARACTERS_RE.sub(lambda m: f"\\u{ord(m[0]):04x}", value)
        cell = WriteOnlyCell(sheet, value)
        if isinstance(value, str):
            cell.data_type = "s"
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        if header:
            cell.fill = PatternFill("solid", fgColor="244E5A")
            cell.font = Font(color="FFFFFF", bold=True)
        cells.append(cell)
    sheet.append(cells)
    sheet._raw_row_count = getattr(sheet, "_raw_row_count", 0) + 1


def _append_diagnostics(book, issues):
    headers = ["序號", "工作表", "來源列號", "欄位位置", "欄位名稱", "原始值", "錯誤原因"]
    sheet = None
    for number, issue in enumerate(issues, 1):
        # Excel's cell length limit is handled with continuation rows, never truncation.
        raw, message = str(issue.get("value", "")), str(issue["message"])
        raw_chunks = [raw[i:i + 30000] for i in range(0, len(raw), 30000)] or [""]
        message_chunks = [message[i:i + 30000] for i in range(0, len(message), 30000)] or [""]
        for part in range(max(len(raw_chunks), len(message_chunks))):
            if sheet is None or sheet._raw_row_count >= 1_048_576:
                sheet = book.create_sheet("完整錯誤清單" if sheet is None else f"完整錯誤清單 {len(book.worksheets)}")
                sheet.freeze_panes = "A2"
                for column, width in {"A": 10, "B": 24, "C": 14, "D": 18, "E": 28, "F": 65, "G": 75}.items():
                    sheet.column_dimensions[column].width = width
                _put_row(sheet, headers, header=True)
            _put_row(sheet, [number if part == 0 else f"{number}（續 {part}）", issue.get("sheet", ""),
                             issue.get("row"), issue.get("column", ""), issue.get("field", ""),
                             raw_chunks[part] if part < len(raw_chunks) else "",
                             message_chunks[part] if part < len(message_chunks) else ""])
    if sheet is None:
        sheet = book.create_sheet("完整錯誤清單")
        _put_row(sheet, headers, header=True)
        _put_row(sheet, ["", "", "", "", "", "", "未發現報表驗證錯誤。"])


def _combine(source, supplement, output, catalog, original_part):
    """Copy all source ZIP members, append worksheets, remap only new cell styles."""
    with ZipFile(source) as raw, ZipFile(supplement) as extra, ZipFile(output, "w", ZIP_DEFLATED, allowZip64=True) as target:
        workbook = _xml(raw, "xl/workbook.xml")
        relationships = _xml(raw, "xl/_rels/workbook.xml.rels")
        types = _xml(raw, "[Content_Types].xml")
        source_root = ET.fromstring(workbook)
        used_names = {s.get("name").casefold() for s in _children(source_root, "sheets")}
        existing_ids = {e.get("Id") for e in ET.fromstring(relationships)}
        next_sheet = max(int(s.get("sheetId")) for s in _children(source_root, "sheets")) + 1
        rels, sheets, overrides, additions = [], [], [], {}
        token = uuid.uuid4().hex
        style_rel = next((r for r in ET.fromstring(relationships) if r.get("Type") == DOCREL + "/styles"), None)
        if style_rel is not None:
            style_name = style_rel.get("Target")
            style_name = style_name.lstrip("/") if style_name.startswith("/") else posixpath.normpath("xl/" + style_name)
            styles, mapping = _merge_styles(_xml(raw, style_name), _xml(extra, "xl/styles.xml"))
        else:
            style_name = f"xl/raw-report-{token}/styles.xml"
            styles = _xml(extra, "xl/styles.xml")
            mapping = {i: i for i, _ in enumerate(_children(ET.fromstring(styles), "cellXfs"))}
            rels.append(ET.Element(f"{{{REL}}}Relationship", {"Id": "rawStyles" + token, "Type": DOCREL + "/styles", "Target": "/" + style_name}))
            overrides.append(ET.Element(f"{{{CT}}}Override", {"PartName": "/" + style_name, "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"}))
        extra_root = ET.fromstring(extra.read("xl/workbook.xml"))
        extra_rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(extra.read("xl/_rels/workbook.xml.rels"))}
        names = {}
        for index, source_sheet in enumerate(_children(extra_root, "sheets")):
            name = _unique_name(source_sheet.get("name"), used_names)
            names[source_sheet.get("name")] = name
            relationship_id = "rawReport" + token + str(index)
            assert relationship_id not in existing_ids
            part_name = f"xl/raw-report-{token}/sheet{index + 1}.xml"
            original_target = extra_rels[source_sheet.get(f"{{{DOCREL}}}id")]
            original_target = original_target.lstrip("/") if original_target.startswith("/") else posixpath.normpath("xl/" + original_target)
            xml = _xml(extra, original_target)
            # Only change style attributes inside element tags, never inline text.
            def remap(match):
                tag = match[0]
                attribute = "style" if re.match(r"<(?:\w+:)?col\b", tag) else "s"
                return re.sub(rf'\b{attribute}="(\d+)"', lambda value: f'{attribute}="{mapping[int(value[1])]}"', tag)
            additions[part_name] = re.sub(r"<(?:\w+:)?(?:c|row|col)\b[^>]*>", remap, xml)
            sheets.append(ET.Element(f"{{{MAIN}}}sheet", {"name": name, "sheetId": str(next_sheet + index), f"{{{DOCREL}}}id": relationship_id}))
            rels.append(ET.Element(f"{{{REL}}}Relationship", {"Id": relationship_id, "Type": DOCREL + "/worksheet", "Target": "/" + part_name}))
            overrides.append(ET.Element(f"{{{CT}}}Override", {"PartName": "/" + part_name, "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}))
        workbook = _append(workbook, "sheets", sheets)
        # Workbook-scoped print settings must point to appended sheets, including collisions.
        defined_names = _children(extra_root, "definedNames")
        for element in defined_names:
            if "localSheetId" in element.attrib:
                element.set("localSheetId", str(int(element.get("localSheetId")) + len(catalog["sheets"])))
            if element.text:
                for old_name, new_name in names.items():
                    element.text = element.text.replace("'" + old_name.replace("'", "''") + "'!", "'" + new_name.replace("'", "''") + "'!")
        workbook = _append(workbook, "definedNames", defined_names,
                           order=("sheets", "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"))
        relationships = _append(relationships, "Relationships", rels + [ET.Element(f"{{{REL}}}Relationship", {
            "Id": "rawOriginal" + token, "Type": "urn:monthly-report:original-input", "Target": "/" + original_part})])
        types = _append(types, "Types", overrides + [ET.Element(f"{{{CT}}}Override", {"PartName": "/" + original_part,
            "ContentType": "application/vnd.ms-excel.sheet.macroEnabled.12" if catalog["output_extension"] == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})])
        changes = {"xl/workbook.xml": workbook, "xl/_rels/workbook.xml.rels": relationships,
                   "[Content_Types].xml": types, style_name: styles}
        target.comment = raw.comment
        for info in raw.infolist():
            if info.filename in changes:
                target.writestr(copy.copy(info), changes.pop(info.filename).encode("utf-8"))
            else:
                with raw.open(info) as incoming, target.open(copy.copy(info), "w") as outgoing:
                    shutil.copyfileobj(incoming, outgoing, length=1024 * 1024)
        for name, xml in {**changes, **additions}.items():
            target.writestr(name, xml.encode("utf-8"))
        target.write(source, original_part)
        return names


def generate_report_with_raw(source, output, year, month, *, sheet=None, fiscal_year=None,
                             rules=None, settings_error=None, overwrite=False, progress=None):
    source = Path(source).resolve()
    catalog = inspect_raw_workbook(source)
    output = Path(output).resolve().with_suffix(catalog["output_extension"])
    if source == output or output.exists() and os.path.samefile(source, output):
        raise ValueError("輸出不可覆寫輸入檔案，請選擇另一個檔名。")
    if output.exists() and not overwrite:
        raise FileExistsError("輸出檔案已存在，請選擇新檔名。")
    output.parent.mkdir(parents=True, exist_ok=True)
    result = {"year": year, "month": month, "sheet": sheet, "fiscal_year": fiscal_year,
              "act_rows": None, "ori_rows": None, "remark_nodes": None, "warnings": []}
    issues = []
    with tempfile.TemporaryDirectory(prefix=".raw-export-", dir=output.parent) as folder:
        supplement = Path(folder) / "supplement.xlsx"
        try:
            if settings_error:
                raise ValueError(settings_error)
            if not isinstance(year, int) or isinstance(year, bool) or not isinstance(month, int) or isinstance(month, bool):
                raise ValueError("年份與月份必須是整數。")
            parsed_rules = rules if isinstance(rules, ReportRules) else ReportRules.from_dict(rules or {})
            result = generate_report(source, supplement, year, month, sheet=sheet,
                                     fiscal_year=fiscal_year, rules=parsed_rules, progress=progress)
            from openpyxl import load_workbook
            book = load_workbook(supplement)
            generated = True
        except Exception as exc:
            issues = exc.issues if isinstance(exc, ReportValidationError) else [
                {"sheet": sheet or "", "row": None, "column": "", "field": "來源／設定", "value": "", "message": str(exc)}]
            book = Workbook(write_only=True)
            generated = False
            result.update(act_rows=None, ori_rows=None, remark_nodes=None)
        original_part = f"raw-input-{uuid.uuid4().hex}/original{source.suffix.lower()}"
        try:
            note = book.create_sheet("匯出說明")
            for row in [
                ["匯出模式", "完整保留 Raw Data"],
                ["統計狀態", "已產生單月報表" if generated else "未產生統計：來源或設定有錯誤，請查看完整錯誤清單。"],
                ["原始檔名", source.name], ["原始工作表數", len(catalog["sheets"])],
                ["保留範圍", "所有原始工作表、所有期間與原有儲存格、公式、樣式及附屬檔案；原始工作表不改名、不篩列、不改值。"],
                ["錯誤數", len(issues)], ["驗證範圍", "所選月份及 FY 的 act／ori 報表資料；原始資料不受此篩選限制。"],
                ["原檔備份", original_part],
                ["備份取回方式", "用 ZIP 工具開啟輸出 Excel，取出上述路徑即可取得與輸入完全相同的原檔。"],
            ]:
                _put_row(note, row)
            note.column_dimensions["A"].width = 22
            note.column_dimensions["B"].width = 110
            _append_diagnostics(book, issues)
            book.save(supplement)
        finally:
            book.close()
        if progress:
            progress(f"正在保留全部 Raw Data 與 {len(issues):,} 項錯誤資訊…")
        temporary = Path(folder) / ("output" + catalog["output_extension"])
        names = _combine(source, supplement, temporary, catalog, original_part)
        if output.exists() and not overwrite:
            raise FileExistsError("輸出檔案已存在，請選擇新檔名。")
        os.replace(temporary, output)
    result.update({"output": str(output), "include_raw": True, "report_generated": generated,
                   "raw_sheet_count": len(catalog["sheets"]), "raw_sheet_names": catalog["sheets"],
                   "info_sheet": names["匯出說明"], "preview_sheet": names.get("monthly pivot"),
                   "issues": issues, "output_extension": catalog["output_extension"]})
    result["warnings"] = [*result["warnings"], f"已完整保留 {len(catalog['sheets'])} 張原始工作表；完整錯誤清單另附在 Excel。"]
    return result
