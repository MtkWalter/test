from pathlib import Path

import pytest
from openpyxl import Workbook
from PySide6.QtCore import Qt

from excel_insights import resources
from excel_insights.core import (
    ImportCancelled, ImportLimitError, ImportLimits, analyze, read_workbook, workbook_sheet_names,
)
from excel_insights.import_ui import ImportSelectionDialog
from test_app import wait_for_import


def create_mixed(path):
    book = Workbook(); raw = book.active; raw.title = "Raw Data"
    raw.append(["Tech Node", "Amount"]); raw.append(["3nm", 120])
    other = book.create_sheet("Other Data")
    other.append(["Tech Node", "Amount"]); other.append(["5nm", 80])
    pivot = book.create_sheet("monthly pivot(CS)")
    pivot.append([f"column-{i}" for i in range(100)])
    pivot.append(list(range(100)))
    book.save(path); book.close()
    return path


def test_selected_sheet_does_not_read_or_charge_unneeded_pivot(tmp_path):
    path = create_mixed(tmp_path / "mixed.xlsx")
    assert workbook_sheet_names(path) == ["Raw Data", "Other Data", "monthly pivot(CS)"]
    limits = ImportLimits(max_cells=10)
    result = read_workbook(path, selected_sheets=["Raw Data"], limits=limits)
    assert len(result) == 1
    assert result[0].data["Amount"].tolist() == [120]
    with pytest.raises(ImportLimitError, match="提高匯入"):
        read_workbook(path, limits=limits)
    with pytest.raises(ValueError, match="找不到"):
        read_workbook(path, selected_sheets=["renamed"])


def test_budget_can_be_increased_without_truncating(tmp_path):
    path = create_mixed(tmp_path / "wide.xlsx")
    with pytest.raises(ImportLimitError):
        read_workbook(path, selected_sheets=["monthly pivot(CS)"], limits=ImportLimits(max_cells=100))
    result = read_workbook(path, selected_sheets=["monthly pivot(CS)"], limits=ImportLimits(max_cells=500))
    assert result[0].data.shape == (1, 100)
    assert result[0].data.iat[0, 99] == 99


def test_catalog_does_not_decode_shared_strings_or_sheet_rows(tmp_path, monkeypatch):
    from openpyxl.reader.excel import ExcelReader
    path = create_mixed(tmp_path / "catalog-only.xlsx")
    def unexpected(*args, **kwargs):
        pytest.fail("Worksheet selection must not load worksheet data or shared strings")
    monkeypatch.setattr(ExcelReader, "read_strings", unexpected)
    monkeypatch.setattr(ExcelReader, "read_worksheets", unexpected)
    assert workbook_sheet_names(path) == ["Raw Data", "Other Data", "monthly pivot(CS)"]


def test_data_exceeding_both_old_limits_is_fully_imported(tmp_path):
    path = tmp_path / "large-real-data.xlsx"
    book = Workbook(write_only=True); sheet = book.create_sheet("Raw Data")
    sheet.append(["Tech Node", "Amount", "scenario", "Category", "CUSTOMER", "P/N", "Year", "Month", "extra1", "extra2"])
    count = 100_001
    for i in range(count):
        sheet.append(["3nm", 1, "act", "TSMC", "Customer A", "P1", 2026, 9, "x", "y"])
    book.save(path)
    result = read_workbook(path)[0]
    assert result.data.shape == (count, 10)
    assert result.cells > 1_000_000
    assert analyze(result.data, "Amount").iloc[0]["總和"] == count


def test_import_can_be_cancelled_and_memory_shortage_is_explained(tmp_path, monkeypatch):
    path = create_mixed(tmp_path / "cancel.xlsx")
    with pytest.raises(ImportCancelled):
        read_workbook(path, cancelled=lambda: True)
    monkeypatch.setattr(resources, "available_memory", lambda: 100)
    with pytest.raises(MemoryError, match="記憶體"):
        read_workbook(path)


def test_sheet_selection_and_editable_limits(qapp, tmp_path):
    path = Path(tmp_path / "book.xlsx")
    widget = ImportSelectionDialog({path: ["Raw Data", "monthly pivot(CS)"]}, ImportLimits())
    widget.show(); qapp.processEvents()
    assert widget.selection() == {path: ["Raw Data"]}
    assert widget.limits().max_cells == 20_000_000
    widget.file_millions.setValue(50); widget.total_millions.setValue(80)
    assert widget.limits().max_cells == 50_000_000
    assert widget.limits().total_cells == 80_000_000
    widget.choose("all")
    assert len(widget.selection()[path]) == 2
    widget.choose("none")
    assert not widget.ok_button.isEnabled()
    widget.close()


def test_importing_another_sheet_from_same_file_does_not_duplicate_existing(window, qapp, tmp_path, monkeypatch):
    path = create_mixed(tmp_path / "additional.xlsx").resolve()
    monkeypatch.setattr(window, "select_import", lambda catalog: ({path: ["Raw Data"]}, ImportLimits(max_cells=10)))
    window.import_paths([path]); wait_for_import(window, qapp)
    assert [sheet.name for sheet in window.sheets] == ["Raw Data"]
    monkeypatch.setattr(window, "select_import", lambda catalog: ({path: ["Raw Data", "Other Data"]}, ImportLimits(max_cells=10)))
    window.import_paths([path]); wait_for_import(window, qapp)
    assert [sheet.name for sheet in window.sheets] == ["Raw Data", "Other Data"]


def test_catalog_cancel_keeps_existing_data(window, qapp, tmp_path, monkeypatch):
    path = create_mixed(tmp_path / "cancel-dialog.xlsx").resolve()
    monkeypatch.setattr(window, "select_import", lambda catalog: None)
    window.import_paths([path]); wait_for_import(window, qapp)
    assert not window.sheets
    assert "取消" in window.status.text()


def test_remaining_total_budget_does_not_discard_previously_loaded_sheets(window, qapp, tmp_path, monkeypatch):
    path = create_mixed(tmp_path / "total-limit.xlsx").resolve()
    monkeypatch.setattr(window, "select_import", lambda catalog: ({path: ["Raw Data"]}, ImportLimits(max_cells=20, total_cells=3)))
    window.import_paths([path]); wait_for_import(window, qapp)
    # File budget includes headers, so 3 cells cannot fit a 2-column header + data.
    assert not window.sheets
    monkeypatch.setattr(window, "select_import", lambda catalog: ({path: ["Raw Data"]}, ImportLimits(max_cells=20, total_cells=5)))
    window.import_paths([path]); wait_for_import(window, qapp)
    assert len(window.sheets) == 1
    monkeypatch.setattr(window, "select_import", lambda catalog: ({path: ["Other Data"]}, ImportLimits(max_cells=20, total_cells=5)))
    window.import_paths([path]); wait_for_import(window, qapp)
    assert [sheet.name for sheet in window.sheets] == ["Raw Data"]
