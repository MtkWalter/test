from copy import deepcopy

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QFileDialog
from openpyxl import load_workbook

from excel_insights.reporting import demo_sheet, save_config, suggest_config
from excel_insights.report_ui import ReportDialog


def test_preview_edit_reconfirm_and_export(qapp, tmp_path, monkeypatch):
    source = demo_sheet(); widget = ReportDialog([source]); widget.show()
    widget.tabs.setCurrentIndex(2); qapp.processEvents()
    assert widget.result.frame.iloc[0]["Total act"] == 191_000_000
    assert not widget.export_button.isEnabled()
    widget.confirm.setChecked(True)
    assert widget.export_button.isEnabled()
    widget.multiplier.setValue(.001)
    assert not widget.confirm.isChecked()
    assert not widget.export_button.isEnabled()
    QTest.qWait(350); qapp.processEvents()
    assert widget.result.frame.iloc[0]["Total act"] == 191000
    widget.confirm.setChecked(True)
    output = tmp_path / "report.xlsx"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: (str(output), "Excel"))
    widget.export_button.click()
    book = load_workbook(output)
    assert book.active["K12"].value == 191000
    book.close(); widget.close()


def test_sources_are_combined_only_when_checked_and_settings_restore(qapp, tmp_path, monkeypatch):
    sources = [demo_sheet(), deepcopy(demo_sheet())]
    widget = ReportDialog(sources); widget.show(); qapp.processEvents()
    assert widget.result.frame.iloc[0]["Total act"] == 382_000_000
    widget.source_list.item(1).setCheckState(Qt.CheckState.Unchecked)
    QTest.qWait(350); qapp.processEvents()
    assert widget.result.frame.iloc[0]["Total act"] == 191_000_000
    config = suggest_config(sources); config.display_unit = 1000; config.total_mode = "classified"
    path = tmp_path / "preset.json"; save_config(path, config)
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(path), "JSON"))
    widget.open_settings()
    assert widget.unit.currentData() == 1000
    assert widget.result.frame.iloc[0]["Total act"] == 362_000_000
    widget.close()


def test_category_values_can_be_changed_without_writing_rules(qapp):
    widget = ReportDialog([demo_sheet()]); widget.show(); qapp.processEvents()
    field, _, values = widget.category_controls["FI"]
    field.setCurrentIndex(field.findData("product"))
    for i in range(values.count()):
        values.item(i).setSelected(values.item(i).text() == "Service")
    QTest.qWait(350); qapp.processEvents()
    assert widget.result.frame.iloc[0]["FI act"] == 10_000_000
    widget.close()


def test_manual_remarks_edit_preserve_reconfirm_export_and_reset(qapp, tmp_path, monkeypatch):
    import json
    widget = ReportDialog([demo_sheet()]); widget.show(); widget.tabs.setCurrentIndex(2)
    qapp.processEvents()
    model = widget.preview.model()
    assert model.flags(model.index(0, 14)) & Qt.ItemFlag.ItemIsEditable
    assert not model.flags(model.index(0, 1)) & Qt.ItemFlag.ItemIsEditable
    widget.confirm.setChecked(True)
    text = "=這是手動備註\n下季追蹤"
    assert model.setData(model.index(0, 14), text)
    assert not widget.confirm.isChecked()
    assert not widget.export_button.isEnabled()
    widget.multiplier.setValue(.001)
    QTest.qWait(350); qapp.processEvents()
    assert widget.result.frame.iloc[0]["Remark"] == text
    preset = tmp_path / "settings.json"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: (str(preset), "JSON"))
    widget.save_settings()
    assert text not in json.dumps(json.loads(preset.read_text()), ensure_ascii=False)
    remarks = tmp_path / "remarks.json"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: (str(remarks), "JSON"))
    widget.save_remarks()
    reopened = ReportDialog([demo_sheet()], widget.config)
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (str(remarks), "JSON"))
    reopened.open_remarks()
    assert reopened.result.frame.iloc[0]["Remark"] == text
    assert not reopened.confirm.isChecked()
    reopened.close()
    path = tmp_path / "with-remark.xlsx"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: (str(path), "Excel"))
    widget.confirm.setChecked(True); widget.export_button.click()
    book = load_workbook(path)
    assert book.active["R12"].value == text
    assert book.active["R12"].data_type == "s"
    book.close()
    widget.preview.setCurrentIndex(widget.preview.model().index(0, 0))
    widget.reset_remark()
    assert "觸發：TSMC 120.0%" in widget.result.frame.iloc[0]["Remark"]
    assert "示範資料，非實際業績" in widget.result.frame.iloc[0]["Remark"]
    assert not widget.confirm.isChecked()
    widget.close()
