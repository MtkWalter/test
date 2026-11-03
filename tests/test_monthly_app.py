import time

from openpyxl import load_workbook
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest

from excel_insights.monthly_app import ClassificationsDialog, ReportWindow
from excel_insights.monthly_report import ReportRules
from excel_insights.report_sample import create_sample


def wait_for_report(window, qapp):
    deadline = time.monotonic() + 15
    while window.busy and time.monotonic() < deadline:
        qapp.processEvents()
        QTest.qWait(10)
    qapp.processEvents()
    assert not window.busy, "Report worker did not finish"


def test_report_gui_generates_workbook_via_button(qapp, tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    output = tmp_path / "output.xlsx"
    window = ReportWindow()
    try:
        window.show()
        window.year.setValue(2026)
        window.month.setCurrentIndex(8)
        window.load_input(source)
        window.output_path.setText(str(output))
        assert window.fiscal_year.text() == "FY27"
        QTest.mouseClick(window.generate_button, Qt.MouseButton.LeftButton)
        wait_for_report(window, qapp)
        assert window.last_result["act_rows"] == 20
        assert window.generate_button.isEnabled() and window.open_button.isEnabled()
        book = load_workbook(output, data_only=True)
        row = next(row for row in book.active if row[0].value == "N20")
        assert row[1].value == 120 and row[3].value == 80
        assert "END_A +40.0M" in row[11].value
        assert "Service" in window.messages.toPlainText()
        book.close()
    finally:
        if window._worker:
            window._worker.wait(15000)
            qapp.processEvents()
        window.close()


def test_report_gui_missing_actual_is_actionable_and_does_not_write(qapp, tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    book.active["AS3"] = None
    book.save(source)
    book.close()
    window = ReportWindow()
    try:
        window.year.setValue(2026)
        window.month.setCurrentIndex(8)
        window.load_input(source)
        output = tmp_path / "output.xlsx"
        window.output_path.setText(str(output))
        window.generate_button.click()
        wait_for_report(window, qapp)
        assert window.last_result is None
        assert "AS 實績金額空白" in window.messages.toPlainText()
        assert window.generate_button.isEnabled()
        assert not window.open_button.isEnabled()
        assert not output.exists()
    finally:
        if window._worker:
            window._worker.wait(15000)
            qapp.processEvents()
        window.close()


def test_classification_editor_preserves_rules(qapp):
    rules = ReportRules(non_t_values=["FI", "non-t"], node_aliases={"N2": "N2/A16", "A16": "N2/A16"})
    dialog = ClassificationsDialog(rules)
    assert dialog.rules().to_dict() == rules.to_dict()
    dialog.non_t.clear()
    assert dialog.rules().non_t_values is None
    dialog.close()
