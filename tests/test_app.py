import csv
import time

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest

from excel_insights.app import FrameModel
from excel_insights.core import read_workbook


def wait_for_import(window, qapp):
    deadline = time.monotonic() + 15
    while window.busy and time.monotonic() < deadline:
        qapp.processEvents()
        QTest.qWait(10)
    assert not window.busy, "Import worker did not finish"
    qapp.processEvents()


def test_desktop_import_filter_group_export_remove(window, qapp, samples, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QFileDialog
    window.import_paths(list(samples.glob("*.xlsx")))
    wait_for_import(window, qapp)
    assert len(window.sheets) == 3
    index = next(i for i, sheet in enumerate(window.sheets) if sheet.path.name.endswith("一月.xlsx") and sheet.name == "銷售明細")
    window.sheet_list.setCurrentRow(index)
    window.numeric_column.setCurrentText("金額")
    window.group_column.setCurrentIndex(window.group_column.findData("地區"))
    window.filter_column.setCurrentText("地區")
    window.filter_text.setText("台北")
    # Exercise the same debounced signal path as actual typing.
    QTest.qWait(350)
    qapp.processEvents()
    assert window.preview.model().rowCount() == 3
    assert window.summary.iloc[0]["總和"] == 8400
    assert window.summary.iloc[0]["平均"] == 4200
    assert window.export_button.isEnabled()
    path = tmp_path / "ui-export.csv"
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: (str(path), "CSV"))
    window.export_button.click()
    with path.open(encoding="utf-8-sig", newline="") as stream:
        result = list(csv.reader(stream))
    assert result[1][1] == "銷售明細"
    assert result[1][-4] == "8400.0"
    window.remove_button.click()
    assert len(window.sheets) == 2
    window.clear_button.click()
    assert window.sheet_list.count() == 0
    assert window.preview_stack.currentIndex() == 0
    assert not window.export_button.isEnabled()


def test_bad_file_does_not_block_valid_file_or_duplicate_import(window, qapp, samples, tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"broken")
    good = samples / "銷售範例-一月.xlsx"
    window.import_paths([bad, good])
    wait_for_import(window, qapp)
    assert len(window.sheets) == 2
    window.import_paths([good, good])
    wait_for_import(window, qapp)
    assert len(window.sheets) == 2


def test_preview_caps_rows_but_statistics_use_all(window, qapp, samples):
    sheet = read_workbook(samples / "銷售範例-二月.xlsx")[0]
    import pandas as pd
    sheet.data = pd.DataFrame({"value": range(6000)})
    window._import_result([sheet], [])
    assert window.preview.model().rowCount() == 5000
    assert window.summary.iloc[0]["資料筆數"] == 6000
    assert window.summary.iloc[0]["總和"] == sum(range(6000))


def test_numeric_and_special_headers_render_in_model(qapp):
    import pandas as pd
    frame = pd.DataFrame({"a.[b]/c": [12.5, None]}, dtype=object)
    model = FrameModel(frame)
    assert model.headerData(0, Qt.Orientation.Horizontal) == "a.[b]/c"
    assert model.data(model.index(0, 0)) == "12.5"
    assert model.data(model.index(1, 0)) == ""
