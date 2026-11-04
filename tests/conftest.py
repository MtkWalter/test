import os
from pathlib import Path

import pytest

# Windows CI can use the native desktop; cloud tests use Qt's real offscreen backend.
if not os.environ.get("DISPLAY") and os.name != "nt":
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture
def samples():
    return Path(__file__).resolve().parents[1] / "samples"


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def window(qapp, monkeypatch):
    from PySide6.QtWidgets import QMessageBox
    from excel_insights.app import MainWindow
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: QMessageBox.StandardButton.Ok)
    monkeypatch.setattr(QMessageBox, "critical", lambda *args: QMessageBox.StandardButton.Ok)
    # Existing interaction tests simulate selecting all sheets in the import dialog.
    monkeypatch.setattr(MainWindow, "select_import", lambda self, catalog: (catalog, self.import_limits))
    widget = MainWindow()
    widget.show()
    qapp.processEvents()
    yield widget
    if widget.busy:
        from PySide6.QtTest import QTest
        import time
        widget.cancel_import()
        deadline = time.monotonic() + 15
        while widget.busy and time.monotonic() < deadline:
            qapp.processEvents(); QTest.qWait(10)
    for worker in (widget._catalog_worker, widget._worker):
        if worker: worker.wait(10000)
    qapp.processEvents()
    widget.close()
    widget.deleteLater()
    qapp.processEvents()
