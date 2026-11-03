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
    widget = MainWindow()
    widget.show()
    qapp.processEvents()
    yield widget
    if widget._worker:
        widget._worker.wait(10000)
        qapp.processEvents()
    widget.close()
    widget.deleteLater()
    qapp.processEvents()
