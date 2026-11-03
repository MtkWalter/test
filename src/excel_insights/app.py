"""PySide6 desktop interface. Workbook I/O runs in a worker thread."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QThread, QTimer, Signal
from PySide6.QtGui import QCloseEvent, QDragEnterEvent, QDropEvent, QFont
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QProgressBar, QPushButton, QSplitter, QStackedWidget,
    QTableView, QVBoxLayout, QWidget, QSpinBox,
)

from .core import (
    SUMMARY_COLUMNS, TOTAL_CELL_LIMIT, Sheet, analyze, display, export_csv,
    filter_rows, is_blank, number, read_workbook,
)

STYLE = """
QMainWindow, QWidget#workspace { background: #f2f5fa; color: #24324a; }
QWidget { font-size: 13px; }
QFrame#sidebar { background: #14233d; }
QFrame#sidebar QLabel { color: #b5c3dc; }
QLabel#brand { color: #96baff; font-size: 16px; font-weight: 700; }
QLabel#sideTitle { color: white; font-size: 23px; font-weight: 600; }
QLabel#heading { font-size: 25px; font-weight: 600; }
QLabel#subheading, QLabel#status { color: #63718a; }
QLabel#summaryHeading { font-size: 17px; font-weight: 600; }
QPushButton { background: #e8eefa; color: #234891; border: 0; border-radius: 5px; padding: 10px 14px; }
QPushButton:hover { background: #d6e3ff; }
QPushButton:disabled { color: #8e9aaf; background: #e7eaf0; }
QPushButton#import { background: #477df4; color: white; font-weight: 600; }
QPushButton#import:hover { background: #366ce3; }
QListWidget { background: #1d304e; color: white; border: 0; border-radius: 5px; }
QListWidget::item { padding: 10px 5px; }
QListWidget::item:selected { background: #35568b; }
QFrame#controls { background: white; border-radius: 8px; }
QComboBox, QLineEdit { background: white; color: #24324a; border: 1px solid #d7dfeb; border-radius: 4px; min-height: 22px; padding: 5px 8px; }
QComboBox QAbstractItemView { background: white; color: #24324a; selection-background-color: #d6e3ff; }
QTableView { background: white; alternate-background-color: #f7f9fc; color: #24324a; border: 1px solid #e0e6f0; gridline-color: #e9edf3; selection-background-color: #d6e3ff; selection-color: #24324a; }
QHeaderView::section { background: #eaf0f8; color: #344963; border: 0; border-bottom: 1px solid #d7dfeb; padding: 8px; }
QLabel#empty { background: white; color: #7c8aa0; border: 1px dashed #c9d3e2; border-radius: 8px; font-size: 16px; }
QProgressBar { border: 0; background: #e0e8f5; max-height: 4px; }
QProgressBar::chunk { background: #598bff; }
"""


class FrameModel(QAbstractTableModel):
    def __init__(self, frame: pd.DataFrame, *, summary: bool = False, parent=None):
        super().__init__(parent)
        self.frame = frame
        self.summary = summary

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.frame)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.frame.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        value = self.frame.iat[index.row(), index.column()]
        if role == Qt.ItemDataRole.DisplayRole:
            if self.summary and index.column() >= 5 and not is_blank(value):
                return f"{value:,.2f}"
            return display(value)
        if role == Qt.ItemDataRole.ToolTipRole:
            return display(value)
        if role == Qt.ItemDataRole.TextAlignmentRole and number(value) is not None:
            return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return str(self.frame.columns[section]) if orientation == Qt.Orientation.Horizontal else str(section + 1)
        return None


class ImportWorker(QThread):
    progress = Signal(str)
    result = Signal(object, object)

    def __init__(self, paths: list[Path], existing: list[Sheet], parent=None, header_row=None):
        super().__init__(parent)
        self.paths = paths
        self.existing = existing.copy()
        self.header_row = header_row

    def run(self):
        loaded: list[Sheet] = []
        errors: list[str] = []
        known = {str(sheet.path).casefold() for sheet in self.existing}
        cells = sum(sheet.cells for sheet in self.existing)
        for path in self.paths:
            if str(path).casefold() in known:
                continue
            self.progress.emit(f"正在匯入 {path.name}…")
            try:
                sheets = read_workbook(path, self.header_row)
                incoming = sum(sheet.cells for sheet in sheets)
                if cells + incoming > TOTAL_CELL_LIMIT:
                    raise ValueError("已載入資料超過 2,000,000 儲存格限制；請先移除部分工作表。")
                if not sheets:
                    raise ValueError("沒有可讀取的工作表。")
                loaded.extend(sheets)
                cells += incoming
                known.add(str(path).casefold())
            except Exception as exc:
                errors.append(f"{path.name}：{exc}")
        self.result.emit(loaded, errors)


def label(text: str, name: str = "") -> QLabel:
    widget = QLabel(text)
    if name:
        widget.setObjectName(name)
    widget.setWordWrap(True)
    return widget


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.sheets: list[Sheet] = []
        self.summary = pd.DataFrame(columns=SUMMARY_COLUMNS)
        self.busy = False
        self._updating = False
        self._worker: ImportWorker | None = None
        self.report_config = None
        self.setWindowTitle("Excel Insights · Python 資料統計工作台")
        self.resize(1240, 850)
        self.setMinimumSize(980, 700)
        self.setAcceptDrops(True)
        self.setStyleSheet(STYLE)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(250)
        self._timer.timeout.connect(self.refresh)
        self._build_ui()

    def _build_ui(self):
        root = QWidget()
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(root)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(270)
        left = QVBoxLayout(sidebar)
        left.setContentsMargins(20, 28, 20, 24)
        left.setSpacing(14)
        left.addWidget(label("▦  EXCEL INSIGHTS", "brand"))
        left.addWidget(label("資料統計工作台", "sideTitle"))
        self.import_button = QPushButton("＋ 匯入 Excel 檔案")
        self.import_button.setObjectName("import")
        self.import_button.clicked.connect(self.choose_files)
        left.addWidget(self.import_button)
        left.addWidget(label("可一次選取多檔，或拖曳檔案至視窗"))
        left.addWidget(label("標題列（修改後重新匯入）"))
        self.header_row = QSpinBox()
        self.header_row.setRange(0, 50)
        self.header_row.setSpecialValueText("自動判讀標題列")
        left.addWidget(self.header_row)
        self.report_button = QPushButton("本機判讀＋報表預覽")
        self.report_button.clicked.connect(self.open_report)
        left.addWidget(self.report_button)
        left.addWidget(label("檔案 / 工作表"))
        self.sheet_list = QListWidget()
        self.sheet_list.currentRowChanged.connect(self.select_sheet)
        left.addWidget(self.sheet_list, 1)
        self.remove_button = QPushButton("移除選取工作表")
        self.remove_button.clicked.connect(self.remove_sheet)
        self.clear_button = QPushButton("清除全部")
        self.clear_button.clicked.connect(self.clear_sheets)
        left.addWidget(self.remove_button)
        left.addWidget(self.clear_button)
        left.addWidget(label("資料在本機處理，不上傳雲端。\n首列視為欄名；不修改來源檔。"))
        layout.addWidget(sidebar)

        workspace = QWidget()
        workspace.setObjectName("workspace")
        right = QVBoxLayout(workspace)
        right.setContentsMargins(26, 24, 26, 22)
        right.setSpacing(12)
        right.addWidget(label("從試算表，到清楚的答案。", "heading"))
        right.addWidget(label("選擇工作表與數值欄位，依需要篩選或分組。統計僅計算目前選取的工作表。", "subheading"))

        controls = QFrame()
        controls.setObjectName("controls")
        grid = QGridLayout(controls)
        grid.setContentsMargins(16, 14, 16, 14)
        self.numeric_column = QComboBox()
        self.group_column = QComboBox()
        self.export_button = QPushButton("匯出統計 CSV")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self.export)
        grid.addWidget(label("數值欄位"), 0, 0)
        grid.addWidget(label("分組欄位（可選）"), 0, 1)
        grid.addWidget(self.numeric_column, 1, 0)
        grid.addWidget(self.group_column, 1, 1)
        grid.addWidget(self.export_button, 1, 2)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        right.addWidget(controls)

        filter_bar = QHBoxLayout()
        self.filter_column = QComboBox()
        self.filter_column.setMinimumWidth(140)
        self.filter_column.setToolTip("篩選欄位")
        self.filter_text = QLineEdit()
        self.filter_text.setPlaceholderText("輸入包含文字，不區分大小寫")
        reset = QPushButton("清除篩選")
        reset.clicked.connect(self.filter_text.clear)
        filter_bar.addWidget(self.filter_column)
        filter_bar.addWidget(self.filter_text, 1)
        filter_bar.addWidget(reset)
        right.addLayout(filter_bar)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self.preview_stack = QStackedWidget()
        empty = label("將 Excel 拖曳至此，或按左側「匯入 Excel 檔案」開始。", "empty")
        empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_stack.addWidget(empty)
        self.preview = self._table()
        self.preview_stack.addWidget(self.preview)
        splitter.addWidget(self.preview_stack)
        summary_panel = QWidget()
        summary_layout = QVBoxLayout(summary_panel)
        summary_layout.setContentsMargins(0, 8, 0, 0)
        self.summary_heading = label("統計結果", "summaryHeading")
        summary_layout.addWidget(self.summary_heading)
        self.summary_table = self._table()
        self.summary_table.horizontalHeader().setDefaultSectionSize(90)
        summary_layout.addWidget(self.summary_table)
        splitter.addWidget(summary_panel)
        splitter.setSizes([320, 240])
        right.addWidget(splitter, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        right.addWidget(self.progress)
        self.status = label("尚未匯入資料。請選取 Excel 檔案開始。", "status")
        right.addWidget(self.status)
        layout.addWidget(workspace, 1)
        for combo in (self.numeric_column, self.group_column, self.filter_column):
            combo.currentIndexChanged.connect(self.schedule_refresh)
        self.filter_text.textChanged.connect(self.schedule_refresh)

    @staticmethod
    def _table():
        table = QTableView()
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        table.verticalHeader().setDefaultSectionSize(30)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setDefaultSectionSize(130)
        return table

    @property
    def selected(self) -> Sheet | None:
        index = self.sheet_list.currentRow()
        return self.sheets[index] if 0 <= index < len(self.sheets) else None

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "匯入 Excel（可多選）", "", "Excel 檔案 (*.xlsx *.xls *.xlsb)")
        if paths:
            self.import_paths(paths)

    def import_paths(self, paths):
        if self.busy:
            return
        paths = list(dict.fromkeys(Path(path).resolve() for path in paths))
        if len(paths) > 20:
            QMessageBox.warning(self, "匯入限制", "每次最多選取 20 個檔案。")
            return
        if not paths:
            return
        self.busy = True
        for button in (self.import_button, self.remove_button, self.clear_button, self.report_button):
            button.setEnabled(False)
        self.progress.show()
        self._worker = ImportWorker(paths, self.sheets, self, self.header_row.value() or None)
        self._worker.progress.connect(self.status.setText)
        self._worker.result.connect(self._import_result)
        self._worker.finished.connect(self._import_finished)
        self._worker.start()

    def _import_result(self, sheets, errors):
        for sheet in sheets:
            self.sheets.append(sheet)
            item = QListWidgetItem(sheet.label)
            item.setToolTip(str(sheet.path) + "\n" + sheet.name)
            self.sheet_list.addItem(item)
        if self.sheet_list.currentRow() < 0 and self.sheets:
            self.sheet_list.setCurrentRow(0)
        self.refresh()
        self.status.setText(self.status.text() + f" 本次新增 {len(sheets)} 張工作表。")
        if errors:
            QMessageBox.warning(self, "部分檔案未匯入", "\n\n".join(errors))

    def _import_finished(self):
        self.busy = False
        for button in (self.import_button, self.remove_button, self.clear_button, self.report_button):
            button.setEnabled(True)
        self.progress.hide()
        if self._worker:
            self._worker.deleteLater()
            self._worker = None

    def select_sheet(self, _index):
        self._updating = True
        for combo in (self.numeric_column, self.group_column, self.filter_column):
            combo.clear()
        self.filter_text.clear()
        sheet = self.selected
        self.group_column.addItem("（不分組）", None)
        if sheet is not None:
            columns = list(sheet.data.columns)
            self.numeric_column.addItems(columns)
            self.filter_column.addItems(columns)
            for column in columns:
                self.group_column.addItem(column, column)
            for index, column in enumerate(columns):
                if any(number(value) is not None for value in sheet.data[column]):
                    self.numeric_column.setCurrentIndex(index)
                    break
        self._updating = False
        self.refresh()

    def schedule_refresh(self, *_args):
        if not self._updating:
            self.export_button.setEnabled(False)
            self._timer.start()

    def refresh(self):
        self._timer.stop()
        self.export_button.setEnabled(False)
        self.summary = pd.DataFrame(columns=SUMMARY_COLUMNS)
        self._replace_model(self.summary_table, FrameModel(self.summary, summary=True))
        sheet = self.selected
        if sheet is None:
            self.preview_stack.setCurrentIndex(0)
            self._replace_model(self.preview, FrameModel(pd.DataFrame()))
            self.summary_heading.setText("統計結果")
            self.status.setText("尚未匯入資料。請選取 Excel 檔案開始。")
            return
        try:
            rows = filter_rows(sheet.data, self.filter_column.currentText(), self.filter_text.text())
            self._replace_model(self.preview, FrameModel(rows.iloc[:5000]))
            self.preview_stack.setCurrentIndex(1)
            column = self.numeric_column.currentText()
            if column:
                self.summary = analyze(rows, column, self.group_column.currentData())
                self._replace_model(self.summary_table, FrameModel(self.summary, summary=True))
                self.summary_heading.setText(f"統計結果 · {column}")
                self.export_button.setEnabled(not self.summary.empty)
            self.status.setText(f"{sheet.label} ｜ 篩選後 {len(rows):,} 列，預覽前 {min(len(rows), 5000):,} 列；統計包含全部篩選結果。")
        except (ValueError, OverflowError) as exc:
            self.status.setText(f"無法計算：{exc}")

    @staticmethod
    def _replace_model(table, model):
        old = table.model()
        model.setParent(table)
        table.setModel(model)
        if old:
            old.deleteLater()

    def export(self):
        self.refresh()
        sheet = self.selected
        if sheet is None or self.summary.empty:
            return
        path, _ = QFileDialog.getSaveFileName(self, "匯出統計", "Excel統計.csv", "CSV 檔案 (*.csv)")
        if not path:
            return
        if not path.lower().endswith(".csv"):
            path += ".csv"
        try:
            query = self.filter_text.text()
            export_csv(path, self.summary, file=sheet.path.name, sheet=sheet.name,
                       column=self.numeric_column.currentText(), group=self.group_column.currentData() or "不分組",
                       query=f"{self.filter_column.currentText()} 包含 {query}" if query.strip() else "無")
            self.status.setText(f"已匯出 {len(self.summary):,} 組統計：{path}")
        except OSError as exc:
            QMessageBox.critical(self, "無法匯出", str(exc))

    def open_report(self):
        from .report_ui import ReportDialog
        dialog = ReportDialog(self.sheets, self.report_config, self)
        dialog.exec()
        # Keep edits for this session; explicit JSON save is available in the dialog.
        # Demo rules are intentionally kept inside the demo dialog only.
        if len(dialog.sheets) == len(self.sheets) and all(a is b for a, b in zip(dialog.sheets, self.sheets)):
            self.report_config = dialog.config
        dialog.deleteLater()

    def remove_sheet(self):
        index = self.sheet_list.currentRow()
        if self.busy or index < 0:
            return
        self._updating = True
        self.sheet_list.blockSignals(True)
        self.sheets.pop(index)
        self.sheet_list.takeItem(index)
        self.sheet_list.blockSignals(False)
        self._updating = False
        self.select_sheet(self.sheet_list.currentRow())

    def clear_sheets(self):
        if self.busy:
            return
        self.sheets.clear()
        self.sheet_list.clear()
        self.select_sheet(-1)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if not self.busy and event.mimeData().hasUrls() and all(url.isLocalFile() for url in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if not self.busy and event.mimeData().hasUrls():
            self.import_paths([url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()])
            event.acceptProposedAction()

    def closeEvent(self, event: QCloseEvent):
        if self.busy:
            self.status.setText("正在讀取 Excel，請等待匯入完成後再關閉。")
            event.ignore()
        else:
            event.accept()


def main() -> int:
    import argparse
    import json
    parser = argparse.ArgumentParser(description="Excel Insights 桌面統計工作台")
    parser.add_argument("--smoke-test", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--smoke-output", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft JhengHei UI" if sys.platform == "win32" else "Noto Sans CJK TC", 10))
    window = MainWindow()
    window.show()
    if args.smoke_test:
        if not args.smoke_output:
            parser.error("--smoke-test requires --smoke-output")
        try:
            window._import_result(read_workbook(args.smoke_test), [])
            if window.numeric_column.findText("金額") >= 0:
                window.numeric_column.setCurrentText("金額")
            window.refresh()
            app.processEvents()
            report = {"sheets": len(window.sheets), "preview_rows": window.preview.model().rowCount(),
                      "summary": window.summary.to_dict(orient="records")}
            from .report_ui import ReportDialog
            from .reporting import build_report, demo_sheet, export_report, suggest_config
            demo = demo_sheet()
            config = suggest_config([demo])
            result = build_report([demo], config)
            report_path = args.smoke_output.with_suffix(".xlsx")
            export_report(report_path, result, config, [demo])
            designer = ReportDialog([demo], config)
            designer.show()
            app.processEvents()
            if designer.result is None or designer.result.frame.empty:
                raise ValueError("Report designer produced no preview")
            designer.confirm.setChecked(True)
            model = designer.preview.model()
            if not model.setData(model.index(0, 14), "smoke Remark") or designer.confirm.isChecked():
                raise ValueError("Remark editing did not update the preview confirmation")
            export_report(report_path, designer.result, designer.config, [demo])
            report["manual_remark"] = designer.result.frame.iat[0, 14]
            designer.close()
            report["template_report_rows"] = len(result.frame)
            report["template_report_first_total"] = result.frame.iloc[0]["Total act"]
            args.smoke_output.write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
            window.close()
            return 0
        except Exception as exc:
            args.smoke_output.write_text(json.dumps({"error": str(exc)}, ensure_ascii=False), encoding="utf-8")
            return 1
    return app.exec()
