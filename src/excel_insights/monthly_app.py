"""Desktop interface for the separate MonthlyReport executable."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PySide6.QtCore import QDate, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QCloseEvent, QDesktopServices, QFont
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox,
    QProgressBar, QPushButton, QSpinBox, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget,
)

from .monthly_report import ReportRules, fiscal_label, generate_report, inspect_workbook


class ClassificationsDialog(QDialog):
    def __init__(self, rules: ReportRules, parent=None):
        super().__init__(parent)
        self.setWindowTitle("客群與 Tech Node 分類設定")
        self.resize(640, 570)
        layout = QVBoxLayout(self)
        explanation = QLabel("設定會記錄在輸出 Excel 的「報表資訊」。未列入對照的 Tech Node 保留原名；Service 的 Tech Node 空白時獨立列為 Service。")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)
        form = QFormLayout()
        self.tsmc = QLineEdit("、".join(rules.tsmc_values))
        self.non_t = QLineEdit("、".join(rules.non_t_values or []))
        self.non_t.setPlaceholderText("留白：所有非空且非 TSMC 的 BF 值，包含 FI")
        self.start_month = QSpinBox()
        self.start_month.setRange(1, 12)
        self.start_month.setValue(rules.fiscal_start_month)
        form.addRow("TSMC 的 BF 值", self.tsmc)
        form.addRow("Non-T 的 BF 值", self.non_t)
        form.addRow("會計年度起始月份", self.start_month)
        layout.addLayout(form)
        layout.addWidget(QLabel("Tech Node 合併對照"))
        self.aliases = QTableWidget(0, 2)
        self.aliases.setHorizontalHeaderLabels(["INPUT 的 Tech Node", "OUTPUT 的 Tech Node"])
        self.aliases.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        for source, target in rules.node_aliases.items():
            self._add(source, target)
        layout.addWidget(self.aliases)
        actions = QHBoxLayout()
        add = QPushButton("新增對照")
        add.clicked.connect(lambda: self._add("", ""))
        remove = QPushButton("移除選取對照")
        remove.clicked.connect(lambda: self.aliases.removeRow(self.aliases.currentRow()) if self.aliases.currentRow() >= 0 else None)
        load = QPushButton("載入設定")
        load.clicked.connect(self.load_rules)
        save = QPushButton("另存設定")
        save.clicked.connect(self.save_rules)
        for button in (add, remove, load, save):
            actions.addWidget(button)
        layout.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _add(self, source, target):
        row = self.aliases.rowCount()
        self.aliases.insertRow(row)
        self.aliases.setItem(row, 0, QTableWidgetItem(source))
        self.aliases.setItem(row, 1, QTableWidgetItem(target))

    def rules(self) -> ReportRules:
        import re
        split = lambda value: [v.strip() for v in re.split(r"[,，、;；\n]", value) if v.strip()]
        pairs = []
        for row in range(self.aliases.rowCount()):
            a = self.aliases.item(row, 0)
            b = self.aliases.item(row, 1)
            source, target = a.text().strip() if a else "", b.text().strip() if b else ""
            if source or target:
                pairs.append((source, target))
        if len({a.casefold() for a, _ in pairs}) != len(pairs):
            raise ValueError("Tech Node 對照不能重複。")
        return ReportRules(tsmc_values=split(self.tsmc.text()), non_t_values=split(self.non_t.text()) or None,
                           node_aliases=dict(pairs), fiscal_start_month=self.start_month.value())

    def accept(self):
        try:
            self.rules()
        except ValueError as exc:
            QMessageBox.warning(self, "分類設定需要修正", str(exc))
            return
        super().accept()

    def load_rules(self):
        path, _ = QFileDialog.getOpenFileName(self, "載入分類設定", "", "分類設定 (*.json)")
        if not path:
            return
        try:
            rules = ReportRules.load(path)
        except Exception as exc:
            QMessageBox.warning(self, "無法載入", str(exc))
            return
        self.tsmc.setText("、".join(rules.tsmc_values))
        self.non_t.setText("、".join(rules.non_t_values or []))
        self.start_month.setValue(rules.fiscal_start_month)
        self.aliases.setRowCount(0)
        for a, b in rules.node_aliases.items():
            self._add(a, b)

    def save_rules(self):
        try:
            data = self.rules().to_dict()
            path, _ = QFileDialog.getSaveFileName(self, "另存分類設定", "monthly-report-rules.json", "分類設定 (*.json)")
            if path:
                Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            QMessageBox.warning(self, "無法儲存", str(exc))


class ReportWorker(QThread):
    progress = Signal(str)
    result = Signal(object)
    failed = Signal(str)

    def __init__(self, arguments: dict, parent=None):
        super().__init__(parent)
        self.arguments = arguments

    def run(self):
        try:
            self.result.emit(generate_report(**self.arguments, progress=self.progress.emit))
        except Exception as exc:
            self.failed.emit(str(exc))


class ReportWindow(QMainWindow):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Monthly Report · 單月 Excel 報表產生器")
        self.resize(900, 700)
        self.setMinimumSize(760, 580)
        self.rules = ReportRules()
        self._worker: ReportWorker | None = None
        self.busy = False
        self.last_result: dict | None = None
        self._confirmed_output: Path | None = None
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(30, 24, 30, 24)
        title = QLabel("單月 Excel 報表產生器")
        title.setStyleSheet("font-size: 24px; font-weight: 600;")
        layout.addWidget(title)
        description = QLabel("選擇 INPUT 與報表年月，產生 monthly pivot、完整差異明細與報表資訊。\n金額來源 AS；主表百萬元整數、Remark 一位小數。資料在本機處理。")
        description.setWordWrap(True)
        layout.addWidget(description)
        form = QFormLayout()
        self.input_path = QLineEdit()
        self.input_path.setReadOnly(True)
        self.input_button = QPushButton("選擇 INPUT")
        self.input_button.clicked.connect(self.choose_input)
        form.addRow("INPUT Excel", self._path_row(self.input_path, self.input_button))
        self.sheet_combo = QComboBox()
        form.addRow("來源工作表", self.sheet_combo)
        period_row = QWidget()
        period_layout = QHBoxLayout(period_row)
        period_layout.setContentsMargins(0, 0, 0, 0)
        today = QDate.currentDate()
        self.year = QSpinBox()
        self.year.setRange(1900, 9999)
        self.year.setValue(today.year())
        self.month = QComboBox()
        self.month.addItems([f"{m:02d}" for m in range(1, 13)])
        self.month.setCurrentIndex(today.month() - 1)
        self.fiscal_year = QLineEdit()
        self.fiscal_year.setMaximumWidth(100)
        for widget in (self.year, QLabel("年"), self.month, QLabel("月　年度標籤"), self.fiscal_year):
            period_layout.addWidget(widget)
        period_layout.addStretch()
        form.addRow("報表期間（單月）", period_row)
        self.year.valueChanged.connect(self.update_period)
        self.month.currentIndexChanged.connect(self.update_period)
        self.update_period()
        self.output_path = QLineEdit()
        self.output_button = QPushButton("選擇 OUTPUT")
        self.output_button.clicked.connect(self.choose_output)
        form.addRow("OUTPUT Excel", self._path_row(self.output_path, self.output_button))
        layout.addLayout(form)
        controls = QHBoxLayout()
        self.rules_button = QPushButton("客群／Tech Node 分類設定")
        self.rules_button.clicked.connect(self.edit_rules)
        controls.addWidget(self.rules_button)
        controls.addStretch()
        self.generate_button = QPushButton("產生 OUTPUT Excel")
        self.generate_button.setStyleSheet("background: #346ce0; color: white; padding: 12px 20px; font-weight: bold;")
        self.generate_button.clicked.connect(self.generate)
        controls.addWidget(self.generate_button)
        layout.addLayout(controls)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.status = QLabel("請選擇 INPUT Excel。")
        self.status.setWordWrap(True)
        layout.addWidget(self.progress)
        layout.addWidget(self.status)
        self.messages = QTextEdit()
        self.messages.setReadOnly(True)
        self.messages.setPlainText("Remark：TSMC／Non-T 各自判斷達成率 ≥115% 或 ≤85%。\n客戶只用 BI END USER，全部列出；同方向優先，再按差額絕對值遞減。\n品號全部列出，按該客戶差額方向優先，再按差額絕對值遞減。\n預算為零、缺值或負數時達成率空白。\n\n預設分類：TSMC 以外的非空 BF 值歸 Non-T（含 FI）。\nN2／A16 合併；N3／N3E 合併；其他節點保留原名。\n如與實際分類不同，請先調整「分類設定」。")
        layout.addWidget(self.messages, 1)
        self.open_button = QPushButton("開啟產出的 Excel")
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(self.open_output)
        layout.addWidget(self.open_button)
        self.setCentralWidget(root)

    @staticmethod
    def _path_row(edit, button):
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(edit, 1)
        layout.addWidget(button)
        return row

    def update_period(self):
        self.fiscal_year.setText(fiscal_label(self.year.value(), self.month.currentIndex() + 1, self.rules.fiscal_start_month))

    def load_input(self, path: str | Path):
        layouts = inspect_workbook(path)
        self.input_path.setText(str(Path(path).resolve()))
        self.sheet_combo.clear()
        self.sheet_combo.addItems([s.name for s in layouts])
        self.output_path.setText(str(Path(path).resolve().with_name(f"monthly-output-{self.year.value()}{self.month.currentIndex() + 1:02d}.xlsx")))
        self._confirmed_output = None
        self.status.setText(f"已找到 {len(layouts)} 張可用工作表，請確認報表年月與分類設定。")

    def choose_input(self):
        path, _ = QFileDialog.getOpenFileName(self, "選擇 INPUT Excel", "", "Excel (*.xlsx *.xlsm)")
        if path:
            try:
                self.load_input(path)
            except Exception as exc:
                QMessageBox.warning(self, "無法讀取 INPUT", str(exc))

    def choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "選擇 OUTPUT Excel", self.output_path.text(), "Excel (*.xlsx)")
        if path:
            if not Path(path).suffix:
                path += ".xlsx"
            self.output_path.setText(path)
            self._confirmed_output = Path(path).resolve()

    def edit_rules(self):
        dialog = ClassificationsDialog(self.rules, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.rules = dialog.rules()
            self.update_period()

    def _set_busy(self, busy: bool):
        self.busy = busy
        for widget in (self.input_button, self.output_button, self.rules_button, self.generate_button,
                       self.year, self.month, self.fiscal_year, self.sheet_combo, self.output_path):
            widget.setEnabled(not busy)
        self.progress.setVisible(busy)
        self.progress.setRange(0, 0 if busy else 100)

    def generate(self):
        if self.busy:
            return
        if not self.input_path.text() or not self.output_path.text():
            self.on_error("請先選擇 INPUT 與 OUTPUT 檔案。")
            return
        output = Path(self.output_path.text()).resolve()
        overwrite = output.exists()
        if overwrite and self._confirmed_output != output:
            answer = QMessageBox.question(self, "OUTPUT 已存在", f"是否覆寫以下檔案？\n{output}",
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                          QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.open_button.setEnabled(False)
        self.last_result = None
        self._set_busy(True)
        self.status.setText("正在讀取與計算…")
        self._worker = ReportWorker({"source": self.input_path.text(), "output": output,
                                     "year": self.year.value(), "month": self.month.currentIndex() + 1,
                                     "sheet": self.sheet_combo.currentText(), "fiscal_year": self.fiscal_year.text(),
                                     "rules": self.rules, "overwrite": overwrite}, self)
        self._worker.progress.connect(self.status.setText)
        self._worker.result.connect(self.on_result)
        self._worker.failed.connect(self.on_error)
        self._worker.finished.connect(self.on_finished)
        self._worker.start()

    def on_result(self, result: dict):
        self.last_result = result
        self.status.setText(f"完成：實績 {result['act_rows']} 列、預算 {result['ori_rows']} 列；{result['remark_nodes']} 個 Tech Node 需要 Remark。")
        message = f"已產生：{result['output']}\n\n"
        message += "\n".join(result["warnings"]) if result["warnings"] else "資料檢查完成。完整差異可在輸出檔的「差異明細」查看。"
        self.messages.setPlainText(message)
        self.open_button.setEnabled(True)
        self.completed.emit(result)

    def on_error(self, message: str):
        self.status.setText("尚未產生報表，請依下方訊息修正。")
        self.messages.setPlainText(message)
        self.failed.emit(message)

    def on_finished(self):
        self._set_busy(False)
        if self._worker:
            self._worker.deleteLater()
            self._worker = None

    def open_output(self):
        if self.last_result:
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.last_result["output"]))

    def closeEvent(self, event: QCloseEvent):
        if self.busy:
            event.ignore()
            self.status.setText("正在處理來源資料，請等待完成後關閉。")
        else:
            super().closeEvent(event)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Monthly Report 桌面報表產生器")
    parser.add_argument("--smoke-test", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--smoke-output", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--output", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--year", type=int, default=2026, help=argparse.SUPPRESS)
    parser.add_argument("--month", type=int, default=9, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.smoke_test and (not args.smoke_output or not args.output):
        parser.error("--smoke-test requires --smoke-output and --output")
    app = QApplication.instance() or QApplication(sys.argv[:1])
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft JhengHei UI" if sys.platform == "win32" else "Noto Sans CJK TC", 10))
    window = ReportWindow()
    window.show()
    if args.smoke_test:
        outcome = {"error": "報表產生器未完成"}
        def success(result):
            outcome.clear()
            outcome.update(result)
        def failure(message):
            outcome.clear()
            outcome["error"] = message
        def finish():
            args.smoke_output.write_text(json.dumps(outcome, ensure_ascii=False, indent=2), encoding="utf-8")
            app.exit(1 if "error" in outcome else 0)
        window.completed.connect(success)
        window.failed.connect(failure)
        try:
            window.year.setValue(args.year)
            window.month.setCurrentIndex(args.month - 1)
            window.load_input(args.smoke_test)
            window.output_path.setText(str(args.output))
            QTimer.singleShot(0, window.generate)
            # Poll only the completion state, letting Qt deliver all worker signals.
            timer = QTimer(window)
            timer.setInterval(25)
            def check_finished():
                if not window.busy and (window.last_result is not None or "error" in outcome and outcome["error"] != "報表產生器未完成"):
                    timer.stop()
                    finish()
            timer.timeout.connect(check_finished)
            timer.start()
        except Exception as exc:
            failure(str(exc))
            QTimer.singleShot(0, finish)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
