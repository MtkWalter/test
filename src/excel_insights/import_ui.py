"""Local worksheet selection and adjustable import budgets."""
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QGridLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QPushButton, QSpinBox, QVBoxLayout,
)

from .core import ImportLimits


def recommended_sheet(name: str) -> bool:
    return "pivot" not in name.casefold() and "樞紐" not in name


class ImportSelectionDialog(QDialog):
    def __init__(self, catalog: dict[Path, list[str]], limits: ImportLimits, loaded_cells=0, parent=None):
        super().__init__(parent)
        self.setWindowTitle("選取工作表與匯入上限")
        self.resize(800, 650)
        self.loaded_cells = loaded_cells
        layout = QVBoxLayout(self)
        note = QLabel("只勾選需要的工作表，不會讀取未勾選表格的資料列。名稱含 Pivot／樞紐的彙總表預設不勾選，可自行變更。")
        note.setWordWrap(True); layout.addWidget(note)
        actions = QHBoxLayout()
        for text, mode in (("只選建議明細", "recommended"), ("全選", "all"), ("全部取消", "none")):
            button = QPushButton(text)
            button.clicked.connect(lambda _checked=False, choice=mode: self.choose(choice))
            actions.addWidget(button)
        layout.addLayout(actions)
        self.list = QListWidget()
        self.entries = []
        for path, names in catalog.items():
            for name in names:
                recommended = recommended_sheet(name)
                item = QListWidgetItem(f"{path.name} / {name}" + (" 〔彙總表，預設略過〕" if not recommended else ""))
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if recommended else Qt.CheckState.Unchecked)
                item.setToolTip(str(path))
                self.list.addItem(item)
                self.entries.append((item, path, name))
        layout.addWidget(self.list, 1)
        grid = QGridLayout()
        self.file_millions = QSpinBox(); self.file_millions.setRange(1, 200)
        self.total_millions = QSpinBox(); self.total_millions.setRange(1, 200)
        self.file_millions.setValue(max(1, (limits.max_cells + 999999) // 1_000_000))
        self.total_millions.setValue(max(1, (limits.total_cells + 999999) // 1_000_000))
        grid.addWidget(QLabel("每檔上限（百萬格）"), 0, 0); grid.addWidget(self.file_millions, 1, 0)
        grid.addWidget(QLabel("總載入上限（百萬格）"), 0, 1); grid.addWidget(self.total_millions, 1, 1)
        layout.addLayout(grid)
        info = QLabel("預設每檔 2,000 萬格、總量 3,000 萬格；列數支援 Excel 上限。內部空白欄仍計入資料矩形。提高上限仍取決於可用記憶體，建議先略過不需要的彙總表。")
        info.setWordWrap(True); layout.addWidget(info)
        self.status = QLabel(); layout.addWidget(self.status)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.ok_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        self.ok_button.setText("匯入選取工作表")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.list.itemChanged.connect(self.refresh)
        self.total_millions.valueChanged.connect(self.refresh)
        self.refresh()

    def selection(self):
        result = {}
        for item, path, name in self.entries:
            if item.checkState() == Qt.CheckState.Checked:
                result.setdefault(path, []).append(name)
        return result

    def limits(self):
        return ImportLimits(max_cells=self.file_millions.value() * 1_000_000,
                            total_cells=self.total_millions.value() * 1_000_000)

    def choose(self, mode):
        for item, _path, name in self.entries:
            checked = mode == "all" or mode == "recommended" and recommended_sheet(name)
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def refresh(self, *_args):
        count = sum(len(names) for names in self.selection().values())
        remaining = self.total_millions.value() * 1_000_000 - self.loaded_cells
        self.status.setText(f"已選 {count} 張工作表；目前已載入 {self.loaded_cells:,} 格，總量尚可載入 {max(0, remaining):,} 格。")
        self.ok_button.setEnabled(count > 0 and remaining > 0)
