"""Point-and-click local report designer; presets contain settings, never input rows."""
from copy import deepcopy
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QDialog, QDoubleSpinBox, QFileDialog,
    QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QListWidget,
    QListWidgetItem, QMessageBox, QPushButton, QScrollArea, QSpinBox, QTabWidget,
    QTableView, QVBoxLayout, QWidget,
)

from .core import Sheet, is_blank
from .reporting import (
    BLUE, CATEGORIES, YELLOW, ReportConfig, Rule, build_report, comparison,
    demo_sheet, export_report, load_config, observed_values, save_config, suggest_config,
)
from .schema import FIELD_LABELS, FIELDS, schema_key


class ReportModelMixin:
    """Color only the same comparison cells that Excel colors."""
    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        value = self.frame.iat[index.row(), index.column()] if index.isValid() else None
        if role == Qt.ItemDataRole.DisplayRole and index.isValid():
            if is_blank(value):
                return ""
            if index.column() in {10, 12}:
                return f"{value:.0%}"
            if 1 <= index.column() <= 13:
                return f"{value / self.config.display_unit:,.{self.config.decimals}f}"
        if role == Qt.ItemDataRole.BackgroundRole and index.isValid():
            column = index.column()
            rate = None
            if column in {9, 10}:
                rate = self.frame.iat[index.row(), 10]
            elif column in {11, 12}:
                rate = self.frame.iat[index.row(), 12]
            elif column == 13:
                rate = comparison(self.frame.iat[index.row(), 5], self.frame.iat[index.row(), 6], self.config)[1]
            if rate is not None and not is_blank(rate):
                if rate >= self.config.upper / 100:
                    return QColor("#" + YELLOW[-6:])
                if rate <= self.config.lower / 100:
                    return QColor("#" + BLUE[-6:])
        return super().data(index, role)


class ReportDialog(QDialog):
    def __init__(self, sheets: list[Sheet], config: ReportConfig | None = None, parent=None):
        super().__init__(parent)
        from .app import FrameModel, STYLE

        class ReportModel(ReportModelMixin, FrameModel):
            pass

        self.model_class = ReportModel
        self.sheets = sheets.copy()
        self.config = deepcopy(config) if config else suggest_config(sheets)
        self.result = None
        self._updating = True
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.setInterval(250)
        self._timer.timeout.connect(self.refresh_preview)
        self.setWindowTitle("本機自動判讀＋預覽確認")
        self.resize(1240, 880)
        self.setMinimumSize(1000, 740)
        self.setStyleSheet(STYLE)
        root = QVBoxLayout(self)
        heading = QLabel("本機自動判讀＋預覽確認")
        heading.setStyleSheet("font-size: 23px; font-weight: 600; color: #24324a")
        root.addWidget(heading)
        info = QLabel("資料全程留在本機。自動配對是建議；你可改欄位、分類、期間與算法。修改後重新預覽並確認，再匯出 Excel。")
        info.setWordWrap(True)
        root.addWidget(info)
        actions = QHBoxLayout()
        for text, callback in (("載入示範資料", self.load_demo), ("開啟設定", self.open_settings), ("儲存設定", self.save_settings)):
            button = QPushButton(text)
            button.clicked.connect(callback)
            actions.addWidget(button)
        actions.addStretch()
        root.addLayout(actions)
        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)
        self._source_tab()
        self._rules_tab()
        self._preview_tab()
        footer = QHBoxLayout()
        self.confirm = QCheckBox("我已檢查本次預覽與設定")
        self.confirm.toggled.connect(self._enable_export)
        footer.addWidget(self.confirm)
        footer.addStretch()
        self.export_button = QPushButton("匯出範本 Excel")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self.export)
        footer.addWidget(self.export_button)
        close = QPushButton("關閉")
        close.clicked.connect(self.accept)
        footer.addWidget(close)
        root.addLayout(footer)
        self._populate_sources()
        self._fill_controls()
        self._updating = False
        self.refresh_preview()

    def _source_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.addWidget(QLabel("勾選要合併的工作表。相同欄位結構共用對應設定；每筆來源列只讀取一次。"))
        self.source_list = QListWidget()
        self.source_list.setMaximumHeight(130)
        self.source_list.itemChanged.connect(self.source_changed)
        layout.addWidget(self.source_list)
        self.source_select = QComboBox()
        self.source_select.currentIndexChanged.connect(self.show_mapping)
        layout.addWidget(self.source_select)
        self.header_info = QLabel()
        layout.addWidget(self.header_info)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        editor = QWidget()
        grid = QGridLayout(editor)
        self.mapping_boxes = {}
        for i, (key, title) in enumerate(FIELD_LABELS.items()):
            box = QComboBox()
            self.mapping_boxes[key] = box
            row, col = (i // 2) * 2, i % 2
            grid.addWidget(QLabel(title + (" *" if key in {"node", "amount", "scenario"} else "")), row, col)
            grid.addWidget(box, row + 1, col)
            box.currentIndexChanged.connect(lambda _index, field=key: self.mapping_changed(field))
        scroll.setWidget(editor)
        layout.addWidget(scroll, 1)
        layout.addWidget(QLabel("* 必填。標題列若判讀錯誤，可回主畫面指定標題列，清除該檔後重新匯入。來源檔不會被修改。"))
        self.tabs.addTab(page, "1 來源與欄位")

    def _rules_tab(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        page = QWidget()
        layout = QVBoxLayout(page)
        versions = QHBoxLayout()
        self.actual = QComboBox()
        self.original = QComboBox()
        versions.addWidget(QLabel("act 來源值")); versions.addWidget(self.actual, 1)
        versions.addWidget(QLabel("ori 來源值")); versions.addWidget(self.original, 1)
        layout.addLayout(versions)
        categories = QHBoxLayout()
        self.category_controls = {}
        for name in CATEGORIES:
            group = QGroupBox(name)
            column = QVBoxLayout(group)
            field = QComboBox()
            for key in ("category", "enduser", "product", "supplier", "part"):
                field.addItem(FIELD_LABELS[key], key)
            operator = QComboBox()
            operator.addItem("等於任一選取值", "equals")
            operator.addItem("包含任一選取值", "contains")
            values = QListWidget()
            values.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
            values.setMinimumHeight(100)
            column.addWidget(field); column.addWidget(operator); column.addWidget(values)
            column.addWidget(QLabel("可點選多個值；未選取時該分類留白。"))
            categories.addWidget(group)
            self.category_controls[name] = (field, operator, values)
            field.currentIndexChanged.connect(lambda _i, category=name: self.category_field_changed(category))
            operator.currentIndexChanged.connect(self.changed)
            values.itemSelectionChanged.connect(self.changed)
        layout.addLayout(categories)
        periods = QGridLayout()
        self.year_mode = QComboBox()
        for text, value in (("全部年度", "all"), ("曆年", "calendar"), ("會計年度（以結束年命名）", "fiscal")):
            self.year_mode.addItem(text, value)
        self.year = QSpinBox(); self.year.setRange(1900, 9999)
        self.fiscal_start = QSpinBox(); self.fiscal_start.setRange(1, 12)
        self.month = QComboBox(); self.month.addItem("全部月份", 0)
        for month in range(1, 13): self.month.addItem(f"{month} 月", month)
        for i, (title, widget) in enumerate((("年度模式", self.year_mode), ("年度／FY 結束年", self.year), ("FY 起始月", self.fiscal_start), ("曆月", self.month))):
            periods.addWidget(QLabel(title), 0, i); periods.addWidget(widget, 1, i)
        layout.addLayout(periods)
        self.period_info = QLabel(); self.period_info.setWordWrap(True); layout.addWidget(self.period_info)
        filters = QGridLayout()
        self.filter_boxes = {}
        for i, key in enumerate(("supplier", "part", "product")):
            box = QComboBox(); box.setEditable(True)
            self.filter_boxes[key] = box
            filters.addWidget(QLabel(FIELD_LABELS[key] + " 篩選（可選）"), 0, i)
            filters.addWidget(box, 1, i)
            box.currentTextChanged.connect(self.changed)
        layout.addLayout(filters)
        math_grid = QGridLayout()
        self.difference = QComboBox()
        for text in ("act-ori", "ori-act"): self.difference.addItem(text, text)
        self.percentage = QComboBox()
        self.percentage.addItem("act ÷ ori", "ratio")
        self.percentage.addItem("所選差額 ÷ ori", "change")
        self.upper = QDoubleSpinBox(); self.upper.setRange(-100000, 100000); self.upper.setSuffix(" %")
        self.lower = QDoubleSpinBox(); self.lower.setRange(-100000, 100000); self.lower.setSuffix(" %")
        self.multiplier = QDoubleSpinBox(); self.multiplier.setDecimals(6); self.multiplier.setRange(.000001, 1_000_000)
        self.unit = QComboBox()
        for text, value in (("原始單位", 1), ("千", 1000), ("百萬", 1_000_000)): self.unit.addItem(text, value)
        self.decimals = QSpinBox(); self.decimals.setRange(0, 6)
        self.total_mode = QComboBox()
        self.total_mode.addItem("全部有效明細（含 Service／未分類）", "all")
        self.total_mode.addItem("僅三分類聯集（重疊列不重複）", "classified")
        self.total_row = QCheckBox("附加總計列（百分比重新計算）")
        self.report_label = QLineEdit(); self.report_label.setPlaceholderText("自動由年度設定產生；可自訂表頭")
        widgets = (("差額算法", self.difference), ("百分比算法", self.percentage), ("黃色門檻 ≥", self.upper), ("藍色門檻 ≤", self.lower), ("來源金額倍率", self.multiplier), ("金額顯示單位", self.unit), ("Total 計算範圍", self.total_mode), ("表頭年度文字", self.report_label), ("金額小數位", self.decimals))
        for i, (title, widget) in enumerate(widgets):
            row, col = (i // 2) * 2, i % 2
            math_grid.addWidget(QLabel(title), row, col); math_grid.addWidget(widget, row + 1, col)
        layout.addLayout(math_grid); layout.addWidget(self.total_row)
        layout.addWidget(QLabel("ori 為 0 或缺值時比率留白。顯示單位只改顯示，來源倍率會先換算金額。所有比較規則都可在此修改。"))
        for combo in (self.actual, self.original, self.year_mode, self.month, self.difference, self.unit, self.total_mode):
            combo.currentIndexChanged.connect(self.changed)
        self.percentage.currentIndexChanged.connect(self.percentage_changed)
        for spin in (self.year, self.fiscal_start, self.upper, self.lower, self.multiplier, self.decimals):
            spin.valueChanged.connect(self.changed)
        self.total_row.toggled.connect(self.changed)
        self.report_label.textChanged.connect(self.changed)
        scroll.setWidget(page)
        self.tabs.addTab(scroll, "2 分類、期間與算法")

    def _preview_tab(self):
        page = QWidget(); layout = QVBoxLayout(page)
        self.preview_info = QLabel(); self.preview_info.setWordWrap(True)
        self.preview_info.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.preview_info)
        self.preview = QTableView()
        self.preview.setAlternatingRowColors(True)
        self.preview.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.preview.horizontalHeader().setDefaultSectionSize(125)
        layout.addWidget(self.preview, 1)
        self.tabs.addTab(page, "3 報表預覽")

    def included(self):
        return [sheet for i, sheet in enumerate(self.sheets) if self.source_list.item(i).checkState() == Qt.CheckState.Checked]

    def _populate_sources(self):
        self.source_list.blockSignals(True); self.source_select.blockSignals(True)
        self.source_list.clear(); self.source_select.clear()
        for sheet in self.sheets:
            item = QListWidgetItem(sheet.label)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            mapping = self.config.mapping(sheet)
            enabled = all(mapping.get(k) for k in ("node", "amount", "scenario"))
            item.setCheckState(Qt.CheckState.Checked if enabled else Qt.CheckState.Unchecked)
            self.source_list.addItem(item)
            self.source_select.addItem(sheet.label)
        self.source_list.blockSignals(False); self.source_select.blockSignals(False)
        self.show_mapping()

    def show_mapping(self, *_args):
        old = self._updating; self._updating = True
        index = self.source_select.currentIndex()
        sheet = self.sheets[index] if index >= 0 else None
        self.header_info.setText(f"本表使用第 {sheet.header_row} 列為標題；資料 {len(sheet.data):,} 筆。" if sheet else "尚無來源。可先載入示範資料查看輸出。")
        for key, box in self.mapping_boxes.items():
            box.clear(); box.addItem("（未設定）", None)
            if sheet:
                for column in sheet.data.columns: box.addItem(column, column)
                box.setCurrentIndex(max(0, box.findData(self.config.mapping(sheet).get(key))))
        self._updating = old

    def mapping_changed(self, key):
        if self._updating or self.source_select.currentIndex() < 0: return
        sheet = self.sheets[self.source_select.currentIndex()]
        mapping = dict(self.config.mapping(sheet))
        mapping[key] = self.mapping_boxes[key].currentData()
        self.config.mappings[schema_key(sheet.data.columns)] = mapping
        self._collect_controls()
        self._fill_values()
        self.changed()

    def source_changed(self, *_args):
        if self._updating: return
        self._collect_controls()
        self._fill_values()
        self.changed()

    def _set_combo_values(self, combo, values, selected, placeholder="（未設定）"):
        combo.blockSignals(True)
        combo.clear(); combo.addItem(placeholder, "")
        for value in values[:2000]: combo.addItem(value, value)
        if selected and combo.findData(selected) < 0: combo.addItem(selected, selected)
        combo.setCurrentIndex(max(0, combo.findData(selected)))
        combo.blockSignals(False)

    def _fill_values(self):
        old = self._updating; self._updating = True
        sheets = self.included()
        versions = observed_values(sheets, self.config, "scenario")
        self._set_combo_values(self.actual, versions, self.config.actual)
        self._set_combo_values(self.original, versions, self.config.original)
        for name, (field, operator, values) in self.category_controls.items():
            rule = self.config.rules[name]
            field.setCurrentIndex(max(0, field.findData(rule.field)))
            operator.setCurrentIndex(max(0, operator.findData(rule.operator)))
            self._fill_category(name)
        for key, box in self.filter_boxes.items():
            self._set_combo_values(box, observed_values(sheets, self.config, key), self.config.filters.get(key, ""), "（全部）")
        self._updating = old

    def _fill_category(self, name):
        field, _, values = self.category_controls[name]
        wanted = self.config.rules[name].values
        values.blockSignals(True); values.clear()
        available = observed_values(self.included(), self.config, field.currentData())
        for value in sorted(set(available[:2000]) | set(wanted), key=str.casefold):
            item = QListWidgetItem(value); values.addItem(item)
            item.setSelected(value in wanted)
        values.blockSignals(False)

    def category_field_changed(self, name):
        if self._updating: return
        self._collect_controls()
        self.config.rules[name].values = []
        self._fill_category(name)
        self.changed()

    def percentage_changed(self):
        if self._updating: return
        self.upper.setValue(115 if self.percentage.currentData() == "ratio" else 15)
        self.lower.setValue(85 if self.percentage.currentData() == "ratio" else -15)
        self.changed()

    def _fill_controls(self):
        old = self._updating; self._updating = True
        for box, value in ((self.year_mode, self.config.year_mode), (self.month, self.config.month), (self.difference, self.config.difference), (self.percentage, self.config.percentage), (self.unit, self.config.display_unit), (self.total_mode, self.config.total_mode)):
            box.setCurrentIndex(box.findData(value))
        for box, value in ((self.year, self.config.year), (self.fiscal_start, self.config.fiscal_start), (self.upper, self.config.upper), (self.lower, self.config.lower), (self.multiplier, self.config.multiplier), (self.decimals, self.config.decimals)):
            box.setValue(value)
        self.total_row.setChecked(self.config.total_row); self.report_label.setText(self.config.label)
        self._fill_values()
        self._updating = old

    def _collect_controls(self):
        self.config.actual = self.actual.currentData() or ""
        self.config.original = self.original.currentData() or ""
        for name, (field, operator, values) in self.category_controls.items():
            self.config.rules[name] = Rule(field.currentData(), operator.currentData(), [item.text() for item in values.selectedItems()])
        for attr, box in (("year_mode", self.year_mode), ("month", self.month), ("difference", self.difference), ("percentage", self.percentage), ("display_unit", self.unit), ("total_mode", self.total_mode)):
            setattr(self.config, attr, box.currentData())
        for attr, box in (("year", self.year), ("fiscal_start", self.fiscal_start), ("upper", self.upper), ("lower", self.lower), ("multiplier", self.multiplier), ("decimals", self.decimals)):
            setattr(self.config, attr, box.value())
        self.config.total_row = self.total_row.isChecked(); self.config.label = self.report_label.text()
        self.config.filters = {key: "" if box.currentIndex() == 0 and box.currentText() == "（全部）" else box.currentText() for key, box in self.filter_boxes.items()}

    def changed(self, *_args):
        if self._updating: return
        self.confirm.setChecked(False)
        self.export_button.setEnabled(False)
        self._timer.start()

    def refresh_preview(self):
        self._timer.stop()
        self.confirm.setChecked(False)
        self.result = None
        self._collect_controls()
        self.year.setEnabled(self.config.year_mode != "all")
        self.fiscal_start.setEnabled(self.config.year_mode == "fiscal")
        if self.config.year_mode == "fiscal":
            start_year = self.config.year if self.config.fiscal_start == 1 else self.config.year - 1
            self.period_info.setText(f"FY{self.config.year % 100:02d}：從 {start_year}/{self.config.fiscal_start:02d}/01 起的 12 個月。月份篩選仍使用曆月。")
        else:
            self.period_info.setText("年度及月份依對應的曆年／月份欄位判讀，缺少時使用日期欄位。不限制期間時不猜測會計年度。")
        try:
            self.result = build_report(self.included(), self.config)
            model = self.model_class(self.result.frame)
            model.config = deepcopy(self.config)
            old = self.preview.model(); model.setParent(self.preview); self.preview.setModel(model)
            if old: old.deleteLater()
            counts = self.result.counts
            message = f"來源 {len(self.included())} 張表，共 {counts['input']:,} 筆；納入 {counts['used']:,} 筆，篩選排除 {counts['filtered']:,} 筆；輸出 {len(self.result.frame):,} 列。金額顯示除以 {self.config.display_unit:,}。"
            self.preview_info.setText(message + ("\n" + "\n".join(self.result.warnings) if self.result.warnings else "\n請核對預覽；若不是預期結果，可回前兩頁修改後再預覽。"))
        except (ValueError, OverflowError) as exc:
            old = self.preview.model(); self.preview.setModel(None)
            if old: old.deleteLater()
            self.preview_info.setText(f"尚無法產生預覽：{exc}")
        self._enable_export()

    def _enable_export(self, *_args):
        self.export_button.setEnabled(self.confirm.isChecked() and not self._timer.isActive() and self.result is not None and not self.result.frame.empty)

    def load_demo(self):
        if self.sheets and QMessageBox.question(self, "載入示範", "在此預覽視窗改用虛構示範資料？主畫面已匯入的檔案會保留。") != QMessageBox.StandardButton.Yes:
            return
        self._updating = True
        self.sheets = [demo_sheet()]; self.config = suggest_config(self.sheets)
        self.config.year_mode = "fiscal"; self.config.year = 2027; self.config.fiscal_start = 4; self.config.month = 9
        self._populate_sources(); self._fill_controls()
        self._updating = False
        self.tabs.setCurrentIndex(2)
        self.refresh_preview()

    def open_settings(self):
        path, _ = QFileDialog.getOpenFileName(self, "開啟本機設定", "", "JSON 設定 (*.json)")
        if not path: return
        try:
            config = load_config(path)
            self._updating = True; self.config = config
            self._populate_sources(); self._fill_controls(); self._updating = False
            self.refresh_preview()
        except (OSError, ValueError) as exc:
            self._updating = False
            QMessageBox.warning(self, "無法開啟設定", str(exc))

    def save_settings(self):
        self._collect_controls()
        path, _ = QFileDialog.getSaveFileName(self, "儲存本機設定（不包含交易列）", "report-settings.json", "JSON 設定 (*.json)")
        if not path: return
        if not path.lower().endswith(".json"): path += ".json"
        try:
            save_config(path, self.config)
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "無法儲存設定", str(exc))

    def export(self):
        if not self.export_button.isEnabled() or self.result is None: return
        path, _ = QFileDialog.getSaveFileName(self, "匯出報表", "報表預覽.xlsx", "Excel 報表 (*.xlsx)")
        if not path: return
        if not path.lower().endswith(".xlsx"): path += ".xlsx"
        try:
            export_report(path, self.result, self.config, self.sheets)
            self.preview_info.setText(self.preview_info.text() + f"\n已匯出：{path}")
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "無法匯出", str(exc))

    def done(self, code):
        self._timer.stop()
        self._collect_controls()
        super().done(code)
