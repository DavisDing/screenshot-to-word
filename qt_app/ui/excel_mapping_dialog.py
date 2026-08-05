from __future__ import annotations

from typing import Dict, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from qt_app.services.excel_service import ExcelPreview, REQUIRED_ROLES
from qt_app.ui.icons import app_icon
from utils.excel_schema import FIELD_SPECS


class ExcelMappingDialog(QDialog):
    def __init__(self, preview: ExcelPreview, parent=None):
        super().__init__(parent)
        self.preview = preview
        icon = app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setWindowTitle("确认 Excel 列含义")
        self.resize(960, 600)
        self.combos: Dict[str, QComboBox] = {}
        self._build_ui()

    def mapping(self) -> Dict[str, Optional[str]]:
        result: Dict[str, Optional[str]] = {}
        for role, combo in self.combos.items():
            text = combo.currentText()
            result[role] = None if text.startswith("不使用") else text
        return result

    def accept(self) -> None:
        mapping = self.mapping()
        missing = [role for role in REQUIRED_ROLES if not mapping.get(role)]
        if missing:
            QMessageBox.warning(self, "缺少必要列", f"请确认这些必要列：{'、'.join(missing)}")
            return
        super().accept()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(12)

        title = QLabel("确认 Excel 列含义")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        subtitle = QLabel(
            f"文件：{self.preview.file_path.name}    工作表：{self.preview.sheet_name}    表头行：第 {self.preview.header_row} 行    "
            "系统会先自动猜测；列名不标准时请手动选择。"
        )
        subtitle.setObjectName("MutedLabel")
        subtitle.setWordWrap(False)
        root.addWidget(subtitle)

        body = QHBoxLayout()
        body.setSpacing(12)
        root.addLayout(body, stretch=1)

        mapping_card = QFrame()
        mapping_card.setObjectName("Card")
        mapping_layout = QGridLayout(mapping_card)
        mapping_layout.setContentsMargins(16, 14, 16, 14)
        mapping_layout.setHorizontalSpacing(10)
        mapping_layout.setVerticalSpacing(9)
        field_head = QLabel("字段")
        field_head.setObjectName("MutedLabel")
        column_head = QLabel("对应 Excel 列")
        column_head.setObjectName("MutedLabel")
        mapping_layout.addWidget(field_head, 0, 0)
        mapping_layout.addWidget(column_head, 0, 1)

        headers = [c.header for c in self.preview.columns]
        options = ["不使用"] + headers
        for row, (role, spec) in enumerate(FIELD_SPECS.items(), start=1):
            required = " *" if role in REQUIRED_ROLES else ""
            reason = self.preview.reasons.get(role, "")
            tooltip = f"{spec.help_text}\n{reason}".strip()
            label = QLabel(f"{spec.label}{required}")
            label.setToolTip(tooltip)
            label.setCursor(Qt.CursorShape.WhatsThisCursor)
            combo = QComboBox()
            combo.setObjectName("MinimalCombo")
            combo.setMinimumHeight(36)
            combo.setMaxVisibleItems(12)
            combo.addItems(options)
            combo.setToolTip(tooltip)
            suggested = self.preview.suggested_mapping.get(role)
            if suggested in headers:
                combo.setCurrentText(suggested)
            mapping_layout.addWidget(label, row, 0, alignment=Qt.AlignmentFlag.AlignVCenter)
            mapping_layout.addWidget(combo, row, 1)
            self.combos[role] = combo
        mapping_layout.setColumnStretch(1, 1)
        body.addWidget(mapping_card, stretch=3)

        preview_card = QFrame()
        preview_card.setObjectName("Card")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(16, 14, 16, 14)
        preview_layout.setSpacing(10)
        preview_title = QLabel("列预览与样例值")
        preview_title.setObjectName("SectionTitle")
        preview_layout.addWidget(preview_title)
        table = QTableWidget(len(self.preview.columns), 3)
        table.setHorizontalHeaderLabels(["序号", "列名", "样例值"])
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for row, col in enumerate(self.preview.columns):
            table.setItem(row, 0, QTableWidgetItem(str(col.index)))
            table.setItem(row, 1, QTableWidgetItem(col.header))
            table.setItem(row, 2, QTableWidgetItem(" / ".join(col.samples)))
        preview_layout.addWidget(table)
        body.addWidget(preview_card, stretch=4)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("确认")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)
