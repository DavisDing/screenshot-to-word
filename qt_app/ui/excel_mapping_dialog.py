from __future__ import annotations

from typing import Dict, Optional

from PySide6.QtCore import QEvent, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QTableWidget,
    QToolTip,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from qt_app.services.excel_service import ExcelPreview, REQUIRED_ROLES
from qt_app.ui.icons import app_icon
from utils.excel_schema import FIELD_SPECS


class ExcelMappingDialog(QDialog):
    DEFAULT_HINT = "把鼠标移动到左侧字段或下拉框，可查看字段说明和自动识别依据。"

    def __init__(self, preview: ExcelPreview, parent=None):
        super().__init__(parent)
        self.preview = preview
        self.combos: Dict[str, QComboBox] = {}
        self._hover_hints: Dict[QWidget, str] = {}
        icon = app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setWindowTitle("确认 Excel 列含义")
        self.resize(980, 600)
        self._build_ui()

    def mapping(self) -> Dict[str, Optional[str]]:
        result: Dict[str, Optional[str]] = {}
        for role, combo in self.combos.items():
            text = combo.currentText()
            result[role] = None if text.startswith("不使用") else text
        return result

    def eventFilter(self, watched, event):  # noqa: N802 - Qt API
        hint = self._hover_hints.get(watched)
        if hint and hasattr(self, "hint_label"):
            if event.type() == QEvent.Type.Enter:
                self.hint_label.setText(hint)
            elif event.type() == QEvent.Type.ToolTip and hasattr(event, "globalPos"):
                self.hint_label.setText(hint)
                QToolTip.showText(event.globalPos(), hint, watched)
                return True
            elif event.type() == QEvent.Type.Leave:
                self.hint_label.setText(self.DEFAULT_HINT)
        return super().eventFilter(watched, event)

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

        body.addWidget(self._build_mapping_card(), stretch=2)
        body.addWidget(self._build_preview_card(), stretch=5)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("确认")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _build_mapping_card(self) -> QFrame:
        mapping_card = QFrame()
        mapping_card.setObjectName("Card")
        mapping_card.setMinimumWidth(340)
        mapping_card.setMaximumWidth(430)
        outer = QVBoxLayout(mapping_card)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("字段映射（* 必填）")
        title.setObjectName("SectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        outer.addLayout(header)

        headers = [c.header for c in self.preview.columns]
        options = ["不使用"] + headers
        for role, spec in FIELD_SPECS.items():
            required = role in REQUIRED_ROLES
            reason = self.preview.reasons.get(role, "")
            hint = f"{spec.label}：{spec.help_text}"
            if reason:
                hint += f"\n识别依据：{reason}"
            row = QFrame()
            row.setObjectName("MappingRow")
            row.setToolTip(hint)
            row.installEventFilter(self)
            self._hover_hints[row] = hint
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(10, 7, 10, 7)
            row_layout.setSpacing(10)

            label_text = spec.label + (" *" if required else "")
            label = QLabel(label_text)
            label.setObjectName("MappingField")
            label.setFixedWidth(72)
            label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            label.setToolTip(hint)
            label.setCursor(Qt.CursorShape.WhatsThisCursor)
            label.installEventFilter(self)
            self._hover_hints[label] = hint

            combo = QComboBox()
            combo.setObjectName("MinimalCombo")
            combo.setMinimumHeight(34)
            combo.setMaxVisibleItems(12)
            combo.addItems(options)
            combo.setToolTip(hint)
            combo.setCursor(Qt.CursorShape.PointingHandCursor)
            combo.installEventFilter(self)
            self._hover_hints[combo] = hint
            suggested = self.preview.suggested_mapping.get(role)
            if suggested in headers:
                combo.setCurrentText(suggested)

            row_layout.addWidget(label)
            row_layout.addWidget(combo, stretch=1)
            outer.addWidget(row)
            self.combos[role] = combo

        self.hint_label = QLabel(self.DEFAULT_HINT)
        self.hint_label.setObjectName("InlineHint")
        self.hint_label.setWordWrap(True)
        self.hint_label.setMinimumHeight(42)
        outer.addWidget(self.hint_label)
        outer.addStretch(1)
        return mapping_card

    def _build_preview_card(self) -> QFrame:
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
        return preview_card
