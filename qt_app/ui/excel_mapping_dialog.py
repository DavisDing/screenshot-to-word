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
    QLabel,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from qt_app.services.excel_service import ExcelPreview, REQUIRED_ROLES
from utils.excel_schema import FIELD_SPECS


class ExcelMappingDialog(QDialog):
    def __init__(self, preview: ExcelPreview, parent=None):
        super().__init__(parent)
        self.preview = preview
        self.setWindowTitle("确认 Excel 列含义")
        self.resize(900, 620)
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
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        title = QLabel("Excel 列名识别不死板：请确认每一列代表什么")
        title.setObjectName("PageTitle")
        root.addWidget(title)

        subtitle = QLabel(
            f"文件：{self.preview.file_path.name}    工作表：{self.preview.sheet_name}    表头行：第 {self.preview.header_row} 行\n"
            "系统会先按别名自动猜测；如果列名不标准，请在下面手动选择。"
        )
        subtitle.setObjectName("MutedLabel")
        subtitle.setWordWrap(True)
        root.addWidget(subtitle)

        body = QHBoxLayout()
        body.setSpacing(14)
        root.addLayout(body, stretch=1)

        mapping_card = QFrame()
        mapping_card.setObjectName("Card")
        mapping_layout = QGridLayout(mapping_card)
        mapping_layout.setContentsMargins(16, 16, 16, 16)
        mapping_layout.setHorizontalSpacing(12)
        mapping_layout.setVerticalSpacing(10)
        mapping_layout.addWidget(QLabel("字段"), 0, 0)
        mapping_layout.addWidget(QLabel("请选择对应列"), 0, 1)
        mapping_layout.addWidget(QLabel("说明 / 识别依据"), 0, 2)

        headers = [c.header for c in self.preview.columns]
        options = ["不使用 / 空"] + headers
        for row, (role, spec) in enumerate(FIELD_SPECS.items(), start=1):
            required = " *" if role in REQUIRED_ROLES else ""
            label = QLabel(f"{spec.label}{required}")
            combo = QComboBox()
            combo.setMinimumHeight(38)
            combo.setMaxVisibleItems(12)
            combo.addItems(options)
            suggested = self.preview.suggested_mapping.get(role)
            if suggested in headers:
                combo.setCurrentText(suggested)
            help_text = QLabel(f"{spec.help_text}\n{self.preview.reasons.get(role, '')}")
            help_text.setObjectName("MutedLabel")
            help_text.setWordWrap(True)
            mapping_layout.addWidget(label, row, 0, alignment=Qt.AlignmentFlag.AlignTop)
            mapping_layout.addWidget(combo, row, 1, alignment=Qt.AlignmentFlag.AlignTop)
            mapping_layout.addWidget(help_text, row, 2)
            self.combos[role] = combo
        body.addWidget(mapping_card, stretch=2)

        preview_card = QFrame()
        preview_card.setObjectName("Card")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(16, 16, 16, 16)
        preview_layout.addWidget(QLabel("列预览与样例值"))
        table = QTableWidget(len(self.preview.columns), 3)
        table.setHorizontalHeaderLabels(["序号", "列名", "样例值"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        for row, col in enumerate(self.preview.columns):
            table.setItem(row, 0, QTableWidgetItem(str(col.index)))
            table.setItem(row, 1, QTableWidgetItem(col.header))
            table.setItem(row, 2, QTableWidgetItem(" / ".join(col.samples)))
        table.resizeColumnsToContents()
        preview_layout.addWidget(table)
        body.addWidget(preview_card, stretch=3)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("确认映射")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)
