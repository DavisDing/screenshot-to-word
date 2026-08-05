from __future__ import annotations

import time
from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
)

from qt_app.models.test_case import TestItem
from qt_app.services.excel_service import ExcelService
from qt_app.services.screenshot_service import ScreenshotService
from qt_app.services.word_service import WordService
from qt_app.ui.annotator_window import AnnotatorWindow


class ControlWindow(QDialog):
    def __init__(
        self,
        items: List[TestItem],
        excel_service: ExcelService,
        screenshot_service: ScreenshotService,
        word_service: WordService,
        logger,
        parent=None,
    ):
        super().__init__(parent)
        self.items = items
        self.excel_service = excel_service
        self.screenshot_service = screenshot_service
        self.word_service = word_service
        self.logger = logger
        self.index = 0
        self.last_image: Optional[Path] = None
        self._hidden_parent = None
        self.setWindowTitle("执行控制面板")
        self.resize(900, 620)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self._build_ui()
        self._refresh()
        QShortcut(QKeySequence("F8"), self, activated=self.capture_and_annotate)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(14)

        header = QFrame()
        header.setObjectName("Card")
        h = QGridLayout(header)
        h.setContentsMargins(18, 16, 18, 16)
        title = QLabel("执行控制面板")
        title.setObjectName("PageTitle")
        self.progress_label = QLabel("-")
        self.progress_label.setObjectName("MutedLabel")
        self.progress = QProgressBar()
        h.addWidget(title, 0, 0)
        h.addWidget(self.progress_label, 0, 1, alignment=Qt.AlignmentFlag.AlignRight)
        h.addWidget(self.progress, 1, 0, 1, 2)
        root.addWidget(header)

        card = QFrame()
        card.setObjectName("Card")
        grid = QGridLayout(card)
        grid.setContentsMargins(18, 18, 18, 18)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(12)
        self.case_label = self._value_label()
        self.checkpoint_label = self._value_label()
        self.step_label = self._value_label()
        self.desc_label = self._value_label()
        self.expected_label = self._value_label()
        rows = [
            ("测试名称", self.case_label),
            ("验证点", self.checkpoint_label),
            ("步骤名称", self.step_label),
            ("步骤描述", self.desc_label),
            ("预期结果", self.expected_label),
        ]
        for row, (name, widget) in enumerate(rows):
            key = QLabel(name)
            key.setObjectName("MutedLabel")
            grid.addWidget(key, row, 0, alignment=Qt.AlignmentFlag.AlignTop)
            grid.addWidget(widget, row, 1)
        grid.setColumnStretch(1, 1)
        root.addWidget(card, stretch=1)

        buttons = QHBoxLayout()
        self.capture_btn = QPushButton("截图并标注 (F8)")
        self.capture_btn.setObjectName("PrimaryButton")
        self.capture_btn.clicked.connect(self.capture_and_annotate)
        self.finish_btn = QPushButton("完成当前")
        self.finish_btn.clicked.connect(self.finish_current)
        self.skip_btn = QPushButton("跳过")
        self.skip_btn.clicked.connect(self.skip_current)
        self.close_btn = QPushButton("结束")
        self.close_btn.setObjectName("DangerButton")
        self.close_btn.clicked.connect(self.reject)
        buttons.addWidget(self.capture_btn)
        buttons.addWidget(self.finish_btn)
        buttons.addWidget(self.skip_btn)
        buttons.addStretch(1)
        buttons.addWidget(self.close_btn)
        root.addLayout(buttons)

    def _value_label(self) -> QLabel:
        label = QLabel("-")
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        return label

    def _current(self) -> Optional[TestItem]:
        if 0 <= self.index < len(self.items):
            return self.items[self.index]
        return None

    def _refresh(self) -> None:
        total = len(self.items)
        item = self._current()
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(min(self.index, total))
        self.progress_label.setText(f"{min(self.index + 1, total)}/{total}" if item else f"{total}/{total}")
        if item is None:
            self._finish_all()
            return
        self.last_image = None
        self.case_label.setText(item.case_name)
        self.checkpoint_label.setText(item.checkpoint)
        self.step_label.setText(item.step_name or "-")
        self.desc_label.setText(item.step_description or "-")
        self.expected_label.setText(item.expected_result or "-")
        self.finish_btn.setEnabled(False)

    def capture_and_annotate(self) -> None:
        item = self._current()
        if not item:
            return
        self.logger.log(f"开始截图：{item.display_title}")
        self.setWindowOpacity(0.0)
        parent = self.parentWidget()
        self._hidden_parent = parent
        if parent is not None:
            parent.setWindowOpacity(0.0)
        # Give the window manager a moment to hide this topmost panel before capture.
        QTimer.singleShot(250, lambda: self._do_capture(item))

    def _do_capture(self, item: TestItem) -> None:
        image_path = self.screenshot_service.capture_fullscreen(item.case_name)
        self.setWindowOpacity(1.0)
        if self._hidden_parent is not None:
            self._hidden_parent.setWindowOpacity(1.0)
            self._hidden_parent = None
        if not image_path:
            QMessageBox.warning(self, "截图失败", "无法完成截图，请检查系统截图权限。")
            return
        annotator = AnnotatorWindow(image_path, self)
        if annotator.exec() == QDialog.DialogCode.Accepted and annotator.result_path:
            self.last_image = annotator.result_path
            self.finish_btn.setEnabled(True)
            self.logger.log(f"标注完成：{self.last_image}")
        else:
            self.logger.log("用户取消标注")

    def finish_current(self) -> None:
        item = self._current()
        if not item or not self.last_image:
            QMessageBox.information(self, "提示", "请先截图并保存标注。")
            return
        try:
            self.word_service.add_image(item.case_name, item.checkpoint, self.last_image, item.step_note())
            self.excel_service.mark_executed([item.row_index])
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            self.logger.log(f"保存当前用例失败：{exc}")
            return
        self.index += 1
        self._refresh()

    def skip_current(self) -> None:
        item = self._current()
        if item:
            self.logger.log(f"跳过：{item.display_title}")
        self.index += 1
        self._refresh()

    def _finish_all(self) -> None:
        self.progress.setValue(len(self.items))
        self.capture_btn.setEnabled(False)
        self.finish_btn.setEnabled(False)
        self.skip_btn.setEnabled(False)
        QMessageBox.information(self, "完成", "所有待执行用例已处理完毕。")
        self.accept()
