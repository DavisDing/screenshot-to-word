from __future__ import annotations

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
    QVBoxLayout,
)

from qt_app.models.test_case import TestItem
from qt_app.services.excel_service import ExcelService
from qt_app.services.screenshot_service import ScreenshotService
from qt_app.services.word_service import WordService
from qt_app.ui.annotator_window import AnnotatorWindow


class ControlWindow(QDialog):
    """Small floating execution panel.

    It intentionally stays compact and semi-transparent so it does not cover the
    system-under-test. During screenshot capture the panel is hidden completely.
    """

    PANEL_OPACITY = 0.88

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
        self.setWindowTitle("执行控制")
        self.resize(430, 500)
        self.setMinimumSize(390, 430)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowOpacity(self.PANEL_OPACITY)
        self._build_ui()
        self._refresh()
        self._move_to_right_side()
        QShortcut(QKeySequence("F8"), self, activated=self.capture_and_annotate)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        header = QFrame()
        header.setObjectName("FloatingCard")
        h = QGridLayout(header)
        h.setContentsMargins(12, 10, 12, 10)
        h.setHorizontalSpacing(8)
        title = QLabel("执行控制")
        title.setObjectName("PanelTitle")
        self.progress_label = QLabel("-")
        self.progress_label.setObjectName("MutedLabel")
        self.progress = QProgressBar()
        h.addWidget(title, 0, 0)
        h.addWidget(self.progress_label, 0, 1, alignment=Qt.AlignmentFlag.AlignRight)
        h.addWidget(self.progress, 1, 0, 1, 2)
        root.addWidget(header)

        card = QFrame()
        card.setObjectName("FloatingCard")
        grid = QGridLayout(card)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(8)
        self.case_label = self._value_label()
        self.checkpoint_label = self._value_label()
        self.step_label = self._value_label()
        self.desc_label = self._value_label()
        self.expected_label = self._value_label()
        rows = [
            ("用例", self.case_label),
            ("验证", self.checkpoint_label),
            ("步骤", self.step_label),
            ("描述", self.desc_label),
            ("预期", self.expected_label),
        ]
        for row, (name, widget) in enumerate(rows):
            key = QLabel(name)
            key.setObjectName("PanelKey")
            grid.addWidget(key, row, 0, alignment=Qt.AlignmentFlag.AlignTop)
            grid.addWidget(widget, row, 1)
        grid.setColumnStretch(1, 1)
        root.addWidget(card, stretch=1)

        buttons1 = QHBoxLayout()
        self.capture_btn = QPushButton("截图并标注 F8")
        self.capture_btn.setObjectName("PrimaryButton")
        self.capture_btn.clicked.connect(self.capture_and_annotate)
        self.finish_btn = QPushButton("完成")
        self.finish_btn.clicked.connect(self.finish_current)
        buttons1.addWidget(self.capture_btn, stretch=2)
        buttons1.addWidget(self.finish_btn, stretch=1)
        root.addLayout(buttons1)

        buttons2 = QHBoxLayout()
        self.skip_btn = QPushButton("跳过")
        self.skip_btn.clicked.connect(self.skip_current)
        self.close_btn = QPushButton("结束")
        self.close_btn.setObjectName("DangerButton")
        self.close_btn.clicked.connect(self.reject)
        buttons2.addWidget(self.skip_btn)
        buttons2.addWidget(self.close_btn)
        root.addLayout(buttons2)

        tip = QLabel("面板半透明；截图时自动隐藏，不会进入截图。")
        tip.setObjectName("PanelTip")
        tip.setWordWrap(True)
        root.addWidget(tip)

    def _value_label(self) -> QLabel:
        label = QLabel("-")
        label.setObjectName("PanelValue")
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        return label

    def _move_to_right_side(self) -> None:
        screen = self.screen()
        if not screen:
            return
        geo = screen.availableGeometry()
        self.move(geo.right() - self.width() - 24, geo.top() + 64)

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
        self.logger.log(f"[截图] 开始截图：{item.display_title}")
        self.hide()
        QTimer.singleShot(280, lambda: self._do_capture(item))

    def _do_capture(self, item: TestItem) -> None:
        image_path = self.screenshot_service.capture_fullscreen(item.case_name)
        self.show()
        self.raise_()
        self.activateWindow()
        if not image_path:
            QMessageBox.warning(self, "截图失败", "无法完成截图，请检查系统截图权限。")
            return
        annotator = AnnotatorWindow(image_path, self)
        annotator.setWindowOpacity(1.0)
        if annotator.exec() == QDialog.DialogCode.Accepted and annotator.result_path:
            self.last_image = annotator.result_path
            self.finish_btn.setEnabled(True)
            self.logger.log(f"[标注] 标注完成：{self.last_image}")
        else:
            self.logger.log("[标注] 用户取消标注")

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
            self.logger.log(f"[错误] 保存当前用例失败：{exc}")
            return
        self.logger.log(f"[完成] 当前用例已完成：{item.display_title}")
        self.index += 1
        self._refresh()

    def skip_current(self) -> None:
        item = self._current()
        if item:
            self.logger.log(f"[跳过] {item.display_title}")
        self.index += 1
        self._refresh()

    def _finish_all(self) -> None:
        self.progress.setValue(len(self.items))
        self.capture_btn.setEnabled(False)
        self.finish_btn.setEnabled(False)
        self.skip_btn.setEnabled(False)
        QMessageBox.information(self, "完成", "所有待执行用例已处理完毕。")
        self.accept()
