from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QApplication,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from qt_app.models.test_case import TestItem
from qt_app.ui.annotator_window import AnnotatorWindow
from qt_app.ui.icons import app_icon

if TYPE_CHECKING:
    from qt_app.services.excel_service import ExcelService
    from qt_app.services.screenshot_service import ScreenshotService
    from qt_app.services.word_service import WordService


class ControlWindow(QDialog):
    """Small floating execution panel shown near the bottom-right corner."""

    PANEL_OPACITY = 0.88

    def __init__(
        self,
        items: List[TestItem],
        excel_service: ExcelService,
        screenshot_service: ScreenshotService,
        word_service: WordService,
        logger,
        parent=None,
        owner_window: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.items = items
        self.excel_service = excel_service
        self.screenshot_service = screenshot_service
        self.word_service = word_service
        self.logger = logger
        self.owner_window = owner_window
        self.index = 0
        self.last_image: Optional[Path] = None
        self.current_saved = False
        self.field_rows: Dict[str, Tuple[QWidget, QLabel, QLabel]] = {}
        icon = app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setWindowTitle("执行控制")
        self.resize(420, 440)
        self.setMinimumSize(380, 360)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowOpacity(self.PANEL_OPACITY)
        self._build_ui()
        self._refresh()
        self._move_to_bottom_right()
        QShortcut(QKeySequence("F8"), self, activated=self.capture_and_annotate)

    def showEvent(self, event):  # noqa: N802 - Qt API
        super().showEvent(event)
        self._move_to_bottom_right()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(9)

        header = QFrame()
        header.setObjectName("FloatingCard")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(7)
        title_row = QHBoxLayout()
        title = QLabel("执行控制")
        title.setObjectName("PanelTitle")
        self.progress_label = QLabel("-")
        self.progress_label.setObjectName("MutedLabel")
        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(self.progress_label)
        self.progress = QProgressBar()
        header_layout.addLayout(title_row)
        header_layout.addWidget(self.progress)
        root.addWidget(header)

        card = QFrame()
        card.setObjectName("FloatingCard")
        self.fields_layout = QVBoxLayout(card)
        self.fields_layout.setContentsMargins(12, 10, 12, 10)
        self.fields_layout.setSpacing(6)
        for key, name in [
            ("case", "用例"),
            ("checkpoint", "验证"),
            ("step", "步骤"),
            ("desc", "描述"),
            ("expected", "预期"),
        ]:
            self._add_field_row(key, name)
        root.addWidget(card, stretch=1)

        buttons1 = QHBoxLayout()
        self.capture_btn = QPushButton("截图并标注 F8")
        self.capture_btn.setObjectName("PrimaryButton")
        self.capture_btn.clicked.connect(self.capture_and_annotate)
        self.finish_btn = QPushButton("下一条")
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

        tip = QLabel("截图时主窗口和面板都会隐藏，完成标注后会自动写入 Word 并回写 Excel，再回到这里。")
        tip.setObjectName("PanelTip")
        tip.setWordWrap(True)
        root.addWidget(tip)

    def _add_field_row(self, key: str, name: str) -> None:
        row = QFrame()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        key_label = QLabel(name)
        key_label.setObjectName("PanelKey")
        key_label.setFixedWidth(36)
        key_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        value_label = QLabel("-")
        value_label.setObjectName("PanelValue")
        value_label.setWordWrap(True)
        value_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        value_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(key_label)
        layout.addWidget(value_label, stretch=1)
        self.fields_layout.addWidget(row)
        self.field_rows[key] = (row, key_label, value_label)

    def _set_field(self, key: str, value: str, required: bool = False) -> None:
        row, _key_label, value_label = self.field_rows[key]
        text = value.strip() if value else ""
        row.setVisible(required or bool(text))
        value_label.setText(text or "-")

    def _move_to_bottom_right(self) -> None:
        screen = self.screen()
        if not screen:
            return
        geo = screen.availableGeometry()
        self.move(geo.right() - self.width() - 24, geo.bottom() - self.height() - 36)

    def _keep_owner_hidden(self) -> None:
        # 执行期间主窗口不再露出，避免进入截图画面；只在整个执行面板结束后由 MainWindow 恢复。
        if self.owner_window is not None:
            self.owner_window.hide()

    def _restore_panel(self) -> None:
        self._keep_owner_hidden()
        self.setWindowOpacity(self.PANEL_OPACITY)
        self.showNormal()
        self.show()
        self.adjustSize()
        self._move_to_bottom_right()
        self.raise_()
        self.activateWindow()

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
        self.current_saved = False
        self._set_field("case", item.case_name, required=True)
        self._set_field("checkpoint", item.checkpoint, required=True)
        self._set_field("step", item.step_name)
        self._set_field("desc", item.step_description)
        self._set_field("expected", item.expected_result)
        self.finish_btn.setEnabled(False)
        self.adjustSize()
        self._move_to_bottom_right()

    def capture_and_annotate(self) -> None:
        item = self._current()
        if not item:
            return
        app = QApplication.instance()
        if app is not None:
            app.setQuitOnLastWindowClosed(False)
        self.logger.log(f"[截图] 开始截图：{item.display_title}")
        self._keep_owner_hidden()
        self.hide()
        # 给系统一点时间真正隐藏窗口，避免主程序或执行面板进入截图。
        QTimer.singleShot(450, lambda: self._do_capture(item))

    def _do_capture(self, item: TestItem) -> None:
        self._keep_owner_hidden()
        try:
            image_path = self.screenshot_service.capture_fullscreen(item.case_name)
        except Exception as exc:
            self.logger.log(f"[错误] 截图失败：{exc}")
            self._restore_panel()
            QMessageBox.critical(self, "截图失败", f"无法完成截图：{exc}")
            return
        self._keep_owner_hidden()
        if not image_path:
            self._restore_panel()
            QMessageBox.warning(self, "截图失败", "无法完成截图，请检查系统截图权限。")
            return

        try:
            annotator = AnnotatorWindow(image_path, None)
            result = annotator.exec()
            if result == QDialog.DialogCode.Accepted and annotator.result_path:
                self.last_image = annotator.result_path
                self.logger.log(f"[标注] 标注完成：{self.last_image}")
                if self._save_current_result(item, self.last_image):
                    self.finish_btn.setEnabled(True)
            else:
                self.logger.log("[标注] 用户取消标注")
        except Exception as exc:
            self.logger.log(f"[错误] 标注失败：{exc}")
            self._restore_panel()
            QMessageBox.critical(self, "标注失败", str(exc))
            return
        # 标注窗口关闭后，下一轮事件循环再恢复执行面板，避免 exec() 收尾阶段抢焦点失败。
        QTimer.singleShot(0, self._restore_panel)

    def _save_current_result(self, item: TestItem, image_path: Path) -> bool:
        try:
            self.word_service.add_image(item.case_name, item.checkpoint, image_path, item.step_note())
            self.excel_service.mark_executed([item.row_index])
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))
            self.logger.log(f"[错误] 保存当前用例失败：{exc}")
            self.current_saved = False
            return False
        self.current_saved = True
        self.logger.log(f"[完成] 当前用例已自动写入 Word 并回写 Excel：{item.display_title}")
        return True

    def finish_current(self) -> None:
        item = self._current()
        if not item or not self.last_image:
            QMessageBox.information(self, "提示", "请先截图并保存标注。")
            return
        if not self.current_saved:
            if not self._save_current_result(item, self.last_image):
                return
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
