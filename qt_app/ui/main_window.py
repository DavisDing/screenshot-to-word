from __future__ import annotations

from pathlib import Path
import html

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from qt_app.services.excel_service import ExcelService
from qt_app.services.logger import AppLogger
from qt_app.services.paths import ensure_app_dirs
from qt_app.services.screenshot_service import ScreenshotService
from qt_app.services.word_service import WordService
from qt_app.ui.control_window import ControlWindow
from qt_app.ui.excel_mapping_dialog import ExcelMappingDialog
from qt_app.ui.icons import app_icon

GITHUB_URL = "https://github.com/DavisDing/screenshot-to-word"


class MainWindow(QMainWindow):
    def __init__(self, base_path: Path, logger: AppLogger):
        super().__init__()
        self.base_path = Path(base_path)
        ensure_app_dirs(self.base_path)
        self.logger = logger
        self.excel_service = ExcelService(self.logger)
        self.screenshot_service = ScreenshotService(self.base_path / "Temp", self.logger)
        self.word_service = WordService(self.base_path / "word_output", self.logger)
        icon = app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        self.setWindowTitle("Screenshot To Word")
        self.resize(1020, 680)
        self.setMinimumSize(860, 560)
        self._build_ui()
        self.logger.set_sink(self._append_log)
        self.logger.log("[启动] 程序启动完成：PySide6 / Qt 重构版")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")
        self.setCentralWidget(root)
        main = QHBoxLayout(root)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(232)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(22, 24, 22, 22)
        side.setSpacing(12)
        brand = QLabel("Screenshot\nTo Word")
        brand.setObjectName("BrandTitle")
        subtitle = QLabel("测试截图报告工具")
        subtitle.setObjectName("BrandSubTitle")
        subtitle.setWordWrap(True)
        side.addWidget(brand)
        side.addWidget(subtitle)
        side.addSpacing(16)
        start_nav = QPushButton("  开始执行")
        start_nav.setObjectName("GhostButton")
        start_nav.clicked.connect(self.select_and_run)
        out_nav = QPushButton("  打开输出目录")
        out_nav.setObjectName("GhostButton")
        out_nav.clicked.connect(lambda: self.open_dir(self.base_path / "word_output"))
        excel_nav = QPushButton("  打开 Excel 目录")
        excel_nav.setObjectName("GhostButton")
        excel_nav.clicked.connect(lambda: self.open_dir(self.base_path / "excel_input"))
        git_nav = QPushButton("  GitHub")
        git_nav.setObjectName("GhostButton")
        git_nav.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(GITHUB_URL)))
        side.addWidget(start_nav)
        side.addWidget(out_nav)
        side.addWidget(excel_nav)
        side.addWidget(git_nav)
        side.addStretch(1)
        version = QLabel("Excel 映射 · 截图标注 · Word 输出")
        version.setObjectName("BrandSubTitle")
        version.setWordWrap(True)
        side.addWidget(version)
        main.addWidget(sidebar)

        content = QVBoxLayout()
        content.setContentsMargins(24, 22, 24, 22)
        content.setSpacing(14)
        main.addLayout(content, stretch=1)

        title = QLabel("截图标注与 Word 报告生成")
        title.setObjectName("PageTitle")
        content.addWidget(title)

        subtitle2 = QLabel("选择 Excel 用例后，确认列含义，进入小型悬浮控制面板执行截图和标注。")
        subtitle2.setObjectName("MutedLabel")
        subtitle2.setWordWrap(True)
        content.addWidget(subtitle2)

        action_card = QFrame()
        action_card.setObjectName("Card")
        action = QGridLayout(action_card)
        action.setContentsMargins(20, 18, 20, 18)
        action.setHorizontalSpacing(14)
        action_title = QLabel("开始执行")
        action_title.setObjectName("SectionTitle")
        action_desc = QLabel("支持列名不标准的 Excel：系统会显示样例值并让你确认每一列是什么。")
        action_desc.setObjectName("MutedLabel")
        action_desc.setWordWrap(True)
        self.start_btn = QPushButton("选择 Excel")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.clicked.connect(self.select_and_run)
        action.addWidget(action_title, 0, 0)
        action.addWidget(action_desc, 1, 0)
        action.addWidget(self.start_btn, 0, 1, 2, 1)
        action.setColumnStretch(0, 1)
        content.addWidget(action_card)

        stat_row = QHBoxLayout()
        stat_row.setSpacing(12)
        stat_row.addWidget(self._stat_card("Excel 输入", str(self.base_path / "excel_input")))
        stat_row.addWidget(self._stat_card("Word 输出", str(self.base_path / "word_output")))
        content.addLayout(stat_row)

        quick_card = QFrame()
        quick_card.setObjectName("Card")
        quick = QGridLayout(quick_card)
        quick.setContentsMargins(18, 16, 18, 16)
        quick.setHorizontalSpacing(10)
        quick.setVerticalSpacing(6)
        flow_title = QLabel("操作流程")
        flow_title.setObjectName("SectionTitle")
        flow = QLabel("1. 选择 Excel  →  2. 确认列含义  →  3. 悬浮面板执行截图  →  4. 标注保存  →  5. 自动写入 Word 并回写 Excel")
        flow.setObjectName("MutedLabel")
        flow.setWordWrap(True)
        quick.addWidget(flow_title, 0, 0)
        quick.addWidget(flow, 1, 0)
        content.addWidget(quick_card)

        log_title = QLabel("运行日志")
        log_title.setObjectName("SectionTitle")
        content.addWidget(log_title)
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(190)
        content.addWidget(self.log_view, stretch=1)

    def _stat_card(self, name: str, value: str) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(6)
        label = QLabel(name)
        label.setObjectName("MutedLabel")
        value_label = QLabel(value)
        value_label.setObjectName("PathLabel")
        value_label.setWordWrap(True)
        layout.addWidget(label)
        layout.addWidget(value_label)
        return card

    def _append_log(self, line: str) -> None:
        color = self._log_color(line)
        safe = html.escape(line)
        self.log_view.append(f'<div style="margin:2px 0; color:{color};">{safe}</div>')

    def _log_color(self, line: str) -> str:
        if any(k in line for k in ("失败", "异常", "错误", "[错误]")):
            return "#dc2626"
        if any(k in line for k in ("截图", "标注", "Excel", "Word", "完成", "保存")):
            return "#1e40af"
        return "#475569"

    def select_and_run(self) -> None:
        excel_dir = self.base_path / "excel_input"
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 用例文件",
            str(excel_dir),
            "Excel 文件 (*.xlsx *.xlsm);;所有文件 (*.*)",
        )
        if not file_name:
            self.logger.log("用户取消选择 Excel")
            return
        try:
            preview = self.excel_service.preview(Path(file_name))
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", f"读取 Excel 失败：{exc}")
            self.logger.log(f"[错误] 读取 Excel 失败：{exc}")
            return

        dialog = ExcelMappingDialog(preview, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.logger.log("用户取消 Excel 列映射")
            return
        try:
            items = self.excel_service.load(Path(file_name), dialog.mapping())
        except Exception as exc:
            QMessageBox.critical(self, "加载失败", str(exc))
            self.logger.log(f"[错误] 加载 Excel 失败：{exc}")
            return
        if not items:
            QMessageBox.information(self, "无需执行", "所有用例均已执行，或 Excel 中没有待执行数据。")
            self.logger.log("[完成] 没有待执行用例")
            return

        self.hide()
        try:
            panel = ControlWindow(
                items,
                self.excel_service,
                self.screenshot_service,
                self.word_service,
                self.logger,
                parent=None,
                owner_window=self,
            )
            panel.exec()
        finally:
            self.setWindowOpacity(1.0)
            self.show()
            self.raise_()
            self.activateWindow()

    def open_dir(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
