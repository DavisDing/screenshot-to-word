from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QListWidget,
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
        self.setWindowTitle("Screenshot To Word")
        self.resize(1080, 720)
        self.setMinimumSize(920, 620)
        self._build_ui()
        self.logger.set_sink(self._append_log)
        self.logger.log("程序启动完成：PySide6 / Qt 重构版")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")
        self.setCentralWidget(root)
        main = QHBoxLayout(root)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(260)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(24, 26, 24, 24)
        side.setSpacing(14)
        brand = QLabel("Screenshot\nTo Word")
        brand.setObjectName("BrandTitle")
        subtitle = QLabel("跨平台测试截图报告工具")
        subtitle.setObjectName("BrandSubTitle")
        subtitle.setWordWrap(True)
        side.addWidget(brand)
        side.addWidget(subtitle)
        side.addSpacing(18)
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
        version = QLabel("Qt UI · Excel 智能映射 · 高级标注")
        version.setObjectName("BrandSubTitle")
        version.setWordWrap(True)
        side.addWidget(version)
        main.addWidget(sidebar)

        content = QVBoxLayout()
        content.setContentsMargins(28, 26, 28, 24)
        content.setSpacing(18)
        main.addLayout(content, stretch=1)

        title = QLabel("专业化截图标注与 Word 报告生成")
        title.setObjectName("PageTitle")
        content.addWidget(title)

        subtitle2 = QLabel("适配 Windows / macOS，支持 Excel 列智能识别、手动列含义确认、截图裁剪、方形/圆形标记、荧光笔、文字标注。")
        subtitle2.setObjectName("MutedLabel")
        subtitle2.setWordWrap(True)
        content.addWidget(subtitle2)

        stat_row = QHBoxLayout()
        stat_row.setSpacing(14)
        stat_row.addWidget(self._stat_card("输入目录", str(self.base_path / "excel_input")))
        stat_row.addWidget(self._stat_card("输出目录", str(self.base_path / "word_output")))
        content.addLayout(stat_row)

        action_card = QFrame()
        action_card.setObjectName("Card")
        action = QGridLayout(action_card)
        action.setContentsMargins(22, 22, 22, 22)
        action.setHorizontalSpacing(14)
        action_title = QLabel("开始")
        action_title.setObjectName("SectionTitle")
        action_desc = QLabel("选择 Excel 后，系统会先预览表头并确认列映射，再进入执行控制面板。")
        action_desc.setObjectName("MutedLabel")
        action_desc.setWordWrap(True)
        self.start_btn = QPushButton("选择 Excel 并开始")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.clicked.connect(self.select_and_run)
        action.addWidget(action_title, 0, 0)
        action.addWidget(action_desc, 1, 0)
        action.addWidget(self.start_btn, 0, 1, 2, 1)
        action.setColumnStretch(0, 1)
        content.addWidget(action_card)

        feature_card = QFrame()
        feature_card.setObjectName("Card")
        feature = QVBoxLayout(feature_card)
        feature.setContentsMargins(22, 18, 22, 18)
        ftitle = QLabel("本次重构重点")
        ftitle.setObjectName("SectionTitle")
        feature.addWidget(ftitle)
        features = QListWidget()
        features.addItems([
            "现代化 Qt 界面：卡片布局、侧边栏、统一视觉风格",
            "Excel 智能列识别：列名不标准时展示样例值并让用户确认每列含义",
            "截图标注增强：裁剪、方形标记、圆形标记、荧光笔、画笔、文字、撤销",
            "跨平台稳定性：移除 Tk/keyboard/pyautogui 依赖，改用 Qt 事件和屏幕 API",
            "可诊断：启动异常会写入 logs/crash.log，便于定位打包闪退问题",
        ])
        feature.addWidget(features)
        content.addWidget(feature_card, stretch=1)

        log_title = QLabel("运行日志")
        log_title.setObjectName("SectionTitle")
        content.addWidget(log_title)
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(120)
        content.addWidget(self.log_view)

    def _stat_card(self, name: str, value: str) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 16)
        label = QLabel(name)
        label.setObjectName("MutedLabel")
        value_label = QLabel(value)
        value_label.setWordWrap(True)
        layout.addWidget(label)
        layout.addWidget(value_label)
        return card

    def _append_log(self, line: str) -> None:
        self.log_view.append(line)

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
            self.logger.log(f"读取 Excel 失败：{exc}")
            return

        dialog = ExcelMappingDialog(preview, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.logger.log("用户取消 Excel 列映射")
            return
        try:
            items = self.excel_service.load(Path(file_name), dialog.mapping())
        except Exception as exc:
            QMessageBox.critical(self, "加载失败", str(exc))
            self.logger.log(f"加载 Excel 失败：{exc}")
            return
        if not items:
            QMessageBox.information(self, "无需执行", "所有用例均已执行，或 Excel 中没有待执行数据。")
            self.logger.log("没有待执行用例")
            return

        panel = ControlWindow(items, self.excel_service, self.screenshot_service, self.word_service, self.logger, self)
        panel.exec()

    def open_dir(self, path: Path) -> None:
        path.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
