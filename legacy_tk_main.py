import os
import threading
import tkinter as tk
from tkinter import ttk
import webbrowser

from core.test_runner import TestRunner
from ui.theme import create_app_root
from utils.excel_handler import ExcelHandler
from utils.logger import Logger
from utils.path_utils import ensure_directories, get_base_path

APP_TITLE = "SCREENSHOT-TO-WORD"
APP_SUBTITLE = "测试截图标注与 Word 报告生成工具"
GITHUB_URL = "https://github.com/DavisDing/screenshot-to-word"


class DesktopTestToolApp:
    def __init__(self):
        self.base_path = get_base_path()
        os.chdir(self.base_path)  # 设置工作目录为exe同级路径，确保所有相对路径正确

        self.theme = create_app_root(APP_TITLE, geometry="820x620", minsize=(720, 520))
        self.root = self.theme.root
        self.root.resizable(True, True)

        self.logger = Logger(self.root)

        try:
            ensure_directories(self.base_path, self.logger)
        except Exception as e:
            self.logger.log(f"初始化目录异常：{e}")

        self.excel_handler = ExcelHandler(self.logger, input_dir=os.path.join(self.base_path, "excel_input"), root=self.root)
        self.test_runner = TestRunner(self.logger, self.excel_handler, self.root)

        self.create_main_ui()
        theme_name = "ttkbootstrap/flatly" if self.theme.using_bootstrap else "ttk/clam fallback"
        self.logger.log(f"程序启动，界面加载完成，主题：{theme_name}")

    def create_main_ui(self):
        main_frame = ttk.Frame(self.root, padding=(24, 22), style="App.TFrame")
        main_frame.pack(side="top", fill="both", expand=False)
        main_frame.columnconfigure(0, weight=1)

        header = ttk.Frame(main_frame, padding=18, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=APP_TITLE, style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=APP_SUBTITLE, style="SubHeader.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(
            header,
            text="支持基础版 / 步骤版 Excel；列名无法自动识别时，可通过映射窗口手动确认。",
            style="Hint.TLabel",
            wraplength=700,
            justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(12, 0))

        action_card = ttk.Frame(main_frame, padding=18, style="Card.TFrame")
        action_card.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        action_card.columnconfigure(0, weight=1)
        action_card.columnconfigure(1, weight=1)
        action_card.columnconfigure(2, weight=1)

        ttk.Button(action_card, text="开始执行", command=self.on_start, style="Primary.TButton").grid(
            row=0, column=0, sticky="ew", padx=(0, 10)
        )
        ttk.Button(action_card, text="打开 GitHub", command=self.open_github, style="Secondary.TButton").grid(
            row=0, column=1, sticky="ew", padx=10
        )
        ttk.Button(action_card, text="退出", command=self.on_exit, style="Danger.TButton").grid(
            row=0, column=2, sticky="ew", padx=(10, 0)
        )

        tips = ttk.Labelframe(main_frame, text="使用提示", padding=14)
        tips.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        tips.columnconfigure(0, weight=1)
        tips_text = (
            "1. 将 Excel 用例放入 excel_input 目录，点击“开始执行”。\n"
            "2. 控制面板会展示当前用例、验证点和步骤信息，F8 可快速截图。\n"
            "3. 标注保存后，截图会自动写入 word_output 下的 Word 报告。"
        )
        ttk.Label(tips, text=tips_text, style="Body.TLabel", justify="left", wraplength=720).grid(row=0, column=0, sticky="w")

    def open_github(self):
        webbrowser.open_new(GITHUB_URL)

    def on_start(self):
        self.logger.log("点击开始执行按钮")
        thread = threading.Thread(target=self.test_runner.run_tests)
        thread.daemon = True
        thread.start()

    def on_exit(self):
        self.logger.log("点击退出按钮，程序结束")
        self.root.quit()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = DesktopTestToolApp()
    app.run()
