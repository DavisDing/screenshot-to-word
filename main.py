import threading
import tkinter as tk
from utils.path_utils import get_base_path, ensure_directories
from utils.logger import Logger
from utils.excel_handler import ExcelHandler
from core.test_runner import TestRunner

APP_TITLE = "测试工具 - V1.0.0 作者: Haoding"

class DesktopTestToolApp:
    def __init__(self):
        import os
        os.chdir(get_base_path())  # 设置工作目录为exe同级路径，确保所有相对路径正确

        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("400x200")
        self.root.resizable(False, False)

        self.logger = Logger(self.root)

        try:
            self.base_path = get_base_path(self.logger)
            ensure_directories(self.base_path, self.logger)
        except Exception as e:
            self.logger.log(f"初始化目录异常：{e}")

        self.excel_handler = ExcelHandler(self.logger, input_dir=self.base_path + "/excel_input", root=self.root)
        self.test_runner = TestRunner(self.logger, self.excel_handler, self.root)

        self.create_main_ui()
        self.logger.log("程序启动，界面加载完成")

    def create_main_ui(self):
        tk.Label(self.root, text=APP_TITLE, font=("Arial", 16)).pack(pady=10)

        start_btn = tk.Button(self.root, text="开始执行", width=20, command=self.on_start)
        start_btn.pack(pady=5)

        exit_btn = tk.Button(self.root, text="退出", width=20, command=self.on_exit)
        exit_btn.pack(pady=5)

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