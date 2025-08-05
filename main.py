import os
import threading
import tkinter as tk
from tkinter import messagebox
from utils.excel_handler import ExcelHandler
from utils.logger import Logger
from ui.control_panel import ControlPanel
from core.test_runner import TestRunner

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("自动化测试工具")
        self.root.geometry("300x150")

        self.logger = Logger("logs/output.log")
        self.excel_handler = ExcelHandler("excel_input")
        self.test_runner = None

        self.status_label = tk.Label(root, text="准备就绪")
        self.status_label.pack(pady=10)

        self.start_button = tk.Button(root, text="开始执行", command=self.start_execution)
        self.start_button.pack(pady=5)

        self.exit_button = tk.Button(root, text="退出", command=root.quit)
        self.exit_button.pack(pady=5)

    def start_execution(self):
        self.root.withdraw()
        self.status_label.config(text="执行中...")
        cases = self.excel_handler.load_cases()
        if not cases:
            messagebox.showinfo("提示", "没有可执行的用例")
            self.root.deiconify()
            return

        self.test_runner = TestRunner(
            self.excel_handler, self.logger, self.status_label, self.root
        )
        threading.Thread(target=self.test_runner.run_tests, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
