# main.py
import os
import tkinter as tk
from tkinter import messagebox
from core.test_runner import TestRunner
from utils.logger import Logger
from utils.excel_handler import ExcelHandler

class DesktopTestTool:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.base_dir)
        self.ensure_directories()

        self.logger = Logger()
        self.excel_handler = ExcelHandler(self.logger)
        self.test_runner = None

        self.root = tk.Tk()
        self.root.title("桌面自动化测试工具")
        self.root.geometry("300x150")
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

    def ensure_directories(self):
        for folder in ['excel_input', 'word_output', 'logs']:
            os.makedirs(folder, exist_ok=True)

    def create_widgets(self):
        tk.Button(self.root, text="开始执行", width=20, height=2, command=self.start_execution).pack(pady=15)
        tk.Button(self.root, text="退出", width=20, height=1, command=self.on_exit).pack()

    def start_execution(self):
        try:
            self.test_runner = TestRunner(self.logger, self.excel_handler, self.root)
            self.test_runner.run()
        except Exception as e:
            self.logger.log(f"启动失败: {str(e)}")
            messagebox.showerror("错误", f"启动执行失败：{e}")

    def on_exit(self):
        if messagebox.askyesno("退出确认", "是否确认退出？"):
            self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == '__main__':
    app = DesktopTestTool()
    app.run()