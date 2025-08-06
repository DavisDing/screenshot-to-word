import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

from utils.logger import Logger
from utils.excel_handler import ExcelHandler
from core.test_runner import TestRunner

BG_COLOR = "#282c34"
FG_COLOR = "#abb2bf"
BTN_BG = "#61afef"
BTN_FG = "#282c34"
FONT = ("Segoe UI", 11)

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
EXCEL_DIR = os.path.join(BASE_DIR, "excel_input")
WORD_OUTPUT_DIR = os.path.join(BASE_DIR, "word_output")
LOG_DIR = os.path.join(BASE_DIR, "logs")
TEMP_DIR = os.path.join(BASE_DIR, "Temp")

for d in [EXCEL_DIR, WORD_OUTPUT_DIR, LOG_DIR, TEMP_DIR]:
    os.makedirs(d, exist_ok=True)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("桌面自动化测试工具")
        self.root.geometry("700x500")
        self.root.configure(bg=BG_COLOR)

        style = ttk.Style(root)
        style.theme_use('clam')
        style.configure('TButton', background=BTN_BG, foreground=BTN_FG, font=FONT)
        style.map('TButton',
                  background=[('active', '#528bff')],
                  foreground=[('active', '#ffffff')])

        self.logger = Logger(LOG_DIR)
        self.logger.init_ui(self.root)

        # 修改ExcelHandler初始化方式，只传递需要的参数
        self.excel_handler = ExcelHandler(excel_dir=EXCEL_DIR, logger=self.logger)
        self.test_runner = TestRunner(
            root=self.root,
            excel_handler=self.excel_handler,
            logger=self.logger,
            word_output_dir=WORD_OUTPUT_DIR,
            temp_dir=TEMP_DIR
        )

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10, fill='x', padx=20)

        self.start_btn = ttk.Button(btn_frame, text="开始执行", command=self.start_test)
        self.start_btn.pack(side="left", padx=10, ipadx=15, ipady=5)

        self.exit_btn = ttk.Button(btn_frame, text="退出", command=self.on_exit)
        self.exit_btn.pack(side="left", padx=10, ipadx=15, ipady=5)

    def start_test(self):
        if not self.excel_handler.load_excel():
            return
        self.start_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.test_runner.run, daemon=True).start()

    def on_exit(self):
        if messagebox.askokcancel("退出确认", "确定退出程序？"):
            self.root.destroy()

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()