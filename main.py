import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox

from utils.logger import Logger
from utils.excel_handler import ExcelHandler
from core.test_runner import TestRunner

def get_base_dir():
    # 运行时获取exe所在目录，打包后生效
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
        self.root.title("测试截图小工具")
        self.root.geometry("700x500")

        self.logger = Logger(LOG_DIR)
        self.logger.init_ui(self.root)

        self.excel_handler = ExcelHandler(EXCEL_DIR, self.logger)
        self.test_runner = TestRunner(
            root=self.root,
            excel_handler=self.excel_handler,
            logger=self.logger,
            word_output_dir=WORD_OUTPUT_DIR,
            temp_dir=TEMP_DIR
        )

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        self.start_btn = tk.Button(btn_frame, text="开始执行", width=15, height=2, command=self.start_test)
        self.start_btn.pack(side="left", padx=10)

        self.exit_btn = tk.Button(btn_frame, text="退出", width=15, height=2, command=self.on_exit)
        self.exit_btn.pack(side="left", padx=10)

        self.root.bind('<Control-s>', lambda event: self.start_test())
        self.root.bind('<Control-q>', lambda event: self.on_exit())

    def start_test(self):
        self.start_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.test_runner.run, daemon=True).start()

    def on_exit(self):
        if messagebox.askokcancel("退出确认", "确定退出程序？"):
            self.logger.close()
            self.root.destroy()

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()