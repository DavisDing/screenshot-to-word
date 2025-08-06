# main.py
import os
import threading
import tkinter as tk
from tkinter import messagebox
from utils.logger import Logger
from utils.excel_handler import ExcelHandler
from core.test_runner import TestRunner

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 创建必要目录
for folder in ['logs', 'excel_input', 'word_output']:
    path = os.path.join(BASE_DIR, folder)
    os.makedirs(path, exist_ok=True)

logger = Logger()
excel_handler = ExcelHandler(logger)

def start_execution():
    runner = TestRunner(logger, excel_handler)
    threading.Thread(target=runner.run, daemon=True).start()

def exit_program():
    root.quit()

# 主窗口
root = tk.Tk()
root.title("桌面自动化测试工具")
root.geometry("400x200")

start_btn = tk.Button(root, text="开始执行", command=start_execution, font=("Arial", 14))
start_btn.pack(pady=20)

exit_btn = tk.Button(root, text="退出", command=exit_program, font=("Arial", 14))
exit_btn.pack(pady=20)

log_box = tk.Text(root, height=6)
log_box.pack(fill="both", expand=True)
logger.attach_text_widget(log_box)

root.mainloop()