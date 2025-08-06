import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import datetime
import os
import sys
from utils.path_utils import get_base_path

class Logger:
    def __init__(self, root, log_file="logs/output.log"):
        self.root = root
        base_path = get_base_path()
        self.log_path = os.path.join(base_path, log_file)
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

        self.text_area = ScrolledText(root, height=6)
        self.text_area.pack(fill="x", padx=10, pady=10)
        self.text_area.configure(state="disabled")

    def log(self, msg):
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        full_msg = f"[{timestamp}] {msg}"

        self.text_area.configure(state="normal")
        self.text_area.insert(tk.END, full_msg + '\n')
        self.text_area.configure(state="disabled")
        self.text_area.yview(tk.END)

        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(full_msg + "\n")
        except Exception:
            # 避免写文件异常阻断程序，忽略
            pass