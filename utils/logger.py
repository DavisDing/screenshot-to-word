import datetime
import os
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from utils.path_utils import get_base_path


class Logger:
    def __init__(self, root, log_file="logs/output.log"):
        self.root = root
        base_path = get_base_path()
        self.log_path = os.path.join(base_path, log_file)
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

        self.container = ttk.Labelframe(root, text="运行日志", padding=(10, 8))
        self.container.pack(side="bottom", fill="both", expand=True, padx=24, pady=(0, 18))
        self.container.columnconfigure(0, weight=1)
        self.container.rowconfigure(0, weight=1)

        self.text_area = ScrolledText(
            self.container,
            height=8,
            borderwidth=0,
            relief="flat",
            font=("Consolas", 10),
            wrap="word",
        )
        self.text_area.grid(row=0, column=0, sticky="nsew")
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
