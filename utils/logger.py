# utils/logger.py
import os
import tkinter as tk
import datetime

class Logger:
    def __init__(self):
        self.log_file_path = os.path.join("logs", "output.log")
        self.text_widget = None

    def bind_text_widget(self, text_widget):
        self.text_widget = text_widget

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        full_message = f"{timestamp} {message}\n"

        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(full_message)
        except PermissionError:
            pass  # 忽略日志写入失败

        print(full_message, end="")
        if self.text_widget:
            self.text_widget.insert(tk.END, full_message)
            self.text_widget.see(tk.END)
