import os
import datetime
import tkinter as tk

class Logger:
    def __init__(self):
        self.text_widget = None
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_path = os.path.join(log_dir, f"log_{timestamp}.txt")

    def attach_text_widget(self, widget: tk.Text):
        self.text_widget = widget

    def log(self, message, level="info"):
        prefix = {"info": "[INFO]", "error": "[ERROR]"}.get(level, "[INFO]")
        text = f"{prefix} {message}\n"

        if self.text_widget:
            self.text_widget.insert(tk.END, text)
            self.text_widget.see(tk.END)

        try:
            with open(self.log_path, 'a', encoding='utf-8') as f:
                f.write(text)
        except Exception:
            pass