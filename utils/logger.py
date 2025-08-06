import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import threading
import datetime
import os

class Logger:
    def __init__(self, log_dir):
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        self.log_file_path = os.path.join(log_dir, "output.log")

        self.lock = threading.Lock()
        self.log_queue = []

        self.root = None
        self.text_widget = None

    def init_ui(self, parent):
        self.text_widget = ScrolledText(parent, state='disabled', height=15, width=80, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 10))
        self.text_widget.pack(fill='both', expand=True)
        self.root = parent
        self._start_update_loop()

    def _start_update_loop(self):
        def update():
            with self.lock:
                while self.log_queue:
                    msg = self.log_queue.pop(0)
                    self._append_text(msg)
            if self.root:
                self.root.after(200, update)
        update()

    def _append_text(self, msg):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + "\n")
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def write(self, msg):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        full_msg = f"[{timestamp}] {msg}"
        with self.lock:
            self.log_queue.append(full_msg)
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(full_msg + "\n")
        except Exception:
            pass