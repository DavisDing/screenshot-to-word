import logging
from tkinter import scrolledtext, Toplevel, BOTH

class Logger:
    def __init__(self, log_path):
        self.log_path = log_path
        self.window = None
        self.text_widget = None
        self._setup_logger()

    def _setup_logger(self):
        logging.basicConfig(filename=self.log_path, level=logging.INFO, format='%(asctime)s - %(message)s')

    def create_window(self):
        self.window = Toplevel()
        self.window.title("执行日志")
        self.window.geometry("500x300")
        self.window.attributes("-topmost", True)
        self.text_widget = scrolledtext.ScrolledText(self.window, bg="black", fg="lime", font=("Courier", 10))
        self.text_widget.pack(fill=BOTH, expand=True)

    def log(self, message):
        logging.info(message)
        if self.text_widget:
            self.text_widget.insert("end", message + "\n")
            self.text_widget.see("end")