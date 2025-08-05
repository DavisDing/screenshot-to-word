import tkinter as tk

class ControlPanel(tk.Toplevel):
    def __init__(self, on_screenshot, on_skip):
        super().__init__()
        self.title("控制面板")
        self.attributes("-topmost", True)
        self.geometry("300x150")
        tk.Button(self, text="截图 (F8)", command=on_screenshot).pack(pady=5)
        tk.Button(self, text="跳过当前用例", command=on_skip).pack(pady=5)