import tkinter as tk
from tkinter import ttk

BG_COLOR = "#282c34"
FG_COLOR = "#abb2bf"
BTN_BG = "#61afef"
BTN_FG = "#282c34"
FONT = ("Segoe UI", 11)

class ControlPanel:
    def __init__(self, on_capture, on_next, on_skip, on_quit):
        self.window = tk.Toplevel()
        self.window.title("测试控制面板")
        self.window.geometry("450x220")
        self.window.configure(bg=BG_COLOR)
        self.window.attributes('-topmost', True)
        self.window.resizable(False, False)

        style = ttk.Style(self.window)
        style.theme_use('clam')
        style.configure('TButton', background=BTN_BG, foreground=BTN_FG, font=FONT)
        style.map('TButton',
                  background=[('active', '#528bff')],
                  foreground=[('active', '#ffffff')])

        self.on_capture = on_capture
        self.on_next = on_next
        self.on_skip = on_skip
        self.on_quit = on_quit

        self.case_label = ttk.Label(self.window, text="用例：", font=(FONT[0], 12), background=BG_COLOR, foreground=FG_COLOR, wraplength=420, justify="left")
        self.case_label.pack(pady=15)

        self.capture_btn = ttk.Button(self.window, text="截图 (F8)", command=self.on_capture)
        self.capture_btn.pack(pady=8, ipadx=15, ipady=5)

        self.next_btn = ttk.Button(self.window, text="完成当前用例", command=self.on_next)
        self.next_btn.pack(pady=8, ipadx=15, ipady=5)

        btn_frame = ttk.Frame(self.window, style='TFrame', padding=10)
        btn_frame.pack(fill='x', pady=10)

        self.skip_btn = ttk.Button(btn_frame, text="跳过", command=self.on_skip)
        self.skip_btn.pack(side="left", expand=True, ipadx=15, ipady=5, padx=10)

        self.quit_btn = ttk.Button(btn_frame, text="退出", command=self.on_quit)
        self.quit_btn.pack(side="right", expand=True, ipadx=15, ipady=5, padx=10)

        # 绑定快捷键
        self.window.bind_all("<F8>", lambda e: self.on_capture())

    def update_case(self, case_name, case_desc):
        text = f"用例名：{case_name}\n验证点：{case_desc}"
        self.case_label.config(text=text)

    def destroy(self):
        self.window.destroy()