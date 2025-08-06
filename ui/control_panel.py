# ui/control_panel.py
import threading
from core.screenshot import Screenshot
from ui.annotator import Annotator
import tkinter as tk
from tkinter import messagebox

class ControlPanel(tk.Toplevel):
    def __init__(self, master, case_name, verify_point, wait_event):
        super().__init__(master)
        self.case_name = case_name
        self.verify_point = verify_point
        self.wait_event = wait_event

        self.screenshot_obj = Screenshot()
        self.screenshot_done = False
        self.screenshot_path = None
        self.skip_flag = False
        self.exit_flag = False

        self.title("控制面板")
        self.geometry("400x180")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_exit)

        self.create_widgets()
        self._init_bindings()

        self.wait_event.set()  # 通知创建完成

    def create_widgets(self):
        tk.Label(self, text=f"用例名：{self.case_name}", font=("Arial", 12)).pack(pady=5)
        tk.Label(self, text=f"验证点：{self.verify_point}", font=("Arial", 10), wraplength=380).pack(pady=5)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        self.btn_screenshot = tk.Button(btn_frame, text="截图 (F8)", width=10, command=self.on_screenshot)
        self.btn_screenshot.grid(row=0, column=0, padx=5)

        self.btn_complete = tk.Button(btn_frame, text="完成", width=10, command=self.on_complete)
        self.btn_complete.grid(row=0, column=1, padx=5)

        self.btn_skip = tk.Button(btn_frame, text="跳过", width=10, command=self.on_skip)
        self.btn_skip.grid(row=0, column=2, padx=5)

        self.btn_exit = tk.Button(self, text="退出", width=10, command=self.on_exit)
        self.btn_exit.pack(pady=5)

    def _init_bindings(self):
        self.bind_all("<F8>", lambda e: self.on_screenshot())

    def on_screenshot(self):
        filepath = self.screenshot_obj.take_screenshot(self.case_name)
        if not filepath:
            messagebox.showerror("错误", "截图失败！")
            return

        def save_callback(saved_path):
            self.screenshot_done = True
            self.screenshot_path = saved_path
            self.wait_event.set()  # 标注完成通知主流程

        def open_annotator():
            annotator = Annotator(self, filepath, save_callback)
            annotator.grab_set()
            annotator.focus_set()
            annotator.wait_window()

        # 在主线程调用标注窗口
        threading.Thread(target=lambda: self.master.after(0, open_annotator)).start()

    def on_complete(self):
        if not self.screenshot_done:
            if not messagebox.askyesno("确认", "未截图，是否确认完成？"):
                return
        self.wait_event.set()

    def on_skip(self):
        if messagebox.askyesno("跳过确认", "确认跳过当前用例？"):
            self.skip_flag = True
            self.wait_event.set()

    def on_exit(self):
        if messagebox.askyesno("退出确认", "确认退出当前测试流程？"):
            self.exit_flag = True
            self.wait_event.set()