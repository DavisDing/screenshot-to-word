# ui/control_panel.py
import tkinter as tk

class ControlPanel:
    def __init__(self, on_capture, on_next, on_skip, on_quit):
        self.window = tk.Toplevel()
        self.window.title("测试控制面板")
        self.window.geometry("400x200")
        self.window.attributes('-topmost', True)
        self.window.resizable(False, False)

        self.on_capture = on_capture
        self.on_next = on_next
        self.on_skip = on_skip
        self.on_quit = on_quit

        self.case_label = tk.Label(self.window, text="用例：", font=("微软雅黑", 12), wraplength=380, justify="left")
        self.case_label.pack(pady=10)

        self.capture_btn = tk.Button(self.window, text="截图 (F8)", width=20, height=2, command=self.on_capture)
        self.capture_btn.pack(pady=5)

        self.next_btn = tk.Button(self.window, text="完成当前用例", width=20, height=2, command=self.on_next)
        self.next_btn.pack(pady=5)

        self.skip_btn = tk.Button(self.window, text="跳过", width=10, command=self.on_skip)
        self.skip_btn.pack(side="left", padx=20, pady=10)

        self.quit_btn = tk.Button(self.window, text="退出", width=10, command=self.on_quit)
        self.quit_btn.pack(side="right", padx=20, pady=10)

        # 绑定快捷键
        self.window.bind_all("<F8>", lambda e: self.on_capture())

    def update_case(self, case_name, case_desc):
        text = f"用例名：{case_name}\n验证点：{case_desc}"
        self.case_label.config(text=text)

    def destroy(self):
        self.window.destroy()