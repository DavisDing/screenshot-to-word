import tkinter as tk
import keyboard
import threading

class ControlPanel:
    def __init__(self, on_screenshot, on_skip, on_exit):
        self.on_screenshot = on_screenshot
        self.on_skip = on_skip
        self.on_exit = on_exit
        self.event = threading.Event()
        self.window = tk.Toplevel()
        self.window.title("控制面板")
        self.window.attributes("-topmost", True)
        self.window.geometry("350x200")

        self.label_case = tk.Label(self.window, text="用例：", font=("Arial", 12))
        self.label_case.pack(pady=5)

        self.label_checkpoint = tk.Label(self.window, text="验证点：", wraplength=300)
        self.label_checkpoint.pack(pady=5)

        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="截图(F8)", command=self.trigger_screenshot, width=10).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="完成", command=self.trigger_done, width=10).grid(row=0, column=3, padx=5)
        tk.Button(btn_frame, text="跳过", command=self.trigger_skip, width=10).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="退出", command=self.trigger_exit, width=10).grid(row=0, column=2, padx=5)

        keyboard.add_hotkey('F8', self.trigger_screenshot)

    def update_case(self, filename, checkpoint, progress_text=""):
        self.label_case.config(text=f"用例：{filename}  [{progress_text}]")
        self.label_checkpoint.config(text=f"验证点：{checkpoint}")

    def trigger_screenshot(self):
        self.on_screenshot()
        self.event.set()

    def trigger_skip(self):
        self.on_skip()
        self.event.set()

    def trigger_exit(self):
        self.on_exit()
        self.event.set()

    def wait_for_action(self):
        self.event.clear()
        self.event.wait()

    def trigger_done(self):
        self.event.set()

    def reset_action(self):
        self.event.set()

    def destroy(self):
        keyboard.unhook_all_hotkeys()
        self.window.destroy()