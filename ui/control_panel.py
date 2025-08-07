# ui/control_panel.py

import tkinter as tk
from tkinter import messagebox
import threading
import time
from core.screenshot import ScreenshotTool
from utils.word_generator import WordGenerator

class ControlPanel(tk.Toplevel):
    def __init__(self, logger, pending_cases, excel_handler, root):
        super().__init__(root)
        self.logger = logger
        self.pending_cases = pending_cases
        self.excel_handler = excel_handler
        self.root = root

        self.title("控制面板")
        self.geometry("450x300")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.resizable(True, True)

        # 状态变量
        self.current_index = 0
        self.current_case = None
        self.screenshot_tool = ScreenshotTool(self.logger, root)
        self.word_generator = WordGenerator(self.logger, root)

        # 事件控制，截图完成通知
        self.screenshot_done_event = threading.Event()

        self.create_widgets()
        self.load_case()

    def create_widgets(self):
        self.lbl_case_name = tk.Label(self, text="用例名: ", wraplength=380, justify="left")
        self.lbl_case_name.pack(pady=5, fill="x")

        self.lbl_checkpoint = tk.Label(self, text="验证点: ", wraplength=380, justify="left")
        self.lbl_checkpoint.pack(pady=5, fill="x")

        self.lbl_progress = tk.Label(self, text="当前进度：", font=("Arial", 10), wraplength=380, justify="left")
        self.lbl_progress.pack(pady=5, fill="x")

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10, fill="x", expand=True, anchor="center")

        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)

        self.btn_screenshot = tk.Button(btn_frame, text="截图 (F8)", command=self.on_screenshot)
        self.btn_screenshot.grid(row=0, column=0, padx=5)
        self.bind_all("<F8>", lambda event: self.on_screenshot())

        self.btn_complete = tk.Button(btn_frame, text="完成", command=self.on_complete)
        self.btn_complete.grid(row=0, column=1, padx=5)

        self.btn_skip = tk.Button(btn_frame, text="跳过", command=self.on_skip)
        self.btn_skip.grid(row=0, column=2, padx=5)

    def load_case(self):
        if self.current_index >= len(self.pending_cases):
            self.finish_all_cases()
            return

        idx, filename, checkpoint = self.pending_cases[self.current_index]
        self.current_case = (idx, filename, checkpoint)
        self.lbl_case_name.config(text=f"用例名: {filename}")
        self.lbl_checkpoint.config(text=f"验证点: {checkpoint}")
        self.lbl_progress.config(
            text=f"当前进度：第 {self.current_index + 1} 条 / 共 {len(self.pending_cases)} 条"
        )
        self.logger.log(f"当前执行用例：{filename} - 验证点：{checkpoint}")

    def on_screenshot(self):
        # 截图 + 标注 + Word生成流程
        def run_screenshot_flow():
            idx, filename, checkpoint = self.current_case
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            file_name = f"{filename}_{timestamp}.png"

            self.logger.log(f"开始截图: {file_name}")

            # 截图
            img_path = self.screenshot_tool.capture_screen(file_name)

            if not img_path:
                self.logger.log("截图失败或取消")
                return

            # 标注，阻塞直到完成
            annotated_path = self.screenshot_tool.annotate(img_path)
            if not annotated_path:
                self.logger.log("标注取消或失败")
                return

            # 生成Word文档或追加图片
            self.word_generator.add_image_to_word(filename, checkpoint, annotated_path)

            # 标记已执行
            self.excel_handler.mark_case_executed(idx)
            self.logger.log(f"用例 {filename} 标记为已执行")

            self.screenshot_done_event.set()

        self.screenshot_done_event.clear()
        threading.Thread(target=run_screenshot_flow, daemon=True).start()

    def on_complete(self):
        if not self.screenshot_done_event.is_set():
            self._show_warning("请先截图并完成标注后，再点击完成。")
            return
        self.current_index += 1
        self.load_case()
        self.screenshot_done_event.clear()

    def on_skip(self):
        self.logger.log(f"跳过用例 {self.current_case[1]}")
        self.current_index += 1
        self.load_case()
        self.screenshot_done_event.clear()

    def finish_all_cases(self):
        self.root.attributes('-topmost', True)
        messagebox.showinfo("完成", "所有用例已执行完毕", parent=self.root)
        self.root.attributes('-topmost', False)
        self.destroy()
        self.root.deiconify()

    def on_exit(self):
        self.root.attributes('-topmost', True)
        if messagebox.askokcancel("退出", "确定退出测试？", parent=self.root):
            self.root.attributes('-topmost', False)
            self.root.destroy()
            self.destroy()
        else:
            self.root.attributes('-topmost', False)

    def _show_warning(self, msg):
        self.root.attributes('-topmost', True)
        messagebox.showwarning("警告", msg, parent=self.root)
        self.root.attributes('-topmost', False)