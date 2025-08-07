# ui/control_panel.py

import tkinter as tk
from tkinter import messagebox
import threading
import time
from core.screenshot import ScreenshotTool
from utils.word_generator import WordGenerator

class ControlPanel(tk.Toplevel):
    def __init__(self, logger, pending_cases, excel_handler, root, is_step_mode=False):
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

        self.is_step_mode = is_step_mode
        if self.is_step_mode:
            self.step_case_keys = list(pending_cases.keys())
            self.current_case_key_index = 0
            self.current_step_index = 0
            self.pending_cases = pending_cases
            self.current_case_steps = []

        self.create_widgets()
        self.load_case()
        self.bind("<Configure>", self._on_resize)

    def create_widgets(self):
        info_frame = tk.Frame(self)
        info_frame.pack(pady=5, fill="both", expand=True)

        def add_row(label_text, attr_name):
            label = tk.Label(info_frame, text=label_text, anchor="w", font=("Arial", 10, "bold"))
            label.grid(row=add_row.row_index, column=0, sticky="nw", padx=5, pady=2)
            value_label = tk.Label(info_frame, text="", anchor="w", wraplength=360, justify="left")
            value_label.grid(row=add_row.row_index, column=1, sticky="nw", padx=5, pady=2)
            setattr(self, attr_name, value_label)
            add_row.row_index += 1

        add_row.row_index = 0
        add_row("用例名：", "lbl_case_name")
        add_row("验证点：", "lbl_checkpoint")
        add_row("步骤名称：", "lbl_step_name")
        add_row("步骤描述：", "lbl_step_desc")
        add_row("预期结果：", "lbl_expected")
        add_row("当前进度：", "lbl_progress")

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10, fill="x", expand=True, anchor="center")

        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)
        btn_frame.grid_columnconfigure(3, weight=1)

        self.btn_screenshot = tk.Button(btn_frame, text="截图 (F8)", command=self.on_screenshot)
        self.btn_screenshot.grid(row=0, column=0, padx=5)
        self.bind_all("<F8>", lambda event: self.on_screenshot())

        self.btn_complete = tk.Button(btn_frame, text="完成", command=self.on_complete)
        self.btn_complete.grid(row=0, column=1, padx=5)

        self.btn_skip = tk.Button(btn_frame, text="跳过", command=self.on_skip)
        self.btn_skip.grid(row=0, column=2, padx=5)

        self.btn_next = tk.Button(btn_frame, text="下一步", command=self.on_next_step)
        self.btn_next.grid(row=0, column=3, padx=5)

    def load_case(self):
        if self.is_step_mode:
            if self.current_case_key_index >= len(self.step_case_keys):
                self.finish_all_cases()
                return

            case_key = self.step_case_keys[self.current_case_key_index]
            self.current_case_steps = self.pending_cases[case_key]
            if self.current_step_index >= len(self.current_case_steps):
                self.btn_complete.config(state="normal")
                return

            step = self.current_case_steps[self.current_step_index]
            self.current_case = (step["index"], case_key[0], case_key[1])
            self.lbl_case_name.config(text=f"用例名: {case_key[0]}")
            self.lbl_checkpoint.config(text=f"验证点: {case_key[1]}")
            self.lbl_step_name.config(text=f"步骤名称: {step.get('步骤名称', '')}")
            self.lbl_step_desc.config(text=f"步骤描述: {step.get('步骤描述', '')}")
            self.lbl_expected.config(text=f"预期结果: {step.get('预期结果', '')}")
            remaining_cases = len(self.step_case_keys) - self.current_case_key_index - 1
            self.lbl_progress.config(
                text=f"当前进度：第 {self.current_step_index + 1} 步 / 共 {len(self.current_case_steps)} 步，剩余 {remaining_cases} 条案例"
            )
            self.btn_complete.config(state="disabled")
            # 显示步骤相关标签
            self.lbl_step_name.grid()
            self.lbl_step_desc.grid()
            self.lbl_expected.grid()
            return

        # 基础版处理
        if self.current_index >= len(self.pending_cases):
            self.finish_all_cases()
            return

        idx, filename, checkpoint = self.pending_cases[self.current_index]
        self.current_case = (idx, filename, checkpoint)
        self.lbl_case_name.config(text=f"用例名: {filename}")
        self.lbl_checkpoint.config(text=f"验证点: {checkpoint}")
        # 隐藏步骤相关标签
        self.lbl_step_name.config(text="")
        self.lbl_step_desc.config(text="")
        self.lbl_expected.config(text="")
        self.lbl_step_name.grid_remove()
        self.lbl_step_desc.grid_remove()
        self.lbl_expected.grid_remove()
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

            # 支持步骤文字
            step_note = ""
            if self.is_step_mode:
                step = self.current_case_steps[self.current_step_index]
                step_note = f"{step['步骤名称']} - {step['步骤描述']}"
            self.word_generator.add_image_to_word(filename, checkpoint, annotated_path, step_note)

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
        if self.is_step_mode:
            self.current_case_key_index += 1
            self.current_step_index = 0
            self.load_case()
            return
        self.current_index += 1
        self.load_case()
        self.screenshot_done_event.clear()

    def on_next_step(self):
        if not self.screenshot_done_event.is_set():
            self._show_warning("请先完成当前步骤截图")
            return
        self.current_step_index += 1
        self.screenshot_done_event.clear()
        self.load_case()

    def on_skip(self):
        if self.is_step_mode:
            self.logger.log(f"跳过步骤 {self.current_step_index + 1}")
            self.current_step_index += 1
            self.screenshot_done_event.clear()

            if self.current_step_index >= len(self.current_case_steps):
                self.logger.log("当前用例所有步骤已跳过，进入下一用例")
                self.current_case_key_index += 1
                self.current_step_index = 0
            self.load_case()
            return

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

    def _on_resize(self, event):
        try:
            new_wraplength = max(event.width - 120, 100)
            for attr in ["lbl_case_name", "lbl_checkpoint", "lbl_step_name", "lbl_step_desc", "lbl_expected", "lbl_progress"]:
                label = getattr(self, attr, None)
                if label:
                    label.config(wraplength=new_wraplength)
        except Exception as e:
            self.logger.log(f"自动调整wraplength失败: {e}")