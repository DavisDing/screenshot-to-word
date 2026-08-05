# ui/control_panel.py

import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk

import keyboard

from core.screenshot import ScreenshotTool
from utils.word_generator import WordGenerator


class ControlPanel(tk.Toplevel):
    def __init__(self, logger, pending_cases, excel_handler, root, is_step_mode=False):
        super().__init__(root)
        self.logger = logger
        self.pending_cases = pending_cases
        self.excel_handler = excel_handler
        self.root = root

        self.title("执行控制面板")
        self.geometry("760x520")
        self.minsize(680, 460)
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.resizable(True, True)

        try:
            self.configure(background="#f5f7fb")
        except Exception:
            pass

        # 状态变量
        self.current_index = 0
        self.current_case = None
        self.screenshot_tool = ScreenshotTool(self.logger, root)
        self.word_generator = WordGenerator(self.logger, root)
        self.screenshot_done_event = threading.Event()

        self.is_step_mode = is_step_mode
        if self.is_step_mode:
            self.step_case_keys = list(pending_cases.keys())
            self.current_case_key_index = 0
            self.current_step_index = 0
            self.pending_cases = pending_cases
            self.current_case_steps = []

        self.status_var = tk.StringVar(value="准备执行")
        self.case_name_var = tk.StringVar(value="-")
        self.checkpoint_var = tk.StringVar(value="-")
        self.step_name_var = tk.StringVar(value="-")
        self.step_desc_var = tk.StringVar(value="-")
        self.expected_var = tk.StringVar(value="-")
        self.progress_var = tk.StringVar(value="-")

        self.create_widgets()
        self.load_case()
        self.bind("<Configure>", self._on_resize)

    def create_widgets(self):
        container = ttk.Frame(self, padding=18, style="App.TFrame")
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        header = ttk.Frame(container, padding=16, style="Card.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="执行控制面板", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=1, sticky="e", padx=(12, 0))
        ttk.Label(
            header,
            text="快捷键：F8 截图；完成标注后再进入下一步 / 完成当前用例。",
            style="Hint.TLabel",
            wraplength=660,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        button_bar = ttk.Frame(container, padding=(0, 14, 0, 0), style="App.TFrame")
        button_bar.grid(row=1, column=0, sticky="ew")
        for index in range(4):
            button_bar.columnconfigure(index, weight=1)

        self.btn_screenshot = ttk.Button(button_bar, text="截图 (F8)", command=self.on_screenshot, style="Primary.TButton")
        self.btn_screenshot.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        try:
            keyboard.add_hotkey('f8', self.on_screenshot)
        except Exception as e:
            self.logger.log(f"注册F8快捷键失败: {e}")

        self.btn_skip = ttk.Button(button_bar, text="跳过", command=self.on_skip, style="Warning.TButton")
        self.btn_skip.grid(row=0, column=1, sticky="ew", padx=8)

        self.btn_next = ttk.Button(button_bar, text="下一步", command=self.on_next_step, style="Secondary.TButton")
        self.btn_next.grid(row=0, column=2, sticky="ew", padx=8)

        self.btn_complete = ttk.Button(button_bar, text="完成", command=self.on_complete, style="Success.TButton")
        self.btn_complete.grid(row=0, column=3, sticky="ew", padx=(8, 0))

        content = ttk.Frame(container, style="App.TFrame")
        content.grid(row=2, column=0, sticky="nsew", pady=(14, 0))
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=0)
        content.rowconfigure(1, weight=1)
        content.rowconfigure(2, weight=0)

        self.case_frame = ttk.Labelframe(content, text="当前用例", padding=14)
        self.case_frame.grid(row=0, column=0, sticky="ew")
        self.case_frame.columnconfigure(1, weight=1)
        self._add_info_row(self.case_frame, 0, "用例名：", self.case_name_var)
        self._add_info_row(self.case_frame, 1, "验证点：", self.checkpoint_var)

        self.step_frame = ttk.Labelframe(content, text="步骤信息", padding=14)
        self.step_frame.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.step_frame.columnconfigure(1, weight=1)
        self._add_info_row(self.step_frame, 0, "步骤名称：", self.step_name_var)
        self._add_info_row(self.step_frame, 1, "步骤描述：", self.step_desc_var)
        self._add_info_row(self.step_frame, 2, "预期结果：", self.expected_var)

        progress_frame = ttk.Frame(content, padding=14, style="SoftCard.TFrame")
        progress_frame.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        progress_frame.columnconfigure(1, weight=1)
        ttk.Label(progress_frame, text="当前进度：", style="Status.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(progress_frame, textvariable=self.progress_var, style="AccentValue.TLabel").grid(row=0, column=1, sticky="w")

        if not self.is_step_mode:
            self.step_frame.grid_remove()
            self.btn_next.grid_remove()

    def _add_info_row(self, parent, row, label_text, variable):
        ttk.Label(parent, text=label_text, style="Section.TLabel").grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=6)
        ttk.Label(parent, textvariable=variable, style="Value.TLabel", wraplength=560, justify="left").grid(
            row=row, column=1, sticky="ew", pady=6
        )

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
            self.case_name_var.set(case_key[0] or "-")
            self.checkpoint_var.set(case_key[1] or "-")
            self.step_name_var.set(step.get('步骤名称', '') or "-")
            self.step_desc_var.set(step.get('步骤描述', '') or "-")
            self.expected_var.set(step.get('预期结果', '') or "-")
            remaining_cases = len(self.step_case_keys) - self.current_case_key_index - 1
            self.progress_var.set(f"第 {self.current_step_index + 1} 步 / 共 {len(self.current_case_steps)} 步，剩余 {remaining_cases} 条案例")
            self.status_var.set("步骤执行中")
            self.btn_complete.config(state="disabled")

            self.step_frame.grid()
            self.btn_next.grid()

            steps = self.current_case_steps
            if self.current_step_index == len(steps) - 1:
                self.btn_next.config(state="disabled")
                self.btn_complete.config(state="normal" if self.screenshot_done_event.is_set() else "disabled")
            else:
                self.btn_next.config(state="normal")
                self.btn_complete.config(state="disabled")
            return

        # 基础版处理
        if self.current_index >= len(self.pending_cases):
            self.finish_all_cases()
            return

        idx, filename, checkpoint = self.pending_cases[self.current_index]
        self.current_case = (idx, filename, checkpoint)
        self.case_name_var.set(filename or "-")
        self.checkpoint_var.set(checkpoint or "-")
        self.step_name_var.set("-")
        self.step_desc_var.set("-")
        self.expected_var.set("-")
        self.progress_var.set(f"第 {self.current_index + 1} 条 / 共 {len(self.pending_cases)} 条")
        self.status_var.set("用例执行中")
        self.logger.log(f"当前执行用例：{filename} - 验证点：{checkpoint}")

        self.step_frame.grid_remove()
        self.btn_next.grid_remove()
        self.btn_complete.config(state="disabled")

    def on_screenshot(self):
        # 截图 + 标注 + Word生成流程
        def run_screenshot_flow():
            idx, filename, checkpoint = self.current_case
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            file_name = f"{filename}_{timestamp}.png"

            self.logger.log(f"开始截图: {file_name}")

            img_path = self.screenshot_tool.capture_screen(file_name)

            if not img_path:
                self.logger.log("截图失败或取消")
                self.screenshot_done_event.clear()
                self.root.after(0, self._restore_control_panel)
                return

            annotated_path = self.screenshot_tool.annotate(img_path)
            if not annotated_path:
                self.logger.log("标注取消或失败")
                self.screenshot_done_event.clear()
                self.root.after(0, self._restore_control_panel)
                return

            step_note = ""
            if self.is_step_mode:
                step = self.current_case_steps[self.current_step_index]
                step_note = f"{step.get('步骤名称', '')} - {step.get('步骤描述', '')}".strip(" -")

            self.word_generator.add_image_to_word(filename, checkpoint, annotated_path, step_note)

            self.excel_handler.mark_case_executed(idx)
            self.logger.log(f"用例 {filename} 标记为已执行")

            self.screenshot_done_event.set()
            self.logger.log("screenshot_done_event.set() 已调用")
            self.root.after(0, self._apply_post_screenshot_button_state)
            self.root.after(0, self._restore_control_panel)

        if not self.current_case:
            self._show_warning("当前没有可执行用例。")
            return

        self.screenshot_done_event.clear()
        self.status_var.set("截图中")
        self.attributes("-topmost", False)  # 取消置顶，防止截图残留
        self.withdraw()  # 隐藏控制面板
        threading.Thread(target=run_screenshot_flow, daemon=True).start()


    def _apply_post_screenshot_button_state(self):
        if self.is_step_mode:
            is_last_step = self.current_step_index == len(self.current_case_steps) - 1
            self.btn_next.config(state="disabled" if is_last_step else "normal")
            self.btn_complete.config(state="normal" if is_last_step else "disabled")
            return
        self.btn_complete.config(state="normal")

    def _restore_control_panel(self):
        self.deiconify()
        self.attributes('-topmost', True)
        self.status_var.set("标注已完成" if self.screenshot_done_event.is_set() else "等待操作")

    def on_complete(self):
        if self.is_step_mode and self.current_step_index == len(self.current_case_steps) - 1:
            if not self.screenshot_done_event.is_set():
                self._show_warning("请先完成当前步骤截图")
                return
            self.current_case_key_index += 1
            self.current_step_index = 0
            self.screenshot_done_event.clear()
            self.load_case()
            return

        if not self.screenshot_done_event.is_set():
            self._show_warning("请先截图并完成标注后，再点击完成。")
            return

        if self.is_step_mode:
            self.current_case_key_index += 1
            self.current_step_index = 0
            self.screenshot_done_event.clear()
            self.load_case()
            return

        self.current_index += 1
        self.screenshot_done_event.clear()
        self.load_case()

    def on_next_step(self):
        self.logger.log(f"on_next_step: screenshot_done_event.is_set()={self.screenshot_done_event.is_set()}")
        if not self.screenshot_done_event.is_set():
            self.logger.log("on_next_step: 检测到 screenshot_done_event 未 set，阻止进入下一步")
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
        try:
            keyboard.remove_hotkey('f8')
            self.logger.log("F8快捷键已移除")
        except Exception as e:
            self.logger.log(f"移除F8快捷键失败: {e}")
        self.root.attributes('-topmost', True)
        messagebox.showinfo("完成", "所有用例已执行完毕", parent=self.root)
        self.root.attributes('-topmost', False)
        self.destroy()
        self.root.deiconify()

    def on_exit(self):
        self.root.attributes('-topmost', True)
        if messagebox.askokcancel("退出", "确定退出测试？", parent=self.root):
            try:
                keyboard.remove_hotkey('f8')
                self.logger.log("F8快捷键已移除")
            except Exception as e:
                self.logger.log(f"移除F8快捷键失败: {e}")
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
            width = self.winfo_width()
            new_wraplength = max(width - 200, 180)
            for child_frame in (self.case_frame, self.step_frame):
                for child in child_frame.winfo_children():
                    if isinstance(child, ttk.Label) and str(child.cget("style")) == "Value.TLabel":
                        child.config(wraplength=new_wraplength)
        except Exception as e:
            self.logger.log(f"自动调整wraplength失败: {e}")
