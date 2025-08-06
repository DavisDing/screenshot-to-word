# core/test_runner.py
import os
import threading
import time
import tkinter as tk
from tkinter import messagebox

from ui.control_panel import ControlPanel
from core.screenshot import take_screenshot
from ui.annotator import launch_annotator
from utils.word_generator import WordGenerator

class TestRunner:
    def __init__(self, root, excel_handler, logger, word_output_dir, temp_dir):
        self.root = root
        self.excel_handler = excel_handler
        self.logger = logger
        self.word_output_dir = word_output_dir
        self.temp_dir = temp_dir

        self.control_panel = None
        self.word_generator = WordGenerator(self.word_output_dir, self.logger)
        self.pending_cases = []
        self.current_index = 0
        self.quit_requested = False
        self.case_executed = False  # 是否已完成截图和插图

    def run(self):
        if not self.excel_handler.load_excel():
            return

        self.pending_cases = list(self.excel_handler.get_pending_cases())
        total = len(self.pending_cases)

        if total == 0:
            messagebox.showinfo("提示", "没有需要执行的用例。")
            return

        self.logger.write(f"共加载 {total} 条待执行用例。")
        self.control_panel = ControlPanel(on_capture=self.capture_screenshot,
                                          on_next=self.complete_case,
                                          on_skip=self.skip_case,
                                          on_quit=self.confirm_quit)
        self.next_case()

    def next_case(self):
        if self.quit_requested:
            return

        if self.current_index >= len(self.pending_cases):
            messagebox.showinfo("完成", "所有用例已执行完毕。")
            self.logger.write("✅ 所有用例执行完成。")
            self.control_panel.destroy()
            return

        idx, row = self.pending_cases[self.current_index]
        self.case_index = idx
        self.case_name = str(row["用例文件名"]).strip()
        self.case_desc = str(row["验证点"]).strip()
        self.case_executed = False  # 重置执行状态

        self.logger.write(f"➡️ 当前用例：{self.case_name} - {self.case_desc}")
        self.control_panel.update_case(self.case_name, self.case_desc)

    def capture_screenshot(self):
        filename = f"{self.case_name}_{int(time.time())}.png"
        temp_path = os.path.join(self.temp_dir, filename)
        take_screenshot(temp_path)

        marked_path = launch_annotator(temp_path)
        if not marked_path:
            self.logger.write("⚠️ 用户取消标注")
            return

        self.word_generator.insert_case_image(self.case_name, self.case_desc, marked_path)
        self.excel_handler.mark_executed(self.case_index, status="已执行")
        self.case_executed = True

        if messagebox.askyesno("截图完成", "是否继续当前用例截图？\n是：继续当前截图\n否：进入下一条用例"):
            self.logger.write("📌 用户选择继续截图")
        else:
            self.current_index += 1
            self.next_case()

    def complete_case(self):
        if not self.case_executed:
            messagebox.showwarning("未完成截图", "请先截图并完成插图后再点击完成。")
            return
        self.logger.write(f"✅ 用例完成：{self.case_name}")
        self.current_index += 1
        self.next_case()

    def skip_case(self):
        self.logger.write(f"⏭️ 跳过用例：{self.case_name}")
        self.current_index += 1
        self.next_case()

    def confirm_quit(self):
        done = self.current_index
        remain = len(self.pending_cases) - done
        if messagebox.askokcancel("确认退出", f"已执行 {done} 条，还剩 {remain} 条未执行，是否退出？"):
            self.quit_requested = True
            self.logger.write("❌ 用户手动退出执行流程。")
            if self.control_panel:
                self.control_panel.destroy()