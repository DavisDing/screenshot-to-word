import os
import threading
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

from ui.control_panel import ControlPanel
from ui.annotator import launch_annotator
from core.screenshot import take_screenshot  # 请确保此函数已实现
from utils.word_generator import WordGenerator
from utils.logger import Logger

class TestRunner:
    def __init__(self, root, excel_handler, logger, word_output_dir, temp_dir):
        self.root = root
        self.excel_handler = excel_handler
        self.logger = logger
        self.word_output_dir = word_output_dir
        self.temp_dir = temp_dir

        self.word_generator = WordGenerator(word_output_dir)
        self.control_panel = None

        self.current_case_index = 0
        self.cases = self.excel_handler.get_cases()  # 应返回[{case_name, case_desc, status}, ...]

        self.stop_flag = threading.Event()
        self.lock = threading.Lock()

    def run(self):
        """主流程启动，开启新线程执行"""
        t = threading.Thread(target=self._run_cases, daemon=True)
        t.start()

    def _run_cases(self):
        if not self.cases:
            messagebox.showinfo("提示", "无可执行用例")
            return

        # 初始化控制面板，绑定按钮事件
        self.root.after(0, self._init_control_panel)
        self.process_case(self.current_case_index)

    def _init_control_panel(self):
        def on_capture():
            self.logger.write("[操作] 点击截图按钮")
            self._trigger_screenshot()

        def on_next():
            self.logger.write("[操作] 点击完成当前用例")
            self._complete_current_case()

        def on_skip():
            self.logger.write("[操作] 点击跳过当前用例")
            self._skip_current_case()

        def on_quit():
            self.logger.write("[操作] 点击退出")
            self._confirm_exit()

        self.control_panel = ControlPanel(on_capture, on_next, on_skip, on_quit)
        self._update_control_panel()

    def _update_control_panel(self):
        if self.current_case_index >= len(self.cases):
            self.control_panel.destroy()
            self.logger.write("[INFO] 所有用例执行完毕")
            messagebox.showinfo("完成", "所有用例已执行完毕！")
            return
        case = self.cases[self.current_case_index]
        self.control_panel.update_case(case['case_name'], case['case_desc'])

    def _trigger_screenshot(self):
        with self.lock:
            if self.current_case_index >= len(self.cases):
                return

            case = self.cases[self.current_case_index]
            case_name = case['case_name']
            case_desc = case['case_desc']

            screenshot_path = self._take_screenshot(case_name)
            if not screenshot_path:
                self.logger.write(f"[WARN] 截图失败，跳过用例：{case_name}")
                self._skip_current_case()
                return

            marked_path = launch_annotator(screenshot_path)
            if not marked_path or not os.path.exists(marked_path):
                self.logger.write(f"[WARN] 标注失败或取消，跳过用例：{case_name}")
                self._skip_current_case()
                return

            # Word文档插入
            try:
                self.word_generator.insert_case_image(case_name, case_desc, marked_path)
                self.logger.write(f"[INFO] 插入标注图片到 Word: {marked_path}")
            except Exception as e:
                self.logger.write(f"[ERROR] Word插入失败: {e}")

            # Excel状态更新
            self.excel_handler.update_case_status(case_name, "已执行")

            # 弹窗提示是否继续当前用例截图
            res = messagebox.askyesno(
                "继续截图",
                f"是否继续当前用例“{case_name}”截图？\n点击“否”进入下一条用例。"
            )
            if res:
                self.logger.write(f"[INFO] 继续当前用例截图：{case_name}")
                # 不改变index，继续当前用例等待下一次截图
                self._update_control_panel()
            else:
                self.logger.write(f"[INFO] 进入下一条用例")
                self._next_case()

    def _complete_current_case(self):
        # 强制完成当前用例（仅当已存在对应Word文档且至少完成一次截图标注）
        case = self.cases[self.current_case_index]
        case_name = case['case_name']
        word_path = os.path.join(self.word_output_dir, f"{case_name}.docx")
        if os.path.exists(word_path):
            self.excel_handler.update_case_status(case_name, "已执行")
            self.logger.write(f"[INFO] 用例已完成：{case_name}")
            self._next_case()
        else:
            messagebox.showwarning("警告", f"用例 {case_name} 尚未截图标注，无法完成。")

    def _skip_current_case(self):
        case = self.cases[self.current_case_index]
        case_name = case['case_name']
        self.logger.write(f"[INFO] 跳过用例：{case_name}")
        self.excel_handler.update_case_status(case_name, "跳过")
        self._next_case()

    def _next_case(self):
        self.current_case_index += 1
        if self.current_case_index >= len(self.cases):
            self.control_panel.destroy()
            self.logger.write("[INFO] 所有用例执行完毕")
            messagebox.showinfo("完成", "所有用例已执行完毕！")
            return
        self._update_control_panel()

    def _confirm_exit(self):
        executed = sum(1 for c in self.cases if c.get('status') == '已执行')
        remain = len(self.cases) - executed
        if messagebox.askokcancel("退出确认", f"当前执行了 {executed} 条，还剩 {remain} 条，确认退出？"):
            self.logger.write("[INFO] 用户退出测试流程")
            if self.control_panel:
                self.control_panel.destroy()
            self.root.quit()

    def _take_screenshot(self, case_name):
        filename = f"{case_name}_{self._current_timestamp()}.png"
        filepath = os.path.join(self.temp_dir, filename)
        try:
            take_screenshot(filepath)  # 你的截图实现
            self.logger.write(f"[INFO] 截图保存: {filepath}")
            return filepath
        except Exception as e:
            self.logger.write(f"[ERROR] 截图失败: {e}")
            return None

    def _current_timestamp(self):
        return datetime.now().strftime("%Y%m%d%H%M%S")