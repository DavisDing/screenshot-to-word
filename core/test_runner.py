import os
import threading
from tkinter import messagebox
from datetime import datetime

from ui.control_panel import ControlPanel
from ui.annotator import launch_annotator
from core.screenshot import take_screenshot
from utils.word_generator import WordGenerator

class TestRunner:
    def __init__(self, root, excel_handler, logger, word_output_dir, temp_dir):
        self.root = root
        self.excel_handler = excel_handler
        self.logger = logger
        self.word_output_dir = word_output_dir
        self.temp_dir = temp_dir

        self.word_generator = WordGenerator(word_output_dir, self.logger)
        self.control_panel = None

        self.cases = []
        self.current_case_index = 0
        self.stop_flag = threading.Event()
        self.lock = threading.Lock()

    def run(self):
        try:
            self.cases = []
            for idx, row in self.excel_handler.get_pending_cases():
                self.cases.append({
                    'index': idx,
                    'case_name': row['用例文件名'],
                    'case_desc': row['验证点'],
                    'status': row['执行结果']
                })
            if not self.cases:
                messagebox.showinfo("提示", "无可执行用例")
                return
            self.root.after(0, self._init_control_panel)
        except Exception as e:
            self.logger.write(f"[异常] 流程启动失败：{e}")

    def _init_control_panel(self):
        def on_capture():
            self._trigger_screenshot()

        def on_next():
            self._complete_current_case()

        def on_skip():
            self._skip_current_case()

        def on_quit():
            self._confirm_exit()

        self.control_panel = ControlPanel(on_capture, on_next, on_skip, on_quit)
        self._update_control_panel()

    def _update_control_panel(self):
        if self.current_case_index >= len(self.cases):
            messagebox.showinfo("提示", "所有用例已执行完毕！")
            self.control_panel.destroy()
            self.root.deiconify()
            return
        case = self.cases[self.current_case_index]
        self.control_panel.update_case(case['case_name'], case['case_desc'])

    def _trigger_screenshot(self):
        with self.lock:
            case = self.cases[self.current_case_index]
            case_name = case['case_name']
            case_desc = case['case_desc']
            filename = f"{case_name}_{self._current_timestamp()}.png"
            filepath = os.path.join(self.temp_dir, filename)
            try:
                take_screenshot(filepath)
                marked_path = launch_annotator(filepath)
                self.word_generator.insert_case_image(case_name, case_desc, marked_path)
                self.logger.write(f"[操作] 完成截图并插入Word：{marked_path}")
                messagebox.showinfo("截图完成", "截图已保存并插入Word文档！")
            except Exception as e:
                self.logger.write(f"[异常] 截图或插图失败：{e}")

    def _complete_current_case(self):
        case = self.cases[self.current_case_index]
        self.excel_handler.update_case_status(case['case_name'], "已执行")
        self._next_case()

    def _skip_current_case(self):
        case = self.cases[self.current_case_index]
        self.logger.write(f"[INFO] 跳过用例：{case['case_name']}")
        self.excel_handler.update_case_status(case['case_name'], "跳过")
        self._next_case()

    def _next_case(self):
        self.current_case_index += 1
        self._update_control_panel()

    def _confirm_exit(self):
        executed = sum(1 for c in self.cases if c.get('status') == '已执行')
        remain = len(self.cases) - executed
        if messagebox.askokcancel("退出确认", f"当前执行了 {executed} 条，还剩 {remain} 条，确认退出？"):
            self.root.destroy()

    def _current_timestamp(self):
        return datetime.now().strftime("%Y%m%d%H%M%S")