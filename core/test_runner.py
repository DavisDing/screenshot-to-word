import threading
from tkinter import messagebox

from ui.control_panel import ControlPanel


class TestRunner:
    def __init__(self, logger, excel_handler, root):
        self.logger = logger
        self.excel_handler = excel_handler
        self.root = root
        self.control_panel = None
        self.pending_cases = []

    def run_tests(self):
        # This method runs in a background thread.
        # All UI operations must be scheduled on the main thread via root.after()
        self.logger.log("测试运行开始")
        if not self.excel_handler.select_excel_file():
            self.logger.log("选择Excel文件失败或用户取消")
            return

        if not self.excel_handler.load_excel():
            self.logger.log("加载 Excel 文件失败")
            return

        def _launch_control_panel():
            try:
                if self.excel_handler.version == "步骤版":
                    step_cases = self.excel_handler.get_step_cases()
                    if not step_cases:
                        self.logger.log("步骤版无有效用例")
                        self.excel_handler._show_info("步骤版无有效用例")
                        return
                    self.control_panel = ControlPanel(self.logger, step_cases, self.excel_handler, self.root, is_step_mode=True)
                else: # 基础版
                    self.pending_cases = list(self.excel_handler.get_pending_cases())
                    if not self.pending_cases:
                        self.excel_handler._show_info("所有用例均已执行，无需重复执行。")
                        self.logger.log("无待执行用例，测试结束")
                        return
                    self.logger.log(f"共加载 {len(self.pending_cases)} 条待执行用例")
                    self.control_panel = ControlPanel(self.logger, self.pending_cases, self.excel_handler, self.root)

                self.root.withdraw()
                self.control_panel.deiconify()
                self.control_panel.grab_set()
            except Exception as e:
                self.logger.log(f"测试运行异常：{e}")
                self.excel_handler._show_error(f"测试运行异常：{e}")

        self.root.after(0, _launch_control_panel)