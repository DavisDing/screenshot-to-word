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
        try:
            self.logger.log("测试运行开始")
            if not self.excel_handler.select_excel_file():
                self.logger.log("选择Excel文件失败或用户取消")
                return

            if not self.excel_handler.load_excel():
                self.logger.log("加载 Excel 文件失败")
                return

            self.pending_cases = list(self.excel_handler.get_pending_cases())
            if not self.pending_cases:
                self.root.attributes('-topmost', True)
                messagebox.showinfo("提示", "所有用例均已执行，无需重复执行。", parent=self.root)
                self.root.attributes('-topmost', False)
                self.logger.log("无待执行用例，测试结束")
                return

            self.logger.log(f"共加载 {len(self.pending_cases)} 条待执行用例")
            self.root.withdraw()

            from ui.control_panel import ControlPanel
            self.control_panel = ControlPanel(self.logger, self.pending_cases, self.excel_handler, self.root)
            self.control_panel.deiconify()
            self.control_panel.grab_set()
        except Exception as e:
            self.logger.log(f"测试运行异常：{e}")
            self.root.attributes('-topmost', True)
            messagebox.showerror("错误", f"测试运行异常：{e}", parent=self.root)
            self.root.attributes('-topmost', False)