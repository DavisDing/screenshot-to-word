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

            if self.excel_handler.version == "步骤版":
                self.pending_cases = list(self.excel_handler.get_pending_cases())
                if not self.pending_cases:
                    self.root.attributes('-topmost', True)
                    messagebox.showinfo("提示", "所有用例均已执行，无需重复执行。", parent=self.root)
                    self.root.attributes('-topmost', False)
                    self.logger.log("无待执行步骤用例，测试结束")
                    return

                pending_keys = {(case_name, checkpoint) for _, case_name, checkpoint in self.pending_cases}
                all_step_cases = self.excel_handler.get_step_cases()
                filtered_cases = {k: v for k, v in all_step_cases.items() if k in pending_keys}

                self.logger.log(f"步骤版待执行用例组数：{len(filtered_cases)}")
                self.control_panel = ControlPanel(self.logger, filtered_cases, self.excel_handler, self.root, is_step_mode=True)
                self.root.withdraw()
                self.control_panel.deiconify()
                self.control_panel.grab_set()
                # 等待control_panel窗口关闭后再判断保存结果
                self.control_panel.wait_window()
                # 修改: 通过control_panel获取annotator对象
                if hasattr(self.control_panel, 'annotator'):
                    annotator = self.control_panel.annotator
                    if not getattr(annotator, "save_result", False):
                        self.logger.log("用户选择不保存标注，跳过当前步骤/用例。")
                        self.root.deiconify()
                        return
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

            self.control_panel = ControlPanel(self.logger, self.pending_cases, self.excel_handler, self.root)
            self.control_panel.deiconify()
            self.control_panel.grab_set()
            # 等待control_panel窗口关闭后再判断保存结果
            self.control_panel.wait_window()

            if hasattr(self.control_panel, 'annotator'):
                annotator = self.control_panel.annotator
                # 当save_result为False或None时，跳过当前截图处理，直接返回控制面板
                if not getattr(annotator, "save_result", True):
                    self.logger.log("用户选择不保存标注，跳过当前步骤/用例。")
                    self.root.deiconify()
                    return
        except Exception as e:
            self.logger.log(f"测试运行异常：{e}")
            self.root.attributes('-topmost', True)
            messagebox.showerror("错误", f"测试运行异常：{e}", parent=self.root)
            self.root.attributes('-topmost', False)