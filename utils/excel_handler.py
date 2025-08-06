import os
import pandas as pd
from tkinter import filedialog, messagebox

class ExcelHandler:
    def __init__(self, excel_dir, logger):
        self.excel_dir = excel_dir
        self.logger = logger
        self.df = None
        self.file_path = None

    def load_excel(self):
        files = [f for f in os.listdir(self.excel_dir) if f.endswith(".xlsx") or f.endswith(".xls")]
        if not files:
            messagebox.showerror("错误", "未在 excel_input 目录中发现 Excel 文件！")
            return False

        if len(files) == 1:
            selected = os.path.join(self.excel_dir, files[0])
        else:
            selected = filedialog.askopenfilename(
                title="请选择要执行的 Excel 文件",
                initialdir=self.excel_dir,
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
            if not selected:
                return False

        try:
            self.df = pd.read_excel(selected, header=0, usecols=[0, 1, 2])
            self.df.columns = ["用例文件名", "验证点", "执行结果"]
            self.file_path = selected
            self.logger.write(f"成功加载 Excel 文件：{selected}")
        except Exception as e:
            messagebox.showerror("错误", f"读取 Excel 文件失败：{e}")
            self.logger.write(f"[异常] 无法读取 Excel：{e}")
            return False

        return True

    def get_pending_cases(self):
        pending = self.df[
            ~self.df["执行结果"].astype(str).str.lower().isin(["已执行", "pass", "passed"])
        ]
        return pending.iterrows()

    def update_case_status(self, case_name, status="已执行"):
        if self.df is not None:
            idx = self.df[self.df["用例文件名"] == case_name].index
            if not idx.empty:
                self.df.at[idx[0], "执行结果"] = status
                try:
                    self.df.to_excel(self.file_path, index=False)
                    self.logger.write(f"已更新用例 {case_name} 状态为：{status}")
                except PermissionError:
                    self.logger.write("[异常] 无法写入 Excel，文件可能被打开")
                except Exception as e:
                    self.logger.write(f"[异常] Excel 写入失败：{e}")