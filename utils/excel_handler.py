import os
import sys
import pandas as pd
from openpyxl import load_workbook
from tkinter import messagebox
from tkinter import filedialog

class ExcelHandler:
    def __init__(self, logger=None):
        self.logger = logger
        self.file_path = None
        self.df = None
        self.workbook = None
        self.sheet_name = None

    def select_excel_file(self):
        base_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, 'frozen', False) else os.getcwd()
        input_dir = os.path.join(base_dir, "excel_input")
        files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
        if not files:
            if self.logger:
                self.logger.log("excel_input 目录中未找到 Excel 文件，准备进入手动选择")
            messagebox.showwarning("提示", "excel_input 目录中未找到任何 Excel 文件，是否手动选择？")
            manual_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel files", "*.xlsx *.xls")])
            if manual_path:
                self.file_path = manual_path
                if self.logger:
                    self.logger.log(f"手动选中Excel文件: {self.file_path}")
                return self.file_path
            return None
        if len(files) == 1:
            self.file_path = os.path.join(input_dir, files[0])
        else:
            from tkinter import simpledialog
            choice = simpledialog.askstring("选择文件", f"请选择文件，输入编号：\n" +
                                         '\n'.join([f'{i+1}. {fn}' for i, fn in enumerate(files)]))
            try:
                idx = int(choice) - 1
                self.file_path = os.path.join(input_dir, files[idx])
            except Exception:
                messagebox.showerror("错误", "文件选择无效！")
                return None
        if self.logger:
            self.logger.log(f"选中Excel文件: {self.file_path}")
        return self.file_path

    def load_excel(self):
        if not self.file_path:
            return False
        try:
            self.workbook = load_workbook(self.file_path)
            self.sheet_name = self.workbook.sheetnames[0]
            ws = self.workbook[self.sheet_name]
            data = ws.values
            columns = next(data)
            self.df = pd.DataFrame(data, columns=columns)
            if self.logger:
                self.logger.log(f"加载Excel表格: {self.file_path}")
            return True
        except PermissionError:
            messagebox.showerror("错误", f"文件被占用或权限不足: {self.file_path}")
            if self.logger:
                self.logger.log(f"无法打开Excel文件（权限或占用）: {self.file_path}")
            return False
        except Exception as e:
            messagebox.showerror("错误", f"Excel加载失败: {e}")
            if self.logger:
                self.logger.log(f"Excel加载异常: {e}")
            return False

    def get_valid_cases(self):
        if self.df is None:
            return []
        # 过滤 C 列为已执行、PASS、Passed
        mask = ~self.df.iloc[:, 2].astype(str).str.lower().isin(['已执行', 'pass', 'passed'])
        filtered_df = self.df[mask]
        return filtered_df.reset_index(drop=True)

    def update_case_status(self, index, status='已执行'):
        try:
            if self.df is None or self.workbook is None:
                return
            # 先更新 pandas DataFrame
            self.df.iat[index, 2] = status
            # 再写回 Excel 文件（第三列）
            ws = self.workbook[self.sheet_name]
            excel_row = index + 2  # 因为表头占1行，DataFrame从0开始
            ws.cell(row=excel_row, column=3, value=status)
            self.workbook.save(self.file_path)
            if self.logger:
                self.logger.log(f"更新用例状态: 行 {excel_row} -> {status}")
        except PermissionError:
            if self.logger:
                self.logger.log("无法保存Excel，文件可能被占用")
        except Exception as e:
            if self.logger:
                self.logger.log(f"更新Excel异常: {e}")
