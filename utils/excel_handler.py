import os
import pandas as pd
from openpyxl import load_workbook

class ExcelHandler:
    def __init__(self, logger):
        self.logger = logger
        self.df = None
        self.filepath = None

    def load_cases(self):
        excel_dir = 'excel_input'
        files = [f for f in os.listdir(excel_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
        if not files:
            self.logger.log("未找到 Excel 文件", level="error")
            return None
        elif len(files) == 1:
            filename = files[0]
        else:
            # 多个文件，默认选择第一个（可扩展为弹窗选择）
            filename = files[0]
            self.logger.log(f"检测到多个 Excel 文件，默认使用第一个：{filename}")

        self.filepath = os.path.join(excel_dir, filename)

        try:
            df = pd.read_excel(self.filepath)
            if not {'测试名称', '验证点', '执行结果'}.issubset(df.columns):
                self.logger.log("Excel 缺少必要列：测试名称、验证点、执行结果", level="error")
                return None

            df = df[~df['执行结果'].astype(str).str.lower().isin(['pass', 'passed', '已执行'])].copy()
            self.df = df.reset_index(drop=True)
            return self.df

        except PermissionError:
            self.logger.log("Excel 文件被占用，无法读取", level="error")
            return None
        except Exception as e:
            self.logger.log(f"加载 Excel 异常：{e}", level="error")
            return None

    def mark_as_executed(self, index):
        try:
            wb = load_workbook(self.filepath)
            ws = wb.active
            row_index = index + 2  # 因为 DataFrame 默认从0起，Excel 从第2行是数据
            col = None
            for i, cell in enumerate(ws[1], start=1):
                if cell.value == '执行结果':
                    col = i
                    break
            if col:
                ws.cell(row=row_index, column=col).value = '已执行'
                wb.save(self.filepath)
            else:
                self.logger.log("找不到执行结果列，无法写入", level="error")
        except PermissionError:
            self.logger.log("Excel 文件正在使用中，无法写入", level="error")
        except Exception as e:
            self.logger.log(f"写入 Excel 异常：{e}", level="error")