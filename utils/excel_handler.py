import os
import pandas as pd
from openpyxl import load_workbook

class ExcelHandler:
    def __init__(self, folder):
        self.folder = folder
        self.file_path = self._get_latest_file()
        self.df = pd.read_excel(self.file_path)
    
    def _get_latest_file(self):
        files = [f for f in os.listdir(self.folder) if f.endswith('.xlsx')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.folder, x)), reverse=True)
        return os.path.join(self.folder, files[0])

    def load_cases(self):
        return self.df[~self.df.iloc[:, 2].astype(str).str.lower().isin(['已执行', 'pass', 'passed'])]

    def update_status(self, row_idx, status='已执行'):
        self.df.iat[row_idx, 2] = status
        self.df.to_excel(self.file_path, index=False)