import os
import threading
import tkinter as tk
from tkinter import messagebox
from core.screenshot import take_screenshot_and_annotate
from utils.excel_handler import ExcelHandler
from utils.logger import Logger
from ui.control_panel import ControlPanel
from docx import Document
from docx.shared import Inches

class TestRunner:
    def __init__(self, excel_path):
        self.excel_handler = ExcelHandler(excel_path)
        self.logger = Logger()
        self.test_cases = self.excel_handler.get_pending_cases()
        self.current_index = 0
        self.control_panel = ControlPanel(self)
        self.word_docs = {}

    def start(self):
        self.logger.log("测试执行开始")
        threading.Thread(target=self.run_tests, daemon=True).start()

    def run_tests(self):
        while self.current_index < len(self.test_cases):
            case = self.test_cases[self.current_index]
            filename = case['filename']
            description = case['description']
            self.logger.log(f"当前用例：{filename}，验证点：{description}")
            self.control_panel.update_case_info(filename, description)
            self.control_panel.enable_controls(True)
            break  # 等待用户点击“截图”或快捷键

    def capture_screenshot(self):
        case = self.test_cases[self.current_index]
        filename = case['filename']
        description = case['description']
        image_path = take_screenshot_and_annotate(filename, description)
        self.logger.log(f"截图保存：{image_path}")
        self.insert_into_word(filename, description, image_path)
        self.excel_handler.mark_case_as_executed(self.current_index)
        self.logger.log("Excel 状态更新：已执行")

        if messagebox.askyesno("继续操作", "是否继续截图当前用例？"):
            return  # 不进入下一条
        else:
            self.next_case()

    def next_case(self):
        self.current_index += 1
        if self.current_index < len(self.test_cases):
            self.run_tests()
        else:
            self.logger.log("所有用例执行完毕")
            messagebox.showinfo("完成", "所有用例已执行完成。")
            self.control_panel.enable_controls(False)

    def skip_case(self):
        self.logger.log("跳过当前用例")
        self.next_case()

    def insert_into_word(self, filename, description, image_path):
        folder = os.path.join("word_output")
        os.makedirs(folder, exist_ok=True)
        doc_path = os.path.join(folder, f"{filename}.docx")

        if filename not in self.word_docs:
            if os.path.exists(doc_path):
                doc = Document(doc_path)
            else:
                doc = Document()
                doc.add_heading(filename, level=1)
                doc.add_paragraph(f"验证点：{description}")
            self.word_docs[filename] = doc

        doc = self.word_docs[filename]
        doc.add_picture(image_path, width=Inches(5.5))
        doc.add_paragraph("")

        doc.save(doc_path)
        self.logger.log(f"截图插入 Word：{doc_path}")