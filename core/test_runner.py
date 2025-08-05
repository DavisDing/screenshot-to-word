import os
import keyboard
import shutil
from docx import Document
from tkinter import messagebox
from ui.annotator import Annotator
from ui.control_panel import ControlPanel
from core.screenshot import take_screenshot

class TestRunner:
    def __init__(self, excel_handler, logger, status_label, root):
        self.excel_handler = excel_handler
        self.logger = logger
        self.status_label = status_label
        self.root = root
        self.cases = excel_handler.load_cases()
        self.current_index = 0
        self.screenshot_ready = False

    def run_tests(self):
        self.logger.create_window()
        self.panel = ControlPanel(self.trigger_screenshot, self.skip_case)
        keyboard.add_hotkey("F8", self.trigger_screenshot)

        for idx, row in self.cases.iterrows():
            self.current_index = idx
            filename, checkpoint = row[0], row[1]
            self.logger.log(f\"用例: {filename} - {checkpoint}\")
            self.status_label.config(text=f\"执行：{filename}\")

            self.screenshot_ready = False
            while not self.screenshot_ready:
                pass  # 等待截图

            self.insert_into_word(filename, checkpoint)
            self.excel_handler.update_status(idx)
            response = messagebox.askyesno(\"继续？\", \"是否继续下一条用例？\") 
            if not response:
                break

        keyboard.remove_hotkey(\"F8\")
        self.panel.destroy()
        self.root.deiconify()

    def trigger_screenshot(self):
        path = take_screenshot()
        Annotator(path, self._on_annotation_done)

    def _on_annotation_done(self, path):
        self.annotated_path = path
        self.screenshot_ready = True

    def insert_into_word(self, filename, checkpoint):
        doc_dir = os.path.join(\"word_output\", filename)
        os.makedirs(doc_dir, exist_ok=True)
        word_path = os.path.join(doc_dir, f\"{filename}.docx\")

        if os.path.exists(word_path):
            doc = Document(word_path)
        else:
            doc = Document()
            doc.add_paragraph(f\"验证点：{checkpoint}\")

        doc.add_picture(self.annotated_path, width=None)
        doc.save(word_path)
        os.remove(self.annotated_path)

    def skip_case(self):
        self.screenshot_ready = True