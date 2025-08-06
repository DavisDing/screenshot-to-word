import os
import time
import threading
from tkinter import messagebox
from ui.control_panel import ControlPanel
from core.screenshot import take_screenshot
from ui.annotator import launch_annotator
from docx import Document
from docx.shared import Inches
from datetime import datetime

class TestRunner:
    def __init__(self, logger, excel_handler):
        self.logger = logger
        self.excel_handler = excel_handler
        self.control_panel = None
        self.stop_flag = threading.Event()

    def run(self):
        cases = self.excel_handler.load_cases()
        if cases is None or cases.empty:
            self.logger.log("无有效用例可执行", level="error")
            return

        self.control_panel = ControlPanel(self.on_screenshot, self.on_skip, self.on_exit)
        self.logger.log(f"加载 {len(cases)} 条用例，开始执行...\n")

        for idx, row in cases.iterrows():
            self.current_index = idx
            filename = row['测试名称']
            checkpoint = row['验证点']
            self.logger.log(f"执行用例：{filename} - {checkpoint}")
            self.control_panel.update_case(filename, checkpoint)
            self.control_panel.wait_for_action()

            if self.stop_flag.is_set():
                break

        self.logger.log("全部用例执行完毕。\n")
        messagebox.showinfo("完成", "所有用例已执行完毕")
        self.control_panel.destroy()

    def on_screenshot(self):
        row = self.excel_handler.df.loc[self.current_index]
        filename = row['测试名称']
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        image_path = os.path.join("word_output", filename, f"{filename}_{timestamp}.png")
        os.makedirs(os.path.dirname(image_path), exist_ok=True)
        take_screenshot(image_path)

        edited_image = launch_annotator(image_path)
        self.insert_to_word(filename, row['验证点'], edited_image)
        self.excel_handler.mark_as_executed(self.current_index)
        self.control_panel.reset_action()

    def on_skip(self):
        self.logger.log("用户跳过当前用例\n")
        self.excel_handler.mark_as_executed(self.current_index)
        self.control_panel.reset_action()

    def on_exit(self):
        self.stop_flag.set()
        self.control_panel.destroy()

    def insert_to_word(self, filename, checkpoint, image_path):
        doc_path = os.path.join("word_output", f"{filename}.docx")
        try:
            if os.path.exists(doc_path):
                doc = Document(doc_path)
            else:
                doc = Document()
                doc.add_heading(f"验证点：{filename}", level=1)
                doc.add_paragraph(checkpoint)

            doc.add_picture(image_path, width=Inches(5))
            doc.add_paragraph("\n")
            doc.save(doc_path)
            self.logger.log(f"截图保存并写入 Word：{doc_path}\n")
        except Exception as e:
            self.logger.log(f"写入 Word 失败：{e}", level="error")
