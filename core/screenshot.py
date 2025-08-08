# core/screenshot.py

import os
import time
import pyautogui
import threading
from ui.annotator import Annotator
import tkinter as tk
from utils.path_utils import get_base_path


class ScreenshotTool:
    def __init__(self, logger, root):
        self.logger = logger
        self.root = root
        base_dir = get_base_path()
        self.temp_dir = os.path.join(base_dir, "Temp")
        os.makedirs(self.temp_dir, exist_ok=True)

    def capture_screen(self, filename):
        """
        全屏截图，保存到Temp目录，返回图片路径
        """
        try:
            filepath = os.path.join(self.temp_dir, filename)
            img = pyautogui.screenshot()
            img.save(filepath)
            self.logger.log(f"截图保存：{filepath}")
            return filepath
        except Exception as e:
            self.logger.log(f"截图失败：{e}")
            return None

    def annotate(self, image_path, done_event):
        """
        在UI主线程中弹出标注窗口.
        用户关闭窗口后，通过 done_event 通知调用方.
        """
        self.logger.log("弹出标注窗口")

        def on_annotator_close():
            self.logger.log("标注窗口关闭")
            done_event.set()

        try:
            if self.root.winfo_exists():
                annotator = Annotator(self.root, image_path, on_close_callback=on_annotator_close)
                annotator.grab_set()
            else:
                done_event.set()
        except Exception as e:
            self.logger.log(f"创建标注窗口失败: {e}")
            done_event.set()