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

    def annotate(self, image_path):
        """
        弹出标注窗口，阻塞等待用户操作完成后返回最终图片路径
        """
        # ✅ 等待图片生成，确保可用（最多等待2秒）
        for _ in range(40):
            if os.path.exists(image_path):
                break
            time.sleep(0.05)
        else:
            self.logger.log(f"标注前图片不存在：{image_path}")
            return None

        result = {'done': False}

        def run_annotator():
            annotator = Annotator(self.root, image_path)
            annotator.grab_set()
            annotator.wait_window()
            result['done'] = True

        thread = threading.Thread(target=run_annotator)
        thread.start()
        # 等待标注完成
        while not result['done']:
            self.root.update()
            time.sleep(0.05)

        self.logger.log(f"标注完成：{image_path}")
        return image_path