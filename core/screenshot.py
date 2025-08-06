# core/screenshot.py
import os
import datetime
import pyautogui

class Screenshot:
    def __init__(self, logger=None):
        self.logger = logger
        self.temp_dir = os.path.join(os.getcwd(), "Temp")
        os.makedirs(self.temp_dir, exist_ok=True)

    def take_screenshot(self, case_name):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{case_name}_{timestamp}.png"
        filepath = os.path.join(self.temp_dir, filename)
        try:
            image = pyautogui.screenshot()
            image.save(filepath)
            if self.logger:
                self.logger.log(f"截图保存成功: {filepath}")
            return filepath
        except Exception as e:
            if self.logger:
                self.logger.log(f"截图失败: {e}")
            return None