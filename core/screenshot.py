# core/screenshot.py
import pyautogui
import time

def take_screenshot(save_path):
    # 等待短暂时间让用户准备好界面
    time.sleep(0.5)
    image = pyautogui.screenshot()
    image.save(save_path)