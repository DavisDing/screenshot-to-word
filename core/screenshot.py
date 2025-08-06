import pyautogui
import time

def take_screenshot(path):
    time.sleep(0.5)  # 短暂延迟以避免截到控制面板
    screenshot = pyautogui.screenshot()
    screenshot.save(path)