import pyautogui
import time

def take_screenshot(save_path):
    time.sleep(2)
    image = pyautogui.screenshot()
    image.save(save_path)