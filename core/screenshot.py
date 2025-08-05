import pyautogui
import tempfile
import os

def take_screenshot():
    screenshot = pyautogui.screenshot()
    temp_file = os.path.join(tempfile.gettempdir(), 'screenshot.png')
    screenshot.save(temp_file)
    return temp_file