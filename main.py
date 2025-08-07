import threading
import tkinter as tk
import webbrowser
from utils.path_utils import get_base_path, ensure_directories
from utils.logger import Logger
from utils.excel_handler import ExcelHandler
from core.test_runner import TestRunner

APP_TITLE = "测试工具 - V1.0.0 作者: DingHao"

def show_splash_screen():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.geometry("300x150+600+300")
    splash.configure(bg="white")
    tk.Label(splash, text="正在启动测试工具...", font=("Arial", 14), bg="white").pack(expand=True)
    splash.update()
    return splash

class DesktopTestToolApp:
    def __init__(self):
        import os
        self.base_path = get_base_path()
        os.chdir(self.base_path)  # 设置工作目录为exe同级路径，确保所有相对路径正确

        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("450x400")
        self.root.resizable(True, True)

        self.logger = Logger(self.root)

        try:
            ensure_directories(self.base_path, self.logger)
        except Exception as e:
            self.logger.log(f"初始化目录异常：{e}")

        self.excel_handler = ExcelHandler(self.logger, input_dir=self.base_path + "/excel_input", root=self.root)
        self.test_runner = TestRunner(self.logger, self.excel_handler, self.root)

        self.create_main_ui()
        self.root.after(100, lambda: self.logger.log("程序启动，界面加载完成"))

    def create_main_ui(self):

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)

        start_btn = tk.Button(btn_frame, text="开始执行", width=15, command=self.on_start)
        start_btn.pack(side="left", padx=10)

        exit_btn = tk.Button(btn_frame, text="退出", width=15, command=self.on_exit)
        exit_btn.pack(side="left", padx=10)
        
        tk.Label(self.root, text=APP_TITLE, font=("Arial", 10), wraplength=380, justify="left").pack(pady=10, fill="x")

        link = tk.Label(self.root, text="GitHub", fg="blue", cursor="hand2", font=("Arial", 10))
        link.pack(pady=0)
        link.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/DavisDing/screenshot-to-word"))
    def on_start(self):
        self.logger.log("点击开始执行按钮")
        # 每次开始前重置 TestRunner 内部状态，确保 event、index 等状态干净
        self.test_runner.reset()
        thread = threading.Thread(target=self.test_runner.run_tests)
        thread.daemon = True
        thread.start()

    def on_exit(self):
        self.logger.log("点击退出按钮，程序结束")
        self.root.quit()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    splash = show_splash_screen()
    app = DesktopTestToolApp()
    splash.destroy()
    app.run()