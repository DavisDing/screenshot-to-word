import os
import time
import datetime
import pandas as pd
import pyautogui
import keyboard
import glob
import re
import tempfile
import uuid
import threading
from tkinter import Tk, Canvas, simpledialog, Button, Frame, Label, messagebox, Text
from PIL import Image, ImageTk
from docx import Document
from docx.shared import Inches
from openpyxl import load_workbook

# 目录与日志文件
INPUT_DIR = 'excel_input'
OUTPUT_DIR = 'output'
EXECUTED_STATUS = '已执行'
os.makedirs(OUTPUT_DIR, exist_ok=True)
LOG_FILE = os.path.join(OUTPUT_DIR, f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

# 自动读取最新Excel文件
xlsx_files = sorted(glob.glob(os.path.join(INPUT_DIR, '*.xlsx')), key=os.path.getmtime, reverse=True)
if not xlsx_files:
    raise FileNotFoundError("在 ./excel_input/ 下未找到 Excel 文件")
EXCEL_PATH = xlsx_files[0]

# 加载Excel
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
df.fillna('', inplace=True)

def create_floating_screenshot_button():
    def do_capture():
        try:
            if 'screenshot_triggered' in globals():
                screenshot_triggered.set()
                return
            messagebox.showwarning("请先开始执行", "请点击主界面“开始执行”以进入用例截图流程。")
        except Exception as e:
            messagebox.showerror("截图失败", str(e))

    float_root = Tk()
    float_root.title("截图按钮")
    float_root.geometry("100x50+50+50")
    float_root.attributes("-topmost", True)
    float_root.resizable(False, False)
    Button(float_root, text="截图", command=do_capture).pack(expand=True, fill='both')
    threading.Thread(target=float_root.mainloop, daemon=True).start()

# 日志窗口
class LogWindow:
    def __init__(self):
        self.root = Tk()
        self.root.title("操作日志")
        self.root.geometry("400x300+100+100")
        self.text = Text(self.root, wrap='word', bg='black', fg='lime', font=("Consolas", 10))
        self.text.pack(expand=True, fill='both')
        self.root.attributes('-topmost', True)
        self.log_lines = []
        # Add screenshot button at the end of __init__
        Button(self.root, text='截图', command=lambda: screenshot_triggered.set()).pack(pady=5)
        self.root.after(100, self.start_mainloop)

    def start_mainloop(self):
        try:
            self.root.mainloop()
        except Exception:
            pass

    def log(self, msg):
        timestamp = time.strftime("[%H:%M:%S]")
        line = f"{timestamp} {msg}"
        self.text.insert('end', line + '\n')
        self.text.see('end')
        self.log_lines.append(line)
        try:
            with open(LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(line + '\n')
        except Exception as e:
            print(f"日志写入失败: {e}")

# 标注工具
class Annotator:
    def __init__(self, image_path, save_path, word_path, case_text):
        self.root = Tk()
        self.root.title("截图标注工具")
        self.image = Image.open(image_path)
        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas = Canvas(self.root, width=self.image.width, height=self.image.height)
        self.canvas.pack()
        self.canvas.create_image(0, 0, anchor='nw', image=self.tk_image)

        self.start_x = None
        self.start_y = None
        self.drawn_objects = []
        self.text_mode = False
        self.save_path = save_path
        self.word_path = word_path
        self.case_text = case_text

        self.canvas.bind('<ButtonPress-1>', self.on_click)
        self.canvas.bind('<B1-Motion>', self.on_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)
        self.canvas.bind('<Button-3>', self.on_right_click)

        ctrl = Frame(self.root)
        ctrl.pack()
        Button(ctrl, text='撤销', command=self.undo).pack(side='left')
        Button(ctrl, text='保存截图并退出', command=self.save_and_exit).pack(side='left')
        self.root.mainloop()

    def on_click(self, event):
        if self.text_mode:
            return
        self.start_x = event.x
        self.start_y = event.y
        self.current_shape = self.canvas.create_oval(event.x, event.y, event.x, event.y, outline='red', width=2)
        self.drawn_objects.append(self.current_shape)

    def on_drag(self, event):
        if self.text_mode:
            return
        self.canvas.coords(self.current_shape, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        pass

    def on_right_click(self, event):
        self.text_mode = True
        text = simpledialog.askstring("输入文字", "请输入标注内容：")
        if text:
            label = self.canvas.create_text(event.x, event.y, text=text, fill="blue", font=("Arial", 14, "bold"))
            self.drawn_objects.append(label)
        self.text_mode = False

    def undo(self):
        if self.drawn_objects:
            self.canvas.delete(self.drawn_objects.pop())

    def save_and_exit(self):
        temp_eps_fd, temp_eps_path = tempfile.mkstemp(suffix='.eps')
        try:
            self.canvas.postscript(file=temp_eps_path)
            img = Image.open(temp_eps_path)
            img.save(self.save_path)
        finally:
            os.close(temp_eps_fd)
            if os.path.exists(temp_eps_path):
                os.remove(temp_eps_path)

        if os.path.exists(self.word_path):
            doc = Document(self.word_path)
        else:
            doc = Document()
            doc.add_heading(os.path.splitext(os.path.basename(self.word_path))[0], level=1)
            doc.add_paragraph(self.case_text)

        doc.add_picture(self.save_path, width=Inches(6))
        doc.add_paragraph('')
        doc.save(self.word_path)
        self.root.destroy()

# 主流程
def run():
    global screenshot_triggered
    screenshot_triggered = threading.Event()
    create_floating_screenshot_button()
    logwin = LogWindow()
    logwin.log(f"开始处理 Excel 文件：{EXCEL_PATH}")
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    for idx, row in df.iterrows():
        if idx == 0:
            continue
        file_name = str(row[0]).strip()
        verify_point = str(row[1]).strip()
        status = str(row[2]).strip()
        if re.search(r'[<>:"/\\|?*\x00-\x1F]', file_name):
            messagebox.showerror("文件名错误", f"无效的文件名: {file_name}")
            continue
        safe_file_name = os.path.normpath(file_name).replace(os.sep, '_')

        # 截图文件夹（以A列命名的子文件夹）
        raw_dir = os.path.join(OUTPUT_DIR, safe_file_name)
        os.makedirs(raw_dir, exist_ok=True)

        # Word文件直接放output根目录
        word_path = os.path.join(OUTPUT_DIR, f"{safe_file_name}.docx")

        if status.lower() in ['passed', EXECUTED_STATUS]:
            continue

        logwin.log(f"[用例名称] {file_name}")
        logwin.log(f"[验证点] {verify_point}")
        messagebox.showinfo("验证点", verify_point)

        while True:
            screenshot_triggered.clear()
            keyboard.add_hotkey('F8', lambda: screenshot_triggered.set())
            logwin.log("等待按 F8 或点击悬浮截图按钮开始截图...")
            while not screenshot_triggered.is_set():
                time.sleep(0.1)
            messagebox.showinfo("截图中", "即将截图当前屏幕...")
            logwin.log("已触发截图，开始截图...")
            time.sleep(0.5)

            try:
                screenshot = pyautogui.screenshot()
            except Exception as e:
                messagebox.showerror("截图错误", f"截图失败: {str(e)}")
                continue

            count = len(glob.glob(os.path.join(raw_dir, "screenshot_*.png"))) + 1
            img_path = os.path.join(raw_dir, f"screenshot_{count}.png")
            tmp_path = os.path.join(raw_dir, f"temp_{uuid.uuid4().hex}.png")
            screenshot.save(tmp_path)

            Annotator(tmp_path, img_path, word_path, verify_point)
            os.remove(tmp_path)

            ws.cell(row=idx+2, column=3, value=EXECUTED_STATUS)
            try:
                wb.save(EXCEL_PATH)
            except Exception as e:
                messagebox.showerror("保存错误", f"Excel保存失败: {str(e)}")
                continue

            choice = messagebox.askyesno("继续截图？", "是否继续截图？\n是 = 当前用例继续截图\n否 = 进入下一条")
            if not choice:
                break

    messagebox.showinfo("执行完成", "所有用例执行完毕！")

# 主界面
class MainWindow:
    def __init__(self):
        self.root = Tk()
        self.root.title("截图自动化工具")
        self.root.geometry("300x150+200+200")
        Label(self.root, text="截图自动化工具", font=("Arial", 14)).pack(pady=10)
        Button(self.root, text="开始执行", command=self.start).pack(pady=10)
        Button(self.root, text="截图").pack(pady=10)
        Button(self.root, text="退出", command=self.root.quit).pack()
        self.root.mainloop()

    def start(self):
        create_floating_screenshot_button()
        self.root.destroy()
        messagebox.showinfo("启动提示", "截图工具已启动，将开始第一条未完成验证点。\n请按 F8 截图。")

        run()

if __name__ == '__main__':
    MainWindow()