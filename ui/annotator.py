import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import threading

def launch_annotator(image_path, callback=None):
    def run_annotator():
        annotator = Annotator(image_path)
        annotator.run()
        # 保存后调用insert_case_image函数
        from utils.word_generator import WordGenerator
        import os
        # 获取基础目录和用例名称
        base_dir = os.path.dirname(os.path.dirname(image_path))
        case_name = os.path.basename(image_path).split('_')[0] 
        case_desc = "截图验证"
        
        # 创建一个简单的logger模拟
        class SimpleLogger:
            def write(self, msg):
                print(msg)
        
        # 调用insert_case_image函数
        word_gen = WordGenerator(base_dir, SimpleLogger())
        marked_image_path = image_path.replace(".png", "_marked.png")
        word_gen.insert_case_image(case_name, case_desc, marked_image_path)
        
        if callback:
            callback(marked_image_path)
    
    # 在新线程中运行标注器，避免阻塞主线程
    thread = threading.Thread(target=run_annotator, daemon=True)
    thread.start()

class Annotator:
    def __init__(self, image_path):
        self.image_path = image_path
        self.root = tk.Toplevel()
        self.root.title("截图标注 - ESC 或 点击保存退出")
        self.root.attributes("-topmost", True)
        
        # 获取图片尺寸并自适应窗口大小
        with Image.open(image_path) as img:
            self.img_width, self.img_height = img.size
        
        # 根据图片大小自适应窗口，设置最大和最小尺寸
        window_width = min(max(self.img_width + 20, 400), 1200)
        window_height = min(max(self.img_height + 60, 300), 800)
        
        self.root.geometry(f"{window_width}x{window_height}+100+100")
        self.root.configure(bg='gray')
        # 允许窗口调整大小
        self.root.resizable(True, True)

        self.original = Image.open(image_path)
        self.image = self.original.copy()
        self.draw = ImageDraw.Draw(self.image)

        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas = tk.Canvas(self.root, width=min(self.img_width, window_width-20), height=min(self.img_height, window_height-60))
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image, tags="image")
        self.canvas.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.actions = []

        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.root.bind("<Escape>", self.on_escape)
        self.root.bind("<Control-z>", self.on_undo)
        self.root.bind("<Control-s>", self.save_and_close)

        toolbar = tk.Frame(self.root, bg='gray')
        toolbar.pack(side="top", fill="x", pady=2)
        save_btn = tk.Button(
            toolbar,
            text="✅ 保存标注并退出",
            command=self.save_and_close,
            bg="#4CAF50", fg="white",
            font=("Arial", 12, "bold")
        )
        save_btn.pack(side="right", padx=10, pady=2)
        
        # 添加说明文本，让用户知道如何保存
        instruction_label = tk.Label(
            toolbar, 
            text="右键添加文字 | 拖拽绘制圆圈 | Ctrl+Z撤销 | Ctrl+S保存 | ESC退出",
            bg='gray',
            fg='white',
            font=("Arial", 9)
        )
        instruction_label.pack(side="left", padx=5, pady=2)

    def _get_image_width(self, path):
        with Image.open(path) as img:
            return img.width

    def _get_image_height(self, path):
        with Image.open(path) as img:
            return img.height

    def on_left_down(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def on_left_drag(self, event):
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_oval(
            self.start_x, self.start_y, event.x, event.y,
            outline='red', width=2
        )

    def on_left_up(self, event):
        self.draw.ellipse([self.start_x, self.start_y, event.x, event.y], outline='red', width=2)
        self.actions.append(('ellipse', (self.start_x, self.start_y, event.x, event.y)))
        self.update_canvas()
        self.rect = None

    def on_right_click(self, event):
        text = self.get_text_input()
        if text:
            self.draw.text((event.x, event.y), text, fill='blue')
            self.actions.append(('text', (event.x, event.y, text)))
            self.update_canvas()

    def get_text_input(self):
        win = tk.Toplevel(self.root)
        win.title("输入文字")
        entry = tk.Entry(win)
        entry.pack(padx=10, pady=10)
        result = {'text': None}
        def confirm():
            result['text'] = entry.get()
            win.destroy()
        tk.Button(win, text="确定", command=confirm).pack()
        win.wait_window()
        return result['text']

    def update_canvas(self):
        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image, tags="image")

    def save_and_close(self, event=None):
        save_path = self.image_path.replace(".png", "_marked.png")
        self.image.save(save_path)
        self.root.destroy()

    def on_escape(self, event=None):
        if messagebox.askokcancel("退出标注", "是否保存当前标注？"):
            self.save_and_close()
        else:
            self.root.destroy()

    def on_undo(self, event=None):
        if self.actions:
            self.image = self.original.copy()
            self.draw = ImageDraw.Draw(self.image)
            self.actions.pop()
            for act in self.actions:
                if act[0] == 'ellipse':
                    self.draw.ellipse(act[1], outline='red', width=2)
                elif act[0] == 'text':
                    x, y, text = act[1]
                    self.draw.text((x, y), text, fill='blue')
            self.update_canvas()

    def run(self):
        self.root.mainloop()