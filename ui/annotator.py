# ui/annotator.py
import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import os

def launch_annotator(image_path):
    Annotator(image_path).run()
    return image_path.replace(".png", "_marked.png")

class Annotator:
    def __init__(self, image_path):
        self.image_path = image_path

        # 读取图片
        self.original = Image.open(image_path)
        self.image = self.original.copy()
        self.draw = ImageDraw.Draw(self.image)

        # 初始化窗口，使用自适应图片尺寸（非全屏）
        self.root = tk.Toplevel()
        self.root.title("截图标注")
        self.root.attributes("-topmost", True)
        window_width = self.image.width + 20
        window_height = self.image.height + 60
        self.root.geometry(f"{window_width}x{window_height}+100+100")
        self.root.configure(bg='gray')
        self.root.resizable(False, False)

        # 顶部工具栏：嵌入保存按钮
        toolbar = tk.Frame(self.root, bg='gray')
        toolbar.pack(side="top", fill="x", pady=5)

        save_btn = tk.Button(
            toolbar,
            text="✅ 保存标注并退出",
            command=self.save_and_close,
            bg="#4CAF50", fg="white",
            font=("Arial", 11, "bold")
        )
        save_btn.pack(side="right", padx=15)

        # 图像显示区域
        self.canvas = tk.Canvas(self.root, width=self.image.width, height=self.image.height)
        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image, tags="image")
        self.canvas.pack(padx=10, pady=5)

        # 标注操作绑定
        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.root.bind("<Escape>", self.save_and_close)

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
        self.update_canvas()
        self.rect = None

    def on_right_click(self, event):
        text = self.get_text_input()
        if text:
            self.draw.text((event.x, event.y), text, fill='blue')
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
        self.canvas.delete("image")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image, tags="image")

    def save_and_close(self, event=None):
        if messagebox.askyesno("退出确认", "是否保存标注并退出？"):
            save_path = self.image_path.replace(".png", "_marked.png")
            self.image.save(save_path)
            print(f"✅ 标注已保存：{save_path}")
        self.root.destroy()

    def run(self):
        self.root.mainloop()