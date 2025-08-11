# ui/annotator.py

import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os

class Annotator(tk.Toplevel):
    def __init__(self, root, image_path):
        super().__init__(root)
        self.root = root
        self.image_path = image_path
        self.title("截图标注")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.resizable(True, True)

        # 载入图片
        self.original_image = Image.open(self.image_path).convert("RGBA")
        self.draw_image = self.original_image.copy()
        self.tk_image = ImageTk.PhotoImage(self.draw_image)

        self.canvas = tk.Canvas(self, width=self.tk_image.width(), height=self.tk_image.height(), cursor="cross")
        self.canvas.pack(fill="both", expand=True)

        self.canvas_image = self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

        # 标注数据
        self.shapes = []  # (type, data)，type: 'circle'/'text'
        self.current_circle = None
        self.start_x = None
        self.start_y = None

        # 绑定事件
        self.canvas.bind("<ButtonPress-1>", self.on_left_button_down)
        self.canvas.bind("<B1-Motion>", self.on_left_button_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_button_up)
        self.canvas.bind("<Button-3>", self.on_right_click)

        self.bind("<Control-z>", self.on_undo)
        self.bind("<Control-Z>", self.on_undo)
        self.bind("<Control-s>", self.on_save)
        self.bind("<Control-S>", self.on_save)
        self.bind("<Escape>", self.on_escape)

        self.focus_set()
        self.save_result = False

    def on_left_button_down(self, event):
        self.start_x, self.start_y = event.x, event.y
        self.current_circle = self.canvas.create_oval(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2)

    def on_left_button_move(self, event):
        if self.current_circle:
            self.canvas.coords(self.current_circle, self.start_x, self.start_y, event.x, event.y)

    def on_left_button_up(self, event):
        if self.current_circle:
            coords = self.canvas.coords(self.current_circle)
            if abs(coords[2] - coords[0]) < 5 or abs(coords[3] - coords[1]) < 5:
                # 太小，视为无效，删除
                self.canvas.delete(self.current_circle)
            else:
                self.shapes.append(('circle', coords, self.current_circle))
            self.current_circle = None

    def on_right_click(self, event):
        text = simpledialog.askstring("输入文本", "请输入标注文字：", parent=self)
        if text:
            item = self.canvas.create_text(event.x, event.y, text=text, fill="blue", font=("Arial", 14))
            self.shapes.append(('text', (event.x, event.y, text), item))

    def on_undo(self, event=None):
        if not self.shapes:
            return
        last = self.shapes.pop()
        self.canvas.delete(last[2])

    def on_save(self, event=None):
        # 生成最终图像，叠加标注
        final_img = self.original_image.copy()
        draw = ImageDraw.Draw(final_img)

        # 尝试加载Arial字体，大小设置为16，如果失败则使用默认字体
        try:
            font = ImageFont.truetype("arial.ttf", 16)
        except IOError:
            font = ImageFont.load_default()

        for shape in self.shapes:
            if shape[0] == 'circle':
                x1, y1, x2, y2 = shape[1]
                draw.ellipse([x1, y1, x2, y2], outline="red", width=3)
            elif shape[0] == 'text':
                x, y, text = shape[1]
                draw.text((x, y), text, fill="blue", font=font)

        # 保存覆盖原文件
        final_img.save(self.image_path)
        self.save_result = True
        self.destroy()

    def on_escape(self, event=None):
        self.root.attributes('-topmost', True)
        if messagebox.askyesno("确认", "是否保存标注后退出？", parent=self):
            self.on_save()
        else:
            self.save_result = False
            self.destroy()
        self.root.attributes('-topmost', False)

    def on_close(self):
        self.save_result = False
        self.destroy()