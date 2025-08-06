# ui/annotator.py
import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import os

class Annotator(tk.Toplevel):
    def __init__(self, master, image_path, save_callback):
        super().__init__(master)
        self.title("截图标注")
        self.attributes("-topmost", True)

        self.image_path = image_path
        self.save_callback = save_callback
        self.orig_image = Image.open(image_path)
        self.draw_image = self.orig_image.copy()
        self.tk_image = ImageTk.PhotoImage(self.draw_image)

        self.canvas = tk.Canvas(self, width=self.tk_image.width(), height=self.tk_image.height(), cursor="cross")
        self.canvas.pack()

        self.canvas_image = self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

        self.shapes = []  # 存储画圈和文字的指令
        self.current_circle = None
        self.start_x = None
        self.start_y = None

        self.canvas.bind("<ButtonPress-1>", self.on_left_button_down)
        self.canvas.bind("<B1-Motion>", self.on_left_button_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_button_up)

        self.canvas.bind("<Button-3>", self.on_right_button)

        self.bind("<Escape>", self.on_escape)
        self.bind("<Control-z>", self.on_undo)
        self.bind("<Control-s>", self.on_save)

    def on_left_button_down(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.current_circle = self.canvas.create_oval(self.start_x, self.start_y, event.x, event.y, outline="red", width=2)

    def on_left_button_move(self, event):
        if self.current_circle:
            self.canvas.coords(self.current_circle, self.start_x, self.start_y, event.x, event.y)

    def on_left_button_up(self, event):
        if self.current_circle:
            coords = self.canvas.coords(self.current_circle)
            self.shapes.append(('oval', coords))
            self.current_circle = None

    def on_right_button(self, event):
        text = simpledialog.askstring("文本标注", "请输入文本：", parent=self)
        if text:
            text_id = self.canvas.create_text(event.x, event.y, text=text, fill="blue", font=("Arial", 14, "bold"))
            self.shapes.append(('text', (event.x, event.y, text)))

    def on_escape(self, event):
        if messagebox.askyesno("退出确认", "是否放弃当前标注并退出？"):
            self.destroy()

    def on_undo(self, event):
        if not self.shapes:
            return
        last = self.shapes.pop()
        if last[0] == 'oval':
            # 找到最近的oval并删除
            items = self.canvas.find_all()
            for item in reversed(items):
                if self.canvas.type(item) == 'oval':
                    self.canvas.delete(item)
                    break
        elif last[0] == 'text':
            # 删除最后一个文本对象
            items = self.canvas.find_all()
            for item in reversed(items):
                if self.canvas.type(item) == 'text':
                    self.canvas.delete(item)
                    break

    def on_save(self, event=None):
        # 将标注绘制到图片上
        draw = ImageDraw.Draw(self.draw_image)
        for shape in self.shapes:
            if shape[0] == 'oval':
                draw.ellipse(shape[1], outline="red", width=3)
            elif shape[0] == 'text':
                x, y, text = shape[1]
                draw.text((x, y), text, fill="blue")
        save_path = self.image_path  # 也可以改成新路径
        self.draw_image.save(save_path)
        if self.save_callback:
            self.save_callback(save_path)
        self.destroy()