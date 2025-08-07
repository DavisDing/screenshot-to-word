# ui/annotator.py

import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import os

class Annotator(tk.Toplevel):
    def __init__(self, root, image_path):
        # 确保存在默认 root（防止 ImageTk 报错）
        if not tk._default_root:
            _default_root = tk.Tk()
            _default_root.withdraw()
        super().__init__(root)
        self.root = root
        self.image_path = image_path
        self.title("截图标注")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.resizable(True, True)

        self.canvas = tk.Canvas(self, cursor="cross")
        self.canvas.pack(fill="both", expand=True)

        # 延迟加载图片，确保窗口初始化完成
        self.after(10, self.load_image)

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

        for shape in self.shapes:
            if shape[0] == 'circle':
                x1, y1, x2, y2 = shape[1]
                draw.ellipse([x1, y1, x2, y2], outline="red", width=3)
            elif shape[0] == 'text':
                x, y, text = shape[1]
                draw.text((x, y), text, fill="blue")

        # 保存覆盖原文件
        final_img.save(self.image_path)
        self.destroy()

    def on_escape(self, event=None):
        self.root.attributes('-topmost', True)
        if messagebox.askyesno("确认", "是否保存标注后退出？", parent=self):
            self.on_save()
        else:
            self.destroy()
        self.root.attributes('-topmost', False)

    def on_close(self):
        self.on_escape()
    def load_image(self):
        try:
            self.original_image = Image.open(self.image_path).convert("RGBA")
            self.draw_image = self.original_image.copy()

            def display_image():
                self.tk_image = ImageTk.PhotoImage(self.draw_image)
                self.canvas.config(width=self.tk_image.width(), height=self.tk_image.height())
                self.canvas_image = self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

            self.after(0, display_image)  # 确保在主循环中执行图像显示
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("图片加载失败", f"无法加载截图图片：\n{e}", parent=self))
            self.after(0, self.destroy)