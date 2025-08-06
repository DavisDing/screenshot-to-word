import tkinter as tk
from PIL import Image, ImageTk, ImageDraw
import os

def launch_annotator(image_path):
    Annotator(image_path).run()
    return image_path.replace(".png", "_marked.png")

class Annotator:
    def __init__(self, image_path):
        self.image_path = image_path
        self.root = tk.Toplevel()
        self.root.attributes('-fullscreen', True)
        self.root.title("截图标注器 - ESC退出")

        self.original = Image.open(image_path)
        self.image = self.original.copy()
        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas = tk.Canvas(self.root, width=self.image.width, height=self.image.height)
        self.canvas.pack()

        self.draw = ImageDraw.Draw(self.image)
        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.create_image(0, 0, anchor='nw', image=self.tk_image)
        self.canvas.bind('<ButtonPress-1>', self.on_left_down)
        self.canvas.bind('<B1-Motion>', self.on_left_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_left_up)
        self.canvas.bind('<Button-3>', self.on_right_click)
        self.root.bind('<Escape>', self.on_escape)
        tk.Button(self.root, text="✅ 保存退出", command=self.save_and_close).pack(pady=8)

    def on_left_down(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def on_left_drag(self, event):
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_oval(self.start_x, self.start_y, event.x, event.y, outline='red', width=2)

    def on_left_up(self, event):
        self.draw.ellipse([self.start_x, self.start_y, event.x, event.y], outline='red', width=2)
        self.update_image()

    def on_right_click(self, event):
        text = self.get_text_input()
        if text:
            self.draw.text((event.x, event.y), text, fill='blue')
            self.update_image()

    def get_text_input(self):
        input_win = tk.Toplevel(self.root)
        input_win.title("输入文字")
        input_box = tk.Entry(input_win)
        input_box.pack(padx=10, pady=10)
        result = {'text': None}

        def confirm():
            result['text'] = input_box.get()
            input_win.destroy()

        tk.Button(input_win, text="确定", command=confirm).pack()
        input_win.wait_window()
        return result['text']

    def on_escape(self, event):
        self.save_and_close()

    def update_image(self):
        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas.create_image(0, 0, anchor='nw', image=self.tk_image)

    def save_and_close(self):
        save_path = self.image_path.replace(".png", "_marked.png")
        self.image.save(save_path)
        self.root.destroy()

    def run(self):
        self.root.mainloop()
