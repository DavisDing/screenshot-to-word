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
        self.root = tk.Toplevel()
        self.root.title("截图标注 - ESC 或 点击保存退出")
        self.root.attributes("-topmost", True)
        # 自适应大小，不全屏
        window_width =  self._get_image_width(image_path) + 20
        window_height = self._get_image_height(image_path) + 60
        self.root.geometry(f"{window_width}x{window_height}+100+100")
        self.root.configure(bg='gray')
        self.root.resizable(False, False)

        self.original = Image.open(image_path)
        self.image = self.original.copy()
        self.draw = ImageDraw.Draw(self.image)

        self.tk_image = ImageTk.PhotoImage(self.image)
        self.canvas = tk.Canvas(self.root, width=self.image.width, height=self.image.height)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image, tags="image")
        self.canvas.pack(padx=10, pady=5)

        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.bind("<ButtonPress-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.root.bind("<Escape>", self.on_escape)

        # 嵌入顶部保存按钮
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
            try:
                self.image.save(save_path)
                print(f"✅ 标注已保存：{save_path}")
            except Exception as e:
                print(f"❌ 保存失败：{e}")
                messagebox.showerror("保存失败", f"保存标注时出错：{e}")
        self.root.destroy()

    def on_escape(self, event=None):
        # ESC 调用保存退出
        self.save_and_close()

    def run(self):
        self.root.mainloop()