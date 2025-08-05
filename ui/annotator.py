import tkinter as tk
from PIL import Image, ImageTk, ImageDraw

class Annotator(tk.Toplevel):
    def __init__(self, image_path, on_save):
        super().__init__()
        self.title("标注截图")
        self.attributes("-fullscreen", True)
        self.image = Image.open(image_path)
        self.draw = ImageDraw.Draw(self.image)
        self.canvas = tk.Canvas(self)
        self.canvas.pack(fill="both", expand=True)
        self.tk_img = ImageTk.PhotoImage(self.image)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        self.bind("<Escape>", lambda e: self.destroy())
        self.canvas.bind("<B1-Motion>", self.paint)
        self.canvas.bind("<Button-3>", self.write_text)
        self.on_save = on_save

    def paint(self, event):
        x, y = event.x, event.y
        r = 20
        self.canvas.create_oval(x-r, y-r, x+r, y+r, outline="red", width=3)
        self.draw.ellipse((x-r, y-r, x+r, y+r), outline="red", width=3)

    def write_text(self, event):
        text = "标注"
        self.canvas.create_text(event.x, event.y, text=text, fill="blue", font=("Arial", 14))
        self.draw.text((event.x, event.y), text, fill="blue")

    def destroy(self):
        path = os.path.join("temp", "annotated.png")
        os.makedirs("temp", exist_ok=True)
        self.image.save(path)
        self.on_save(path)
        super().destroy()