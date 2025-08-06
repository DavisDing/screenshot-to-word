# ui/annotator.py
import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import os

def launch_annotator(image_path):
    if not os.path.exists(image_path):
        print("[标注器] 图片不存在：", image_path)
        return None

    try:
        img = Image.open(image_path)
        img.load()
    except Exception as e:
        print("[标注器] 图像加载失败：", e)
        return None

    saved = False
    annotated_path = image_path.replace(".png", "_marked.png")
    undo_stack = []
    annotations = []

    root = tk.Tk()
    root.title("截图标注 - Esc退出，Ctrl+S保存")
    root.geometry(f"{img.width+20}x{img.height+20}+100+100")
    root.resizable(True, True)

    draw = ImageDraw.Draw(img)
    tk_img = ImageTk.PhotoImage(img)

    canvas = tk.Canvas(root, width=img.width, height=img.height, bg="gray")
    canvas.pack(padx=10, pady=10)

    canvas.create_image(0, 0, anchor="nw", image=tk_img)
    canvas.image = tk_img  # 防止 pyimage2 错误

    def save_and_exit(event=None):
        nonlocal saved
        saved = True
        for item in annotations:
            if item[0] == "oval":
                draw.ellipse(item[1], outline="red", width=2)
            elif item[0] == "text":
                x, y, text = item[1]
                draw.text((x, y), text, fill="blue")
        img.save(annotated_path)
        root.destroy()
        if hasattr(root, "float_btn"):
            root.float_btn.destroy()

    def undo(event=None):
        if undo_stack:
            last = undo_stack.pop()
            canvas.delete(last)
            if annotations:
                annotations.pop()

    def on_escape(event=None):
        if not saved:
            if messagebox.askyesno("退出", "尚未保存标注，是否保存？"):
                save_and_exit()
            else:
                root.destroy()
                if hasattr(root, "float_btn"):
                    root.float_btn.destroy()
        else:
            root.destroy()
            if hasattr(root, "float_btn"):
                root.float_btn.destroy()

    def on_left_press(event):
        canvas.start_x = event.x
        canvas.start_y = event.y

    def on_left_release(event):
        x0, y0 = canvas.start_x, canvas.start_y
        x1, y1 = event.x, event.y
        oval = canvas.create_oval(x0, y0, x1, y1, outline="red", width=2)
        annotations.append(("oval", (x0, y0, x1, y1)))
        undo_stack.append(oval)

    def on_right_click(event):
        text = simpledialog.askstring("添加文字", "请输入标注文字：")
        if text:
            txt = canvas.create_text(event.x, event.y, text=text, fill="blue", font=("Arial", 14, "bold"), anchor="nw")
            annotations.append(("text", (event.x, event.y, text)))
            undo_stack.append(txt)

    # ⚠️ 全局绑定快捷键 + 强制焦点
    root.focus_force()
    root.bind_all("<Control-s>", save_and_exit)
    root.bind_all("<Control-z>", undo)
    root.bind_all("<Escape>", on_escape)
    root.bind("<ButtonPress-1>", on_left_press)
    root.bind("<ButtonRelease-1>", on_left_release)
    root.bind("<Button-3>", on_right_click)

    # ✅ 悬浮保存按钮（延迟显示）
    def create_float_button():
        float_btn = tk.Toplevel()
        float_btn.overrideredirect(True)
        float_btn.attributes('-topmost', True)
        float_btn.geometry("+30+30")

        save_btn = tk.Button(float_btn, text="保存标注 (Ctrl+S)", command=save_and_exit)
        save_btn.pack()

        # 可拖动按钮
        def start_move(event):
            float_btn.x = event.x
            float_btn.y = event.y
        def stop_move(event):
            float_btn.x = None
            float_btn.y = None
        def on_motion(event):
            x = float_btn.winfo_x() - float_btn.x + event.x
            y = float_btn.winfo_y() - float_btn.y + event.y
            float_btn.geometry(f"+{x}+{y}")
        save_btn.bind("<ButtonPress-1>", start_move)
        save_btn.bind("<ButtonRelease-1>", stop_move)
        save_btn.bind("<B1-Motion>", on_motion)

        root.float_btn = float_btn
        print("[调试] 悬浮保存按钮已创建")

    root.after(300, create_float_button)
    root.mainloop()

    return annotated_path if saved else None