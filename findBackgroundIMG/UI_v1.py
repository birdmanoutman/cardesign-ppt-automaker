import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image, __version__ as PIL_version
import csv
import os
import subprocess
import platform
import threading


if int(PIL_version.split('.')[0]) >= 9:  # Pillow 9.0.0及以上版本
    resample_method = Image.LANCZOS
else:
    resample_method = Image.ANTIALIAS


class ImageGalleryApp:
    def __init__(self, master, csv_file_path):
        self.master = master
        self.master.title("Image Gallery")
        # 根据屏幕大小调整窗口大小
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        self.master.geometry(f"{int(screen_width * 0.8)}x{int(screen_height * 0.8)}+{int(screen_width * 0.1)}+{int(screen_height * 0.1)}")
        self.csv_file_path = csv_file_path
        self.images = self.load_images_from_csv()
        self.image_buttons = []  # 初始化image_buttons列表
        self.selected_image = None  # 用于跟踪当前选中的图片
        self.create_ui()  # create_ui现在负责调用populate_initial
        self.master.bind("<Configure>", self.on_window_resize)  # 绑定窗口大小变化事件

    def load_images_from_csv(self):
        images = []
        with open(self.csv_file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                images.append((row['Image File'], row['PPTX File']))
        return images

    def create_ui(self):
        self.canvas = tk.Canvas(self.master)
        self.scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.master.bind_all("<MouseWheel>", self._on_mousewheel)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.populate()
        self.master.bind("<Configure>", self.on_window_resize)

    def populate_initial(self):
        rows, columns = 0, 0
        for img_path, pptx_path in self.images:
            if columns == 3:  # Adjust based on desired number of columns
                columns = 0
                rows += 1
            try:
                img = Image.open(img_path)
                img.thumbnail((200, 200), Image.LANCZOS)  # Adjust thumbnail size as needed, using LANCZOS for high quality
                tk_img = ImageTk.PhotoImage(img)
                btn = tk.Button(self.scrollable_frame, image=tk_img)
                btn.image = tk_img
                btn.grid(row=rows, column=columns, padx=10, pady=10)
                columns += 1
                # Bind double-click event to open PPTX file in a background thread
                btn.bind("<Double-1>", lambda e, p=pptx_path: self.open_pptx(p))
                # Bind single-click event for selecting the image
                btn.bind("<Button-1>", lambda e, btn=btn: self.select_image(btn))
            except Exception as e:
                print(f"Error loading image: {img_path}, {e}")

    def _on_mousewheel(self, event):
        if platform.system() == "Darwin":  # macOS
            self.canvas.yview_scroll(int(-1 * (event.delta)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def open_pptx(self, pptx_path):
        """Open PPTX file in a background thread to improve responsiveness."""

        def run():
            try:
                if platform.system() == "Windows":
                    os.startfile(pptx_path)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.call(['open', pptx_path])
                else:  # Assume Linux or other
                    subprocess.call(['xdg-open', pptx_path])
            except Exception as e:
                print(f"Failed to open file: {pptx_path}, {e}")

        threading.Thread(target=run).start()

    def select_image(self, frame):
        # 重置之前选中的Frame的样式
        if self.selected_image:
            self.selected_image.config(borderwidth=0, relief="flat", bg='SystemButtonFace')  # Reset to default

        # 更新当前选中的Frame的样式以显示选中效果
        frame.config(borderwidth=2, relief="solid", bg='#d9d9d9')  # 设置边框宽度和样式

        self.selected_image = frame

    def arrange_images(self):
        rows, columns = 0, 0
        max_columns = self.calculate_columns()  # Calculate max columns based on current window width
        for btn, img_path, pptx_path in self.image_buttons:
            if columns >= max_columns:
                columns = 0
                rows += 1
            btn.grid(row=rows, column=columns, padx=10, pady=10)  # Use dynamic padding if necessary
            columns += 1

    def calculate_columns(self):
        """根据窗口宽度计算可以容纳的列数。"""
        width = self.master.winfo_width()
        column_width = 200 + 10 * 2  # 假定每个图片按钮的宽度为200像素，两边的pad为10
        columns = width // column_width
        return max(1, columns)  # 至少保持一列

    def populate(self):
        for img_path, pptx_path in self.images:
            try:
                img = Image.open(img_path)
                img.thumbnail((200, 200), resample_method)
                tk_img = ImageTk.PhotoImage(img)

                # 创建一个Frame作为图片按钮的容器
                frame = tk.Frame(self.scrollable_frame, borderwidth=0, relief="flat")
                btn = tk.Button(frame, image=tk_img, relief="flat", borderwidth=0, bg='SystemButtonFace')
                btn.image = tk_img  # 保持对图像对象的引用
                btn.pack(padx=5, pady=5)  # 根据需要调整，以在Frame内部为图片提供额外的空间

                frame.grid(row=len(self.image_buttons) // 3, column=len(self.image_buttons) % 3, padx=10, pady=10)

                btn.bind("<Double-1>", lambda e, p=pptx_path: self.open_pptx(p))
                btn.bind("<Button-1>", lambda e, btn=frame: self.select_image(btn))  # 将Frame作为参数传递给select_image

                self.image_buttons.append((frame, img_path, pptx_path))  # 存储Frame而非按钮本身
            except Exception as e:
                print(f"Error loading image: {img_path}, {e}")

    def on_window_resize(self, event):
        """窗口大小改变时的处理函数。"""
        self.arrange_images()

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageGalleryApp(root, "/Users/birdmanoutman/Desktop/outputTest/image_ppt_mapping.csv")
    root.mainloop()
