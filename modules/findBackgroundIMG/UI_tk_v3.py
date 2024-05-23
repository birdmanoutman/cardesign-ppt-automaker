import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image
import csv
import os
import platform
import threading
from functools import partial


class ImageGalleryApp:
    def __init__(self, master, csv_file_path):
        self.master = master
        self.master.title("Image Gallery")
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        self.master.geometry(f"{int(screen_width * 0.8)}x{int(screen_height * 0.8)}+{int(screen_width * 0.1)}+{int(screen_height * 0.1)}")

        self.csv_file_path = csv_file_path
        self.image_cache = {}  # 图片缓存
        self.image_load_queue = []  # 待加载图片队列
        self.loading_threads = []  # 图片加载线程列表
        self.create_ui()

    def create_ui(self):
        self.canvas = tk.Canvas(self.master)
        self.scrollbar = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # 绑定滚动事件，用于懒加载
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.master.bind("<MouseWheel>", self.on_mouse_wheel)

        # 初始化时加载图片列表，但不实际加载图片内容
        self.load_image_list_from_csv()

    def load_image_list_from_csv(self):
        with open(self.csv_file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                img_path = row['Image File']
                self.image_load_queue.append(img_path)  # 将图片路径添加到加载队列

    def load_images_in_background(self):
        """在后台线程中加载图片。"""
        while self.image_load_queue:
            img_path = self.image_load_queue.pop(0)
            if img_path not in self.image_cache:
                thread = threading.Thread(target=self.load_image, args=(img_path,))
                thread.start()
                self.loading_threads.append(thread)

        # 等待所有加载线程完成
        for thread in self.loading_threads:
            thread.join()

    def load_image(self, img_path):
        """加载单个图片，并将其存储在缓存中。"""
        try:
            img = Image.open(img_path)
            img.thumbnail((200, 200), Image.LANCZOS)
            self.image_cache[img_path] = ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Error loading image: {img_path}, {e}")

    def on_frame_configure(self, event=None):
        """当滚动框架配置更改时调用，用于实现懒加载。"""
        # TODO: 基于当前滚动位置判断需要加载的图片，并调用load_images_in_background加载图片。
        self.load_images_in_background()

    def on_mouse_wheel(self, event):
        if platform.system() == "Darwin":  # macOS
            self.canvas.yview_scroll(int(-1 * (event.delta)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageGalleryApp(root, "//Users/birdmanoutman/上汽/backgroundIMGsource/image_ppt_mapping.csv")
    root.mainloop()
