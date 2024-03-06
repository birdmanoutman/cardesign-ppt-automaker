import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import csv
import os
import subprocess
import platform
from PIL import Image, __version__ as PIL_version

if int(PIL_version.split('.')[0]) >= 9:  # Pillow 9.0.0及以上版本
    resample_method = Image.LANCZOS
else:
    resample_method = Image.ANTIALIAS




class ImageGalleryApp:
    def __init__(self, master, csv_file_path):
        self.master = master
        self.master.title("Image Gallery")
        self.csv_file_path = csv_file_path
        self.images = self.load_images_from_csv()
        self.create_ui()

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

        # Improved scrolling support on macOS and Windows
        # 移除了 "<Touch>" 事件绑定，因为它在一些环境中不被支持
        self.master.bind_all("<MouseWheel>", self._on_mousewheel)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.populate()

    def populate(self):
        rows, columns = 0, 0
        for img_path, pptx_path in self.images:
            if columns == 3:  # Adjust based on desired number of columns
                columns = 0
                rows += 1
            try:
                img = Image.open(img_path)
                img.thumbnail((200, 200), resample_method)  # Adjust thumbnail size as needed
                tk_img = ImageTk.PhotoImage(img)
                btn = tk.Button(self.scrollable_frame, image=tk_img, command=lambda p=pptx_path: self.open_pptx(p))
                btn.image = tk_img
                btn.grid(row=rows, column=columns, padx=10, pady=10)
                columns += 1
            except Exception as e:
                print(f"Error loading image: {img_path}, {e}")

    def _on_mousewheel(self, event):
        if platform.system() == "Darwin":  # macOS
            self.canvas.yview_scroll(int(-1 * (event.delta)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def open_pptx(self, pptx_path):
        try:
            if platform.system() == "Windows":
                os.startfile(pptx_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.call(['open', pptx_path])
            else:  # Assume Linux or other
                subprocess.call(['xdg-open', pptx_path])
        except Exception as e:
            print(f"Failed to open file: {pptx_path}, {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageGalleryApp(root, "/Users/birdmanoutman/Desktop/outputTest/image_ppt_mapping.csv")
    root.mainloop()
