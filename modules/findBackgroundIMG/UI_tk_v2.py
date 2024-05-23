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
        images = {}
        with open(self.csv_file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                img_hash = row['Image Hash']
                # 无论是否为重复项，都将PPTX链接添加到对应哈希值的列表中
                if img_hash not in images:
                    images[img_hash] = {'img_path': row['Image File'], 'pptx_paths': [row['PPTX File']]}
                else:
                    # 避免重复添加相同的PPTX文件路径
                    if row['PPTX File'] not in images[img_hash]['pptx_paths']:
                        images[img_hash]['pptx_paths'].append(row['PPTX File'])
        return list(images.values())

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
        frame.config(borderwidth=2, relief="solid", bg='black')  # 用黑色边框显示选中效果

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
        for image_data in self.images:
            img_path = image_data['img_path']
            pptx_paths = image_data['pptx_paths']
            try:
                img = Image.open(img_path)
                img.thumbnail((200, 200), resample_method)
                tk_img = ImageTk.PhotoImage(img)

                frame = tk.Frame(self.scrollable_frame, borderwidth=0, relief="flat")
                btn = tk.Button(frame, image=tk_img, relief="flat", borderwidth=0, bg='SystemButtonFace')
                btn.image = tk_img
                btn.pack(padx=5, pady=5)  # Adjust as needed to provide additional space within the Frame for the image

                # Bind single-click event for selecting the image (the whole frame)
                btn.bind("<Button-1>", lambda e, btn=frame: self.select_image(btn))

                # For each PPTX path, create a clickable label or button
                for pptx_path in pptx_paths:
                    pptx_btn = tk.Button(frame, text=os.path.basename(pptx_path), command=lambda p=pptx_path: self.open_pptx(p))
                    pptx_btn.pack()  # Place PPTX link buttons inside the Frame, right below the image

                frame.grid(row=len(self.image_buttons) // 3, column=len(self.image_buttons) % 3, padx=10, pady=10)

                self.image_buttons.append((frame, img_path, pptx_paths))  # Store the entire Frame containing the image and links
            except Exception as e:
                print(f"Error loading image: {img_path}, {e}")

    def on_window_resize(self, event):
        """窗口大小改变时的处理函数。"""
        self.arrange_images()

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageGalleryApp(root, "/Users/birdmanoutman/上汽/backgroundIMGsource/image_ppt_mapping.csv")
    root.mainloop()
