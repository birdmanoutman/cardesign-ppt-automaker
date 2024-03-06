import csv
import os
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import subprocess
import platform

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
        self.canvas = tk.Canvas(self.master, borderwidth=0)
        self.frame = tk.Frame(self.canvas)
        self.vsb = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((4,4), window=self.frame, anchor="nw", tags="self.frame")

        self.frame.bind("<Configure>", self.on_frame_configure)
        self.populate()

    def populate(self):
        for img_path, pptx_path in self.images:
            try:
                img = Image.open(img_path)
                img.thumbnail((100, 100), Image.ANTIALIAS)
                tk_img = ImageTk.PhotoImage(img)
                btn = tk.Button(self.frame, image=tk_img)
                btn.image = tk_img
                btn.bind("<Double-1>", lambda e, p=pptx_path: self.open_pptx(p))  # Double-click to open
                btn.pack(side="top", fill="both", expand=True)
            except Exception as e:
                print(f"Error loading image: {img_path}, {e}")

    def on_frame_configure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

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
