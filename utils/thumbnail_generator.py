import os
import tempfile
from PIL import Image, ImageTk, ImageOps
import fitz  # PyMuPDF

class ThumbnailGenerator:
    def __init__(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        # 创建一个空白图像，用于无预览的情况
        self.no_preview_image = Image.new('RGB', (50, 50), color='#D3D3D3')  # 灰色背景
        self.no_preview_photo = ImageTk.PhotoImage(self.no_preview_image)

    def create_thumbnail(self, filepath, label):
        file_ext = os.path.splitext(filepath)[1].lower()
        if file_ext in ['.jpg', '.jpeg', '.png', '.tiff', '.webp']:
            self.generate_image_thumbnail(filepath, label)
        elif file_ext == '.pdf':
            self.generate_pdf_thumbnail(filepath, label)
        else:
            self.set_no_preview(label)

    def generate_image_thumbnail(self, filepath, label):
        try:
            img = Image.open(filepath)
            img.thumbnail((50, 50))
            img = ImageOps.pad(img, (50, 50), method=Image.Resampling.LANCZOS)
            temp_file = os.path.join(self.temp_dir.name, os.path.basename(filepath) + '_thumb.png')
            img.save(temp_file, format='PNG')
            img = ImageTk.PhotoImage(file=temp_file)
            label.config(image=img, text='', width=50, height=50)
            label.image = img
        except Exception:
            self.set_no_preview(label)

    def generate_pdf_thumbnail(self, filepath, label):
        try:
            doc = fitz.open(filepath)
            page = doc.load_page(0)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((50, 50))
            img = ImageOps.pad(img, (50, 50), method=Image.Resampling.LANCZOS)
            temp_file = os.path.join(self.temp_dir.name, os.path.basename(filepath) + '_thumb.png')
            img.save(temp_file, format='PNG')
            img = ImageTk.PhotoImage(file=temp_file)
            label.config(image=img, text='', width=50, height=50)
            label.image = img
        except Exception:
            self.set_no_preview(label)

    def set_no_preview(self, label):
        label.config(image=self.no_preview_photo, text='', width=50, height=50)
        label.image = self.no_preview_photo
