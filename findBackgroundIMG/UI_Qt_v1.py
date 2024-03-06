import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QScrollArea, QGridLayout
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt
import csv
import os
import subprocess
import threading


class ImageGalleryApp(QWidget):
    def __init__(self, csv_file_path):
        super().__init__()
        self.setWindowTitle("Image Gallery")

        # 获取屏幕尺寸，调整窗口大小
        screen_size = app.primaryScreen().size()
        self.resize(int(screen_size.width() * 0.8), int(screen_size.height() * 0.8))

        self.csv_file_path = csv_file_path
        self.images = self.load_images_from_csv()

        self.create_ui()

    def load_images_from_csv(self):
        images = {}
        with open(self.csv_file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                img_hash = row['Image Hash']
                if img_hash not in images:
                    images[img_hash] = {'img_path': row['Image File'], 'pptx_paths': [row['PPTX File']]}
                else:
                    if row['PPTX File'] not in images[img_hash]['pptx_paths']:
                        images[img_hash]['pptx_paths'].append(row['PPTX File'])
        return list(images.values())

    def create_ui(self):
        # 使用滚动区域
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)

        # 创建滚动区域内的容器
        self.container = QWidget()
        self.grid_layout = QGridLayout(self.container)

        self.populate()

        # 设置滚动区域的窗口
        self.scroll_area.setWidget(self.container)

        layout = QVBoxLayout(self)
        layout.addWidget(self.scroll_area)
        self.setLayout(layout)

    def populate(self):
        for row_number, image_data in enumerate(self.images):
            img_path = image_data['img_path']
            pptx_paths = image_data['pptx_paths']

            pixmap = QPixmap(img_path)
            pixmap = pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)

            label = QLabel()
            label.setPixmap(pixmap)

            self.grid_layout.addWidget(label, row_number, 0, 1, 1)

            # 为每个PPTX路径创建一个按钮
            for col_number, pptx_path in enumerate(pptx_paths):
                pptx_btn = QPushButton(os.path.basename(pptx_path))
                pptx_btn.clicked.connect(lambda checked, path=pptx_path: self.open_pptx(path))
                self.grid_layout.addWidget(pptx_btn, row_number, col_number + 1, 1, 1)

    def open_pptx(self, pptx_path):
        # 在后台线程中打开PPTX文件
        def run():
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(pptx_path)
                elif sys.platform == "darwin":  # macOS
                    subprocess.call(['open', pptx_path])
                else:  # Assume Linux or other
                    subprocess.call(['xdg-open', pptx_path])
            except Exception as e:
                print(f"Failed to open file: {pptx_path}, {e}")

        threading.Thread(target=run).start()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    gallery_app = ImageGalleryApp("source/image_ppt_mapping.csv")
    gallery_app.show()
    sys.exit(app.exec_())
