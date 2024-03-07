import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QScrollArea, QGridLayout
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import csv
import os
import subprocess
import threading
from PyQt5.QtWidgets import QSizePolicy
from PyQt5.QtCore import QSize


class ImageGalleryApp(QWidget):
    def __init__(self, csv_file_path):
        super().__init__()
        self.setWindowTitle("Image Gallery")

        # 获取屏幕尺寸，调整窗口大小
        screen_size = QApplication.instance().primaryScreen().size()
        self.resize(int(screen_size.width() * 0.8), int(screen_size.height() * 0.8))

        self.csv_file_path = csv_file_path
        self.images = self.load_images_from_csv()
        self.create_ui()
        self.setup_connections()
        self.pixmap_cache = {}  # 图片缓存
        self.current_widgets = []  # 当前显示的部件列表


    @staticmethod
    def clear_layout(layout):
        for i in reversed(range(layout.count())):
            widget = layout.itemAt(i).widget()
            if widget is not None:
                layout.removeWidget(widget)
                widget.deleteLater()

    def create_ui(self):
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)

        self.container = QWidget()
        self.grid_layout = QGridLayout(self.container)
        self.container.setLayout(self.grid_layout)

        self.scroll_area.setWidget(self.container)

        layout = QVBoxLayout(self)
        layout.addWidget(self.scroll_area)
        self.setLayout(layout)

    def update_widget_with_image_data(self, vertical_layout, image_data):
        img_path = image_data['img_path']
        pptx_paths = image_data['pptx_paths']

        # 首先更新或添加图片
        if vertical_layout.count() > 0:
            # 假设第一个widget是图片标签
            label = vertical_layout.itemAt(0).widget()
        else:
            label = QLabel()
            vertical_layout.addWidget(label)

        # 检查图片缓存，更新或添加图片
        if img_path not in self.pixmap_cache:
            pixmap = QPixmap(img_path).scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.pixmap_cache[img_path] = pixmap
        else:
            pixmap = self.pixmap_cache[img_path]
        label.setPixmap(pixmap)

        # 确保所有旧的按钮被清除，除了图片标签以外
        while vertical_layout.count() > 1:
            widget_to_remove = vertical_layout.itemAt(1).widget()
            vertical_layout.removeWidget(widget_to_remove)
            widget_to_remove.deleteLater()

        # 添加新的PPTX按钮
        for pptx_path in pptx_paths:
            pptx_btn_text = os.path.basename(pptx_path)
            pptx_btn = QPushButton(pptx_btn_text)
            pptx_btn.clicked.connect(lambda checked, path=pptx_path: self.open_pptx(path))
            vertical_layout.addWidget(pptx_btn)

    def showEvent(self, event):
        super().showEvent(event)
        self.populate()  # 确保窗口显示后填充元素

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

    def calculate_columns(self):
        container_width = self.scroll_area.width()  # 获取滚动区域的当前宽度
        item_width = 250  # 假设每个图片加上内边距等需要的宽度
        columns = max(1, container_width // item_width)  # 计算可以容纳的列数，至少为1
        return columns

    def populate(self):
        max_columns = self.calculate_columns()
        print(f"Max Columns: {max_columns}")

        required_widgets_count = len(self.images)
        current_widgets_count = len(self.current_widgets)

        # 移除多余的部件
        for i in range(required_widgets_count, current_widgets_count):
            widget = self.current_widgets.pop()
            self.grid_layout.removeWidget(widget)
            widget.deleteLater()

        row_number = 0
        col_number = 0

        for index, image_data in enumerate(self.images):
            if index < current_widgets_count:
                # 复用现有的widget和layout
                container_widget = self.current_widgets[index]
                vertical_layout = container_widget.layout()
            else:
                # 创建新的widget和layout
                container_widget = QWidget()
                vertical_layout = QVBoxLayout()
                container_widget.setLayout(vertical_layout)
                self.current_widgets.append(container_widget)

            # 更新图片和按钮
            self.update_widget_with_image_data(vertical_layout, image_data)

            # 更新widget位置
            if self.grid_layout.itemAtPosition(row_number, col_number) is not None:
                self.grid_layout.removeItem(self.grid_layout.itemAtPosition(row_number, col_number))
            self.grid_layout.addWidget(container_widget, row_number, col_number, 1, 1, Qt.AlignTop)

            col_number += 1
            if col_number >= max_columns:
                col_number = 0
                row_number += 1

        # 确保所有部件都是可见的
        for widget in self.current_widgets:
            widget.setVisible(True)

    def resizeEvent(self, event):
        self.populate()  # 重新排列图片和按钮
        super().resizeEvent(event)

    def setup_connections(self):
        pass

    def calculate_dynamic_padding(self):
        # 根据窗口大小计算动态内边距，这里只是一个基本示例
        window_width = self.width()
        if window_width < 800:  # 小窗口
            return 5, 5
        elif window_width < 1200:  # 中等窗口
            return 10, 10
        else:  # 大窗口
            return 15, 15

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
