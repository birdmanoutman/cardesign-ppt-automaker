import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QScrollArea, QGridLayout
from PyQt5.QtGui import QPixmap
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
        screen_size = QApplication.instance().primaryScreen().size()
        self.resize(int(screen_size.width() * 0.8), int(screen_size.height() * 0.8))

        self.csv_file_path = csv_file_path
        self.images = self.load_images_from_csv()

        self.create_ui()
        self.setup_connections()  # 设置信号和槽的连接（如果有）

    def create_ui(self):
        self.scroll_area = QScrollArea(self)  # 创建滚动区域
        self.scroll_area.setWidgetResizable(True)

        self.container = QWidget()  # 创建用于放置内容的容器
        self.grid_layout = QGridLayout()  # 创建网格布局
        self.container.setLayout(self.grid_layout)  # 将网格布局设置给容器

        self.populate()  # 填充内容

        self.scroll_area.setWidget(self.container)  # 将容器设置为滚动区域的子Widget

        layout = QVBoxLayout(self)  # 创建一个垂直布局
        layout.addWidget(self.scroll_area)  # 将滚动区域添加到垂直布局中
        self.setLayout(layout)  # 将垂直布局设置为主窗口的布局

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

    def clear_layout(self, layout):
        for i in reversed(range(layout.count())):
            widget = layout.itemAt(i).widget()
            if widget is not None:
                layout.removeWidget(widget)
                widget.deleteLater()


    def calculate_columns(self):
        container_width = self.scroll_area.width()  # 获取滚动区域的当前宽度
        item_width = 250  # 假设每个图片加上内边距等需要的宽度
        columns = max(1, container_width // item_width)  # 计算可以容纳的列数，至少为1
        return columns

    def populate(self):
        self.clear_layout(self.grid_layout)  # 清除当前网格布局中的所有项

        rows = 0
        max_columns = self.calculate_columns()  # 动态计算最大列数
        dynamic_padx, dynamic_pady = self.calculate_dynamic_padding()  # 动态计算内边距

        for row_number, image_data in enumerate(self.images):
            img_path = image_data['img_path']
            pptx_paths = image_data['pptx_paths']

            pixmap = QPixmap(img_path).scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            label = QLabel()
            label.setPixmap(pixmap)

            vertical_layout = QVBoxLayout()
            vertical_layout.addWidget(label)

            for pptx_path in pptx_paths:
                pptx_btn_text = os.path.basename(pptx_path)
                if len(pptx_btn_text) > 16:
                    pptx_btn_text = pptx_btn_text[:8] + "..." + pptx_btn_text[-5:]
                pptx_btn = QPushButton(pptx_btn_text)
                pptx_btn.clicked.connect(lambda checked, path=pptx_path: self.open_pptx(path))
                vertical_layout.addWidget(pptx_btn)

            container_widget = QWidget()
            container_widget.setLayout(vertical_layout)
            # 注意，这里没有显式处理列的增加，假设每个 container_widget 占据一列
            self.grid_layout.addWidget(container_widget, rows, row_number % max_columns, 1, 1, Qt.AlignTop)

            if (row_number + 1) % max_columns == 0:
                rows += 1

    def resizeEvent(self, event):
        self.populate()  # 重新排列图片和按钮
        super().resizeEvent(event)

    def arrange_images(self):
        rows, columns = 0, 0
        max_columns = self.calculate_columns()  # 动态计算列数

        for btn, img_path, pptx_path in self.image_buttons:
            if columns >= max_columns:
                columns = 0
                rows += 1

            # 假设这里的 btn 是一个自定义的 QWidget，其中包含了图片和按钮
            # 动态调整内边距，可以根据需要调整 10 的值
            dynamic_padx, dynamic_pady = self.calculate_dynamic_padding()

            btn.grid(row=rows, column=columns, padx=dynamic_padx, pady=dynamic_pady)
            columns += 1

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
