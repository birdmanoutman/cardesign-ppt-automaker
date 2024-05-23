# UI_Qt_v1.py
import csv
import os
import sys
import tempfile

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QScrollArea, QGridLayout, QMenu
from PyQt5.QtWidgets import QMessageBox

from modules.findBackgroundIMG.utils import open_pptx


class ImageGalleryApp(QWidget):
    def __init__(self, csv_file_path):
        super().__init__()
        self.grid_layout = None
        self.container = None
        self.scroll_area = None
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
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
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
        img_hash = image_data.get('img_hash')  # 假设image_data字典中有img_hash这一项

        # 创建或更新ClickableLabel时，现在还需要传递img_path
        if vertical_layout.count() > 0:
            label = vertical_layout.itemAt(0).widget()
            if not isinstance(label, ClickableLabel):
                label.deleteLater()  # 删除原有的Label
                label = ClickableLabel(img_hash, img_path, self)  # 创建新的ClickableLabel并传递img_path
                vertical_layout.insertWidget(0, label)
        else:
            label = ClickableLabel(img_hash, img_path, self)
            vertical_layout.addWidget(label)

        label.image_hash = img_hash

        try:
            pixmap = self.pixmap_cache.get(img_path)

            if pixmap is None:
                pixmap = QPixmap(img_path)
                if pixmap.isNull():
                    # 如果图片无法加载，则打印错误信息并尝试继续
                    print(f"Unable to load the image at path: {img_path}")
                else:
                    # 如果图片加载成功，则将其加入到缓存
                    self.pixmap_cache[img_path] = pixmap

            if pixmap and not pixmap.isNull():
                label.setPixmap(pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                # 获取图片宽度用于后续按钮宽度调整
                image_width = pixmap.width()
            else:
                raise IOError(f"Cannot load image: {img_path}")

        except Exception as e:
            print(f"Error loading image {img_path}: {e}")
            # 在此处设置默认图片或错误标志
            label.setText("Image not available")

        # 确保所有旧的按钮被清除，除了图片标签以外
        while vertical_layout.count() > 1:
            widget_to_remove = vertical_layout.itemAt(1).widget()
            vertical_layout.removeWidget(widget_to_remove)
            widget_to_remove.deleteLater()

        # 添加新的PPTX按钮
        for pptx_path in pptx_paths:
            pptx_btn_text = os.path.basename(pptx_path)
            # 检查文本长度，并在必要时调整它
            if len(pptx_btn_text) > 20:
                pptx_btn_text = pptx_btn_text[:9] + "..." + pptx_btn_text[-7:]
            pptx_btn = QPushButton(pptx_btn_text)
            pptx_btn.clicked.connect(lambda checked, path=pptx_path: open_pptx(path))
            pptx_btn.setFixedWidth(image_width)  # 设置按钮宽度与图片宽度一致
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

            # 更新widget位置前，先移除旧的widget
            if self.grid_layout.itemAtPosition(row_number, col_number) is not None:
                # 移除并删除当前位置的widget
                item = self.grid_layout.itemAtPosition(row_number, col_number)
                widget_to_remove = item.widget()
                if widget_to_remove is not None:
                    self.grid_layout.removeWidget(widget_to_remove)
                    widget_to_remove.deleteLater()

            # 然后添加新的widget到GridLayout
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


    def delete_csv_entry(self, image_hash):
        # 创建一个临时文件
        temp_file, temp_file_path = tempfile.mkstemp()

        with open(self.csv_file_path, newline='', encoding='utf-8') as csvfile, os.fdopen(temp_file, 'w', newline='',
                                                                                          encoding='utf-8') as temp_csv:
            reader = csv.DictReader(csvfile)
            fieldnames = reader.fieldnames
            writer = csv.DictWriter(temp_csv, fieldnames=fieldnames)
            writer.writeheader()

            # 将不需要删除的行写入临时文件
            for row in reader:
                if row['Image Hash'] != image_hash:
                    writer.writerow(row)

        # 替换原始文件
        os.replace(temp_file_path, self.csv_file_path)

    def deleteImage(self, image_hash):
        # 查找要删除的图片路径
        img_path_to_delete = None
        for img in self.images:
            if img.get('Image Hash') == image_hash:  # 使用get方法来避免KeyError
                img_path_to_delete = img.get('img_path')
                break

        # 更新内存中的images数据结构，移除对应图片
        self.images = [img for img in self.images if img.get('Image Hash') != image_hash]

        # 重新写入CSV文件
        with open(self.csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Image Hash', 'Image File', 'PPTX File']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for img in self.images:
                for pptx_path in img.get('pptx_paths', []):
                    writer.writerow({'Image Hash': img.get('Image Hash'), 'Image File': img.get('img_path'),
                                     'PPTX File': pptx_path})

        # 从文件系统中删除图片文件
        if img_path_to_delete and os.path.exists(img_path_to_delete):
            os.remove(img_path_to_delete)

        # 重新加载界面
        self.populate()


class ClickableLabel(QLabel):
    def __init__(self, image_hash, img_path, parent=None):
        super().__init__(parent)
        self.image_hash = image_hash  # 保存图片的哈希值或其他唯一标识
        self.img_path = img_path  # 保存原图的路径
        self.gallery_app = gallery_app  # 存储对ImageGalleryApp实例的引用

    def copyToClipboard(self):
        # 加载原图
        pixmap = QPixmap(self.img_path)
        # 复制到剪贴板
        QApplication.clipboard().setPixmap(pixmap)

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.showContextMenu(event.pos())
        else:
            super().mousePressEvent(event)

    def showContextMenu(self, position):
        menu = QMenu()
        copy_action = menu.addAction("复制图片")
        delete_action = menu.addAction("删除图片")
        action = menu.exec_(self.mapToGlobal(position))
        if action == copy_action:
            self.copyToClipboard()
        elif action == delete_action:
            self.deleteImage()

    def deleteImage(self):
        reply = QMessageBox.question(self, '删除确认', "你确定要删除这张图片吗？", QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.gallery_app.delete_csv_entry(self.image_hash)  # 通过gallery_app引用调用方法


if __name__ == "__main__":
    app = QApplication(sys.argv)
    gallery_app = ImageGalleryApp("/Users/birdmanoutman/上汽/backgroundIMGsource/image_ppt_mapping.csv")
    gallery_app.show()
    sys.exit(app.exec_())
