import os
import sqlite3
import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QScrollArea, QGridLayout, QMenu, \
    QMessageBox

from modules.findBackgroundIMG.utils import open_pptx


class ImageGalleryApp(QWidget):
    def __init__(self, db_path):
        super().__init__()
        self.setWindowTitle("Image Gallery")
        self.db_path = db_path
        self.images = self.load_images_from_db()

        screen_size = QApplication.instance().primaryScreen().size()
        self.resize(int(screen_size.width() * 0.8), int(screen_size.height() * 0.8))
        self.create_ui()
        self.setup_connections()
        self.pixmap_cache = {}
        self.current_widgets = []

    def create_ui(self):
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.container = QWidget()
        self.grid_layout = QGridLayout(self.container)
        self.container.setLayout(self.grid_layout)
        layout = QVBoxLayout(self)
        layout.addWidget(self.scroll_area)
        self.scroll_area.setWidget(self.container)
        self.setLayout(layout)

    def clear_layout(self, layout):
        for i in reversed(range(layout.count())):
            widget = layout.itemAt(i).widget()
            if widget is not None:
                layout.removeWidget(widget)
                widget.deleteLater()

    def update_widget_with_image_data(self, vertical_layout, image_data):
        img_path = image_data['img_path']
        pptx_paths = image_data['pptx_paths']
        img_hash = image_data.get('img_hash')

        if vertical_layout.count() > 0:
            label = vertical_layout.itemAt(0).widget()
            if not isinstance(label, ClickableLabel):
                label.deleteLater()
                label = ClickableLabel(img_hash, img_path, self)
                vertical_layout.insertWidget(0, label)
        else:
            label = ClickableLabel(img_hash, img_path, self)
            vertical_layout.addWidget(label)

        label.image_hash = img_hash

        if img_path not in self.pixmap_cache:
            pixmap = QPixmap(img_path).scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.pixmap_cache[img_path] = pixmap
        else:
            pixmap = self.pixmap_cache[img_path]
        label.setPixmap(pixmap)

        image_width = pixmap.width()

        while vertical_layout.count() > 1:
            widget_to_remove = vertical_layout.itemAt(1).widget()
            vertical_layout.removeWidget(widget_to_remove)
            widget_to_remove.deleteLater()

        for pptx_path in pptx_paths:
            pptx_btn_text = os.path.basename(pptx_path)
            if len(pptx_btn_text) > 20:
                pptx_btn_text = pptx_btn_text[:9] + "..." + pptx_btn_text[-7:]
            pptx_btn = QPushButton(pptx_btn_text)
            pptx_btn.clicked.connect(lambda checked, path=pptx_path: open_pptx(path))
            pptx_btn.setFixedWidth(image_width)
            vertical_layout.addWidget(pptx_btn)

    def load_images_from_db(self):
        images = {}
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT img_hash, img_path, pptx_path, is_duplicate FROM image_ppt_mapping")
        for row in cursor.fetchall():
            img_hash, img_path, pptx_path, is_duplicate = row
            if img_path.lower().endswith(('.png', '.jpg', '.jpeg')):
                if QImage(img_path).isNull() or is_duplicate:
                    print(f"Unable to load image: {img_path}")
                    continue
                if img_hash not in images:
                    images[img_hash] = {'img_path': img_path, 'pptx_paths': [pptx_path]}
                else:
                    if pptx_path not in images[img_hash]['pptx_paths']:
                        images[img_hash]['pptx_paths'].append(pptx_path)
        conn.close()
        return list(images.values())

    def calculate_columns(self):
        container_width = self.scroll_area.width()
        item_width = 250
        columns = max(1, container_width // item_width)
        return columns

    def populate(self):
        self.clear_unused_widgets()
        max_columns = self.calculate_columns()

        row_number = 0
        col_number = 0
        for index, image_data in enumerate(self.images):
            if index < len(self.current_widgets):
                container_widget = self.current_widgets[index]
            else:
                container_widget = QWidget()
                vertical_layout = QVBoxLayout()
                container_widget.setLayout(vertical_layout)
                self.current_widgets.append(container_widget)
                self.grid_layout.addWidget(container_widget, row_number, col_number)

            vertical_layout = container_widget.layout()
            self.update_widget_with_image_data(vertical_layout, image_data)

            col_number += 1
            if col_number >= max_columns:
                col_number = 0
                row_number += 1

        self.remove_excess_widgets(row_number, col_number, max_columns)

    def clear_unused_widgets(self):
        for i in reversed(range(len(self.current_widgets))):
            widget = self.current_widgets[i]
            self.grid_layout.removeWidget(widget)
            widget.deleteLater()
        self.current_widgets.clear()

    def remove_excess_widgets(self, row_number, col_number, max_columns):
        for index in range(row_number * max_columns + col_number, self.grid_layout.count()):
            layout_item = self.grid_layout.itemAt(index)
            if layout_item is not None:
                widget = layout_item.widget()
                if widget:
                    self.grid_layout.removeWidget(widget)
                    widget.deleteLater()

    def resizeEvent(self, event):
        self.populate()
        super().resizeEvent(event)

    def setup_connections(self):
        pass

    def delete_db_entry(self, image_hash):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM image_ppt_mapping WHERE img_hash=?", (image_hash,))
        conn.commit()
        conn.close()

    def delete_image(self, image_hash):
        img_path_to_delete = None
        for img in self.images:
            if img.get('img_hash') == image_hash:
                img_path_to_delete = img.get('img_path')
                break

        self.images = [img for img in self.images if img.get('img_hash') != image_hash]

        if img_path_to_delete and os.path.exists(img_path_to_delete):
            os.remove(img_path_to_delete)

        self.populate()


class ClickableLabel(QLabel):
    def __init__(self, image_hash, img_path, parent=None):
        super().__init__(parent)
        self.image_hash = image_hash
        self.img_path = img_path
        self.parent = parent

    def copy_to_clipboard(self):
        pixmap = QPixmap(self.img_path)
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
            self.copy_to_clipboard()
        elif action == delete_action:
            self.deleteImage()

    def deleteImage(self):
        reply = QMessageBox.question(self, '删除确认', "你确定要删除这张图片吗？", QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.parent.delete_db_entry(self.image_hash)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    gallery_app = ImageGalleryApp("image_gallery.db")
    gallery_app.show()
    sys.exit(app.exec_())
