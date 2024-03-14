from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QScrollArea, QGridLayout, QPushButton, QApplication, QMessageBox, QMenu
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt, QRunnable, QThreadPool, pyqtSignal, QObject
from findBackgroundIMG.helper_functions import *
import csv
from PyQt5.QtCore import QThread, pyqtSignal


class ImageLoaderTask(QRunnable):
    def __init__(self, imagePath, callback):
        super().__init__()
        self.imagePath = imagePath
        self.callback = callback

    def run(self):
        try:
            pixmap = QPixmap(self.imagePath).scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            success = True
        except Exception as e:
            print(f"Failed to load image {self.imagePath}: {e}")
            pixmap = QPixmap()
            success = False
        self.callback(self.imagePath, pixmap, success)



class ImageGalleryApp(QWidget):
    def __init__(self, csv_file_path):
        super().__init__()
        self.setWindowTitle("Image Gallery")
        self.threadPool = QThreadPool()

        # 获取屏幕尺寸，调整窗口大小
        screen_size = QApplication.instance().primaryScreen().size()
        self.resize(int(screen_size.width() * 0.8), int(screen_size.height() * 0.8))

        self.csv_file_path = csv_file_path
        self.images = self.load_images_from_csv()
        self.columns = 4  # Adjust based on your layout preference or calculate dynamically

        self.create_ui()

    def create_ui(self):
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        container = QWidget()
        self.grid_layout = QGridLayout(container)
        container.setLayout(self.grid_layout)
        self.scroll_area.setWidget(container)

        layout = QVBoxLayout(self)
        layout.addWidget(self.scroll_area)
        self.setLayout(layout)

        self.populate()  # 填充界面

    def load_images_from_csv(self):
        images = []
        with open(self.csv_file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                img_path = row['Image File']
                if img_path.lower().endswith(('.png', '.jpg', '.jpeg')) and not QImage(img_path).isNull():
                    images.append({'img_path': img_path, 'pptx_paths': [row['PPTX File']]})
        return images

    def populate(self):
        for image_data in self.images:
            callback = lambda img_path, pixmap, success, image_data=image_data: self.onImageLoaded(img_path, pixmap,
                                                                                                   success, image_data)
            loaderTask = ImageLoaderTask(image_data['img_path'], callback)
            self.threadPool.start(loaderTask)


    def add_image_widget(self, image_data):
        """为每个图片数据添加图片显示和操作按钮。

        输入:
        - image_data: 包含单个图片路径和相关PPTX文件路径的字典。

        功能: 创建显示图片的Label和对应的打开PPTX文件的按钮。不返回任何输出。
        """
        img_path = image_data['img_path']
        pptx_paths = image_data['pptx_paths']

        # 创建图片显示Label
        label = QLabel(self)
        pixmap = QPixmap(img_path)
        label.setPixmap(pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        # 添加到布局
        row = len(self.current_widgets)
        self.grid_layout.addWidget(label, row, 0)
        self.current_widgets.append(label)

        # 为每个PPTX文件添加打开按钮
        for pptx_path in pptx_paths:
            btn = QPushButton(os.path.basename(pptx_path), self)
            btn.clicked.connect(lambda checked, path=pptx_path: open_pptx(path))
            self.grid_layout.addWidget(btn, row, 1)
            self.current_widgets.append(btn)

    def onImageLoaded(self, img_path, pixmap, success, image_data):
        if not success:
            print(f"Error loading image: {img_path}")
            return
        # 通过图片路径找到对应的图片数据
        image_data = next((item for item in self.images if item['img_path'] == img_path), None)
        if not image_data:
            return  # 如果找不到对应的数据，直接返回

        # 找到对应图片数据在self.images中的索引，用于确定放置位置
        index = self.images.index(image_data)

        # 创建一个新的QWidget作为容器，里面包含图片Label和对应的PPTX打开按钮
        container_widget = QWidget()
        vertical_layout = QVBoxLayout(container_widget)

        # 创建并设置图片Label
        label = QLabel()
        label.setPixmap(pixmap)
        vertical_layout.addWidget(label)

        # 为每个PPTX文件添加打开按钮
        for pptx_path in image_data['pptx_paths']:
            pptx_btn = QPushButton(os.path.basename(pptx_path))
            pptx_btn.clicked.connect(lambda checked, path=pptx_path: open_pptx(path))
            vertical_layout.addWidget(pptx_btn)

        # 计算放置容器的行和列
        row = index // self.columns  # self.columns 需要根据实际布局确定，表示每行能容纳的图片数量
        column = index % self.columns

        # 将容器添加到网格布局中
        self.grid_layout.addWidget(container_widget, row, column)
        self.current_widgets.append(container_widget)  # 将容器添加到当前部件列表中，以便后续管理


class ClickableLabel(QLabel):
    """
    可点击的标签，用于展示图片，并提供右键菜单进行复制或删除操作。

    属性:
    - image_hash: 图片的哈希值，用于唯一标识图片。
    - img_path: 图片的文件路径。
    - gallery_app: 对ImageGalleryApp实例的引用，用于执行图片的删除操作。
    """

    def __init__(self, image_hash, img_path, gallery_app, parent=None):
        super().__init__(parent)
        self.image_hash = image_hash
        self.img_path = img_path
        self.gallery_app = gallery_app

        self.setPixmap(QPixmap(img_path).scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.showContextMenu)

    def showContextMenu(self, position):
        """
        在标签位置显示右键菜单。

        输入:
        - position: 右键点击的位置。

        功能: 显示一个包含“复制图片”和“删除图片”选项的菜单。
        """
        menu = QMenu(self)
        copy_action = menu.addAction("复制图片")
        delete_action = menu.addAction("删除图片")
        action = menu.exec_(self.mapToGlobal(position))

        if action == copy_action:
            self.copy_to_clipboard()
        elif action == delete_action:
            self.confirm_delete()

    def copy_to_clipboard(self):
        """
        将图片复制到剪贴板。

        功能: 加载图片并复制到系统剪贴板。不需要输入和输出。
        """
        clipboard = QApplication.clipboard()
        clipboard.setPixmap(QPixmap(self.img_path))

    def confirm_delete(self):
        """
        确认是否删除图片。

        功能: 弹出确认对话框，若确认则调用ImageGalleryApp实例删除图片。不需要输入和输出。
        """
        reply = QMessageBox.question(self, '删除确认', "你确定要删除这张图片吗？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.gallery_app.delete_image(self.image_hash)
