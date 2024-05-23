import sys
from PyQt5.QtWidgets import QApplication
from UIforWindows import ImageGalleryApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    csv_file_path =r'D:\PycharmProjects\cardesign-ppt-automaker\findBackgroundIMG\source\image_ppt_mapping.csv'
    gallery_app = ImageGalleryApp(csv_file_path)
    gallery_app.show()
    sys.exit(app.exec_())
