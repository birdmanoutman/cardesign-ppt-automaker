import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QLabel, QProgressBar, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal
from findBackgroundIMG_sqliteV1 import SQLiteManager, ImageManager, ImageExtractor


class Worker(QThread):
    progress = pyqtSignal(int)

    def __init__(self, db_path, src_folder=None, dest_folder=None, csv_file_path=None, table_name=None):
        super().__init__()
        self.db_path = db_path
        self.src_folder = src_folder
        self.dest_folder = dest_folder
        self.csv_file_path = csv_file_path
        self.table_name = table_name

    def run(self):
        if self.csv_file_path:
            self.import_csv_data()
        if self.src_folder and self.dest_folder:
            self.process_pptx_files()

    def import_csv_data(self):
        db_manager = SQLiteManager(self.db_path)
        db_manager.connect()
        db_manager.create_table(self.table_name)
        db_manager.import_csv_to_database(self.csv_file_path, self.table_name)
        db_manager.close()
        self.progress.emit(100)  # Assuming CSV import is a single step for simplicity

    def process_pptx_files(self):
        db_manager = SQLiteManager(self.db_path)
        db_manager.connect()
        image_manager = ImageManager(db_manager, self.table_name)
        extractor = ImageExtractor(self.src_folder, self.dest_folder, image_manager)
        extractor.extract_images()
        db_manager.close()
        self.progress.emit(100)  # Simplification: Update progress as complete after processing


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'PPTX Image Processor'
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(100, 100, 400, 300)
        layout = QVBoxLayout()

        # Database Path
        layout.addWidget(QLabel('Database Path:'))
        self.dbPathEdit = QLineEdit()
        layout.addWidget(self.dbPathEdit)

        # CSV File Path
        layout.addWidget(QLabel('CSV File Path:'))
        self.csvPathEdit = QLineEdit()  # Separate QLineEdit for CSV file path
        self.csvPathBtn = QPushButton('Select CSV File')
        self.csvPathBtn.clicked.connect(self.openFileNameDialog)
        layout.addWidget(self.csvPathEdit)
        layout.addWidget(self.csvPathBtn)

        # Source Folder Path
        layout.addWidget(QLabel('Source Folder:'))
        self.srcFolderEdit = QLineEdit()
        layout.addWidget(self.srcFolderEdit)

        # Destination Folder Path
        layout.addWidget(QLabel('Destination Folder:'))
        self.destFolderEdit = QLineEdit()
        layout.addWidget(self.destFolderEdit)

        # Table Name
        layout.addWidget(QLabel('Table Name:'))
        self.tableNameEdit = QLineEdit()
        layout.addWidget(self.tableNameEdit)

        # Start Button
        self.startBtn = QPushButton('Start CSV Import')
        self.startBtn.clicked.connect(self.startCsvImport)
        layout.addWidget(self.startBtn)

        # Progress Bar
        self.progressBar = QProgressBar(self)
        layout.addWidget(self.progressBar)
        self.setLayout(layout)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "CSV Files (*.csv)", options=options)
        if fileName:
            self.dbPathEdit.setText(fileName)

    def startCsvImport(self):
        db_path = self.dbPathEdit.text()
        csv_file_path = self.csvPathEdit.text()  # Assuming you have a QLineEdit for CSV path
        table_name = "your_table_name"  # Set your table name here
        worker = Worker(db_path=db_path, csv_file_path=csv_file_path, table_name=table_name)
        worker.progress.connect(self.progressBar.setValue)
        worker.start()

    def startPptxProcessing(self):
        db_path = self.dbPathEdit.text()
        src_folder = self.srcFolderEdit.text()  # Assuming you have a QLineEdit for source folder
        dest_folder = self.destFolderEdit.text()  # Assuming you have a QLineEdit for destination folder
        table_name = "your_table_name"  # Set your table name here
        worker = Worker(db_path=db_path, src_folder=src_folder, dest_folder=dest_folder, table_name=table_name)
        worker.progress.connect(self.progressBar.setValue)
        worker.start()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())
