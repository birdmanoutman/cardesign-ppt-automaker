# findBackgroundIMG_V3.py
# 数据库存储版本
import io
import logging
import os

import mysql.connector
import pandas as pd
from PIL import Image
from mysql.connector import errorcode
from pptx import Presentation

from utils import calculate_hash, is_size_similar


class MySQLManager:
    def __init__(self, db_config):
        self.db_config = db_config
        self.conn = None
        self.cursor = None

    def connect(self):
        try:
            # 尝试连接到指定数据库
            self.conn = mysql.connector.connect(**self.db_config)
        except mysql.connector.Error as err:
            if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                print("Something is wrong with your user name or password")
            elif err.errno == errorcode.ER_BAD_DB_ERROR:
                # 如果数据库不存在，则尝试创建数据库
                self.create_database()
                # 重新连接到新创建的数据库
                self.conn = mysql.connector.connect(**self.db_config)
            else:
                print(err)
                return
        self.cursor = self.conn.cursor()
        print("Database connection established.")

    def create_database(self):
        try:
            tmp_conn = mysql.connector.connect(host=self.db_config['host'],
                                               user=self.db_config['user'],
                                               passwd=self.db_config['passwd'])
            tmp_cursor = tmp_conn.cursor()
            tmp_cursor.execute(f"CREATE DATABASE IF NOT EXISTS {self.db_config['database']}")
            print(f"Database {self.db_config['database']} created successfully.")
        except mysql.connector.Error as err:
            print(f"Failed creating database: {err}")
        finally:
            tmp_cursor.close()
            tmp_conn.close()

    def create_table(self, table_sql):
        try:
            self.cursor.execute(table_sql)
            self.conn.commit()
            print("Table is ready.")
        except mysql.connector.Error as err:
            print(f"Failed creating table: {err}")

    def import_csv_to_database(self, csv_file_path, table_name):
        # 读取CSV文件
        df = pd.read_csv(csv_file_path)

        # 导入数据到数据库
        for _, row in df.iterrows():
            img_hash = row['Image Hash']
            img_path = row['Image Path']  # 根据实际CSV列名进行调整
            pptx_path = row['PPTX Path']

            try:
                self.cursor.execute(f'INSERT INTO {table_name} (img_hash, img_path, pptx_path) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE img_hash=img_hash', (img_hash, img_path, pptx_path))
            except mysql.connector.Error as err:
                print(f"Failed to insert row: {err}")

        self.conn.commit()
        print("CSV import completed.")

    def close(self):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def check_duplicate(self, img_hash):
        query = f"SELECT EXISTS(SELECT 1 FROM your_table_name WHERE img_hash=%s LIMIT 1)"
        self.cursor.execute(query, (img_hash,))
        return self.cursor.fetchone()[0]

    def insert_image(self, img_hash, img_path, pptx_path):
        try:
            query = f"INSERT INTO your_table_name (img_hash, img_path, pptx_path) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE img_hash=img_hash"
            self.cursor.execute(query, (img_hash, img_path, pptx_path))
            self.conn.commit()
        except mysql.connector.Error as err:
            print(f"Failed to insert row: {err}")


class ImageManager:
    def __init__(self, db_manager):
        self.db_manager = db_manager

    def is_duplicate(self, img_hash):
        return self.db_manager.check_duplicate(img_hash)

    def save_image(self, img, img_hash, img_name, dest_folder):
        if self.is_duplicate(img_hash):
            logging.info(f"Duplicate image detected, not saved: {img_name}")
            return False
        else:
            if img.mode == 'RGBA' or img.mode == 'P':
                img = img.convert('RGB')

            img_path = os.path.join(dest_folder, img_name)
            img.save(img_path)
            logging.info(f"Saved image: {img_path}")
            self.db_manager.insert_image(img_hash, img_path, "source_pptx_placeholder")  # Add your logic here to include the pptx_path
            return True


class ImageExtractor:
    def __init__(self, src_folder, dest_folder, image_manager):
        self.src_folder = src_folder
        self.dest_folder = dest_folder
        self.image_manager = image_manager

    def extract_images(self):
        for root, _, files in os.walk(self.src_folder):
            for file in files:
                if file.endswith(".pptx") and not file.startswith("~$"):
                    pptx_path = os.path.join(root, file)
                    self.process_pptx(pptx_path)

    def process_pptx(self, pptx_path):
        presentation = Presentation(pptx_path)
        for slide_idx, slide in enumerate(presentation.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if shape.shape_type == 13:  # Picture type
                    img = Image.open(io.BytesIO(shape.image.blob))
                    if is_size_similar(img, presentation):
                        img_hash = calculate_hash(shape.image.blob)
                        if img_hash:
                            img_name = f"{os.path.splitext(os.path.basename(pptx_path))[0]}_{slide_idx + 1}_{shape_idx + 1}.jpg"
                            self.image_manager.save_image(img, img_hash, img_name, self.dest_folder)


# Define a new class for processing
class ImageProcessor:
    def __init__(self, db_config, src_folder, dest_folder, csv_file_path):
        self.db_config = db_config
        self.src_folder = src_folder
        self.dest_folder = dest_folder
        self.csv_file_path = csv_file_path

    def run(self):
        db_manager = MySQLManager(self.db_config)
        db_manager.connect()
        table_sql = '''
        CREATE TABLE IF NOT EXISTS cardesign-ppt-automaker_findBackgroundIMG (
            id INT AUTO_INCREMENT PRIMARY KEY,
            img_hash VARCHAR(255) UNIQUE,
            img_path VARCHAR(255),
            pptx_path VARCHAR(255)
        )'''
        db_manager.create_table(table_sql)

        if self.csv_file_path:
            db_manager.import_csv_to_database(self.csv_file_path, "your_table_name")

        image_manager = ImageManager(db_manager)

        if not os.path.exists(self.dest_folder):
            os.makedirs(self.dest_folder)

        extractor = ImageExtractor(self.src_folder, self.dest_folder, image_manager)
        extractor.extract_images()

        db_manager.close()


def main():
    src_folder = "your_source_folder"
    dest_folder = "your_destination_folder"
    db_config = {
        'host': 'localhost',
        'user': 'birdmanoutman',
        'passwd': 'Zez12345687！',
        'database': 'cardesign-ppt-automaker_findBackgroundIMG'
    }

    db_manager = MySQLManager(db_config)
    db_manager.connect()
    table_sql = '''
    CREATE TABLE IF NOT EXISTS your_table_name (
        id INT AUTO_INCREMENT PRIMARY KEY,
        img_hash VARCHAR(255) UNIQUE,
        img_path VARCHAR(255),
        pptx_path VARCHAR(255)
    )'''
    db_manager.create_table(table_sql)

    image_manager = ImageManager(db_manager)

    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)

    extractor = ImageExtractor(src_folder, dest_folder, image_manager)
    extractor.extract_images()

    db_manager.close()


if __name__ == "__main__":
    main()
