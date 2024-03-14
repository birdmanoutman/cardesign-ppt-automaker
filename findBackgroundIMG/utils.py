import os
import subprocess
import sys
import threading
import io
import imagehash
from PIL import Image
import logging


def open_pptx(pptx_path):
    """在后台线程中打开PPTX文件。

    输入:
    - pptx_path: PPTX文件的路径

    功能: 根据操作系统打开对应的PPTX文件。不返回任何输出。
    """

    def run():
        try:
            if os.name == 'nt':
                os.startfile(pptx_path)
            elif sys.platform == "darwin":
                subprocess.call(['open', pptx_path])
            else:
                subprocess.call(['xdg-open', pptx_path])
        except Exception as e:
            print(f"Failed to open file: {pptx_path}, {e}")

    threading.Thread(target=run).start()


def calculate_hash(image_blob):
    try:
        image_bytes = io.BytesIO(image_blob)
        img = Image.open(image_bytes)
        return str(imagehash.average_hash(img))
    except Exception as e:
        logging.error(f"Error calculating image hash: {e}")
        return None


def is_size_similar(img_width, img_height, slide_width, slide_height):
    """判断图片尺寸是否符合要求。"""
    tolerance = 500
    return (abs(img_width - slide_width) < tolerance and abs(img_height - slide_height) < tolerance) or (img_width >= slide_width and img_height >= slide_height)

