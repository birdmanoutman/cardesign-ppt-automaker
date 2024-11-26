# resource_manager.py
import os
import sys
from PIL import Image, ImageTk, ImageDraw

class ResourceManager:
    def __init__(self):
        pass

    def resource_path(self, relative_path):
        """获取资源文件的绝对路径，兼容 PyInstaller 打包后的路径"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def get_folder_icon(self):
        try:
            folder_icon_path = self.resource_path('icon_folder.png')
            if not os.path.exists(folder_icon_path):
                # 如果文件不存在，创建一个简单的文件夹图标
                icon = Image.new('RGB', (50, 50), color='gray')
                draw = ImageDraw.Draw(icon)
                draw.rectangle([10, 20, 40, 40], fill='yellow', outline='black')
                draw.rectangle([15, 15, 35, 25], fill='yellow', outline='black')
                folder_icon = ImageTk.PhotoImage(icon)
                return folder_icon
            else:
                folder_icon = Image.open(folder_icon_path)
                folder_icon = folder_icon.resize((50, 50), Image.LANCZOS)
                folder_icon = ImageTk.PhotoImage(folder_icon)
                return folder_icon
        except Exception as e:
            print(f"Failed to load folder icon: {e}")
            return None
