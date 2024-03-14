import os
import subprocess
import threading
import sys


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