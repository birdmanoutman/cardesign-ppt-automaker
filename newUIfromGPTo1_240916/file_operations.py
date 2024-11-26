# file_operations.py
import os
import shutil
import subprocess
import pythoncom
import win32com.client
from tkinter import messagebox

class FileOperations:
    def __init__(self):
        self.rename_history = []
        self.redo_history = []

    def is_hidden(self, filepath):
        try:
            attrs = win32com.client.Dispatch("Scripting.FileSystemObject").GetFile(filepath).Attributes
            return attrs & 2  # FILE_ATTRIBUTE_HIDDEN == 2
        except Exception:
            return False

    def generate_standardized_name(self, path, add_date, rename_folders):
        base_name = os.path.basename(path)
        name, ext = os.path.splitext(base_name)
        if add_date:
            date_prefix = self.get_date_prefix(path)
            name = f"{date_prefix}_{name}"
        return name + ext

    def get_date_prefix(self, path):
        import datetime
        timestamp = os.path.getmtime(path)
        date = datetime.datetime.fromtimestamp(timestamp)
        return date.strftime("%Y%m%d")

    def resolve_shortcut(self, path):
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path)
            target = shortcut.Targetpath
            return target
        except Exception as e:
            print(f"Failed to resolve shortcut for {path}: {e}")
            return None
        finally:
            pythoncom.CoUninitialize()

    def open_file(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：{e}")

    def rename_files(self, folder_path, selected_files):
        self.rename_history.append([])
        for orig_name, new_name in selected_files:
            orig_path = os.path.join(folder_path, orig_name)
            new_path = os.path.join(folder_path, new_name)
            if os.path.exists(new_path):
                messagebox.showwarning("警告", f"目标文件已存在，跳过：{new_name}")
                continue
            try:
                shutil.move(orig_path, new_path)
                self.rename_history[-1].append((new_path, orig_path))
            except Exception as e:
                messagebox.showerror("错误", f"无法重命名文件：{orig_name}\n错误信息：{e}")

    def undo_rename(self):
        if self.rename_history:
            last_operation = self.rename_history.pop()
            self.redo_history.append([])
            for new_path, orig_path in reversed(last_operation):
                try:
                    shutil.move(new_path, orig_path)
                    self.redo_history[-1].append((orig_path, new_path))
                except Exception as e:
                    messagebox.showerror("错误", f"无法撤销重命名：{new_path}\n错误信息：{e}")
            return True
        return False

    def redo_rename(self):
        if self.redo_history:
            last_operation = self.redo_history.pop()
            self.rename_history.append([])
            for orig_path, new_path in last_operation:
                try:
                    shutil.move(orig_path, new_path)
                    self.rename_history[-1].append((new_path, orig_path))
                except Exception as e:
                    messagebox.showerror("错误", f"无法重做重命名：{orig_path}\n错误信息：{e}")
            return True
        return False

    def open_ppt_selection_window(self, action, ui):
        # 这里应该实现打开 PPT 文件并执行相应操作的逻辑
        messagebox.showinfo("提示", f"执行操作：{action}")
