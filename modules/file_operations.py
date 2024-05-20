import ctypes
import os
import subprocess  # 确保导入 subprocess 模块
from tkinter import messagebox
import tkinter as tk
import win32com.client

from modules.clean_graph import adjust_text_boxes
from modules.clean_master import clean_master_in_open_presentation
from modules.extract_slides import extract_text_from_presentation
from modules.standardize_filename_v7 import generate_standardized_name


class FileOperations:
    def open_ppt_selection_window(self, action, ui):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentations_info = self.list_open_presentations()

        top = tk.Toplevel(ui.root)
        top.title("选择演示文稿")

        if not presentations_info:
            messagebox.showinfo("提示", "没有找到打开的演示文稿。")
            top.destroy()
            return

        lbl = tk.Label(top, text="选择一个演示文稿:")
        lbl.pack(pady=10)

        lb = tk.Listbox(top, width=50, height=10)
        for item in presentations_info:
            lb.insert(tk.END, item)
        lb.pack(pady=5, padx=10)

        def on_select():
            selection = lb.curselection()
            if selection:
                selected_index = selection[0]
                presentation = powerpoint.Presentations[selected_index]
                action_func = {
                    'extract': extract_text_from_presentation,
                    'adjust': adjust_text_boxes,
                    'clean_master': clean_master_in_open_presentation
                }
                result = action_func[action](presentation)
                if action == 'extract':
                    ui.text_output.delete('1.0', tk.END)
                    ui.text_output.insert(tk.END, "\n".join(result))
                top.destroy()

        btn_select = tk.Button(top, text="选择", command=on_select)
        btn_select.pack(pady=5)

    def list_open_presentations(self):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        return [presentation.Name for presentation in powerpoint.Presentations]

    @staticmethod
    def is_hidden(filepath):
        """判断文件是否为隐藏文件"""
        attribute = ctypes.windll.kernel32.GetFileAttributesW(filepath)
        return attribute & 2  # FILE_ATTRIBUTE_HIDDEN = 0x2

    @staticmethod
    def open_file(filepath):
        """使用系统默认程序打开文件"""
        os.startfile(filepath)

    def generate_standardized_name(self, file_path, add_date, rename_folders):
        return generate_standardized_name(file_path, add_date, rename_folders)

    def rename_files(self, ui):
        batch_history = []
        for var, orig_name, std_name_var in ui.file_vars:
            if var.get():
                try:
                    orig_path = os.path.join(ui.folder_path.get(), orig_name)
                    new_path = os.path.join(ui.folder_path.get(), std_name_var.get())
                    os.rename(orig_path, new_path)
                    batch_history.append((orig_path, new_path))  # 记录重命名操作
                except Exception as e:
                    messagebox.showerror("错误", f"重命名 {orig_name} 到 {std_name_var.get()} 失败: {e}")
        if batch_history:
            ui.rename_history.append(batch_history)
            if len(ui.rename_history) > 10:
                ui.rename_history.pop(0)  # 保持最多10个批次的撤销操作
            ui.redo_history.clear()  # 清空重做历史
        messagebox.showinfo("完成", "文件重命名完成。")
        ui.refresh_file_list()

    def undo_rename(self, ui):
        if not ui.rename_history:
            messagebox.showinfo("提示", "没有可以撤销的重命名操作。")
            return
        last_batch = ui.rename_history.pop()
        for orig_path, new_path in reversed(last_batch):
            try:
                os.rename(new_path, orig_path)
            except Exception as e:
                messagebox.showerror("错误", f"撤销重命名 {new_path} 到 {orig_path} 失败: {e}")
        ui.redo_history.append(last_batch)
        ui.refresh_file_list()
        ui.update_buttons()

    def redo_rename(self, ui):
        if not ui.redo_history:
            messagebox.showinfo("提示", "没有可以重做的重命名操作。")
            return
        last_batch = ui.redo_history.pop()
        for orig_path, new_path in last_batch:
            try:
                os.rename(orig_path, new_path)
            except Exception as e:
                messagebox.showerror("错误", f"重做重命名 {orig_path} 到 {new_path} 失败: {e}")
        ui.rename_history.append(last_batch)
        ui.refresh_file_list()
        ui.update_buttons()
