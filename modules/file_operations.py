import ctypes
import os
from tkinter import messagebox
import tkinter as tk
import win32com.client

from modules.clean_graph import adjust_text_boxes
from modules.clean_master import clean_master_in_open_presentation
from modules.extract_slides import extract_text_from_presentation
from modules.standardize_filename_v7 import generate_standardized_name

class FileOperations:
    def __init__(self, ui=None):
        self.ui = ui

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

    def open_file(self, path):
        os.startfile(path)

    def is_hidden(self, filepath):
        return bool(os.stat(filepath).st_file_attributes & (ctypes.windll.kernel32.GetFileAttributesW(filepath) & 2))

    def generate_standardized_name(self, original_path, add_date, rename_folders):
        return generate_standardized_name(original_path, add_date, rename_folders)

    def rename_files(self, ui):
        for var, orig_name, std_name_var in ui.file_vars:
            if var.get():
                original_path = os.path.join(ui.folder_path.get(), orig_name)
                new_path = os.path.join(ui.folder_path.get(), std_name_var.get())
                if not os.path.exists(new_path):
                    os.rename(original_path, new_path)
        ui.refresh_file_list()

    def undo_rename(self, ui):
        if ui.rename_history:
            for original_path, new_path in ui.rename_history.pop():
                os.rename(new_path, original_path)
            ui.redo_history.append(ui.rename_history.pop())
            ui.refresh_file_list()

    def redo_rename(self, ui):
        if ui.redo_history:
            for original_path, new_path in ui.redo_history.pop():
                os.rename(original_path, new_path)
            ui.rename_history.append(ui.redo_history.pop())
            ui.refresh_file_list()
