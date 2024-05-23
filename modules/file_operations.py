import ctypes
import os
import re
import shutil
import tempfile
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

import win32com.client

from modules.clean_graph import adjust_text_boxes
from modules.clean_master import clean_master_in_open_presentation
from modules.extract_slides import extract_text_from_presentation


class FileOperations:
    def __init__(self, ui=None):
        self.ui = ui

    def standardize_filename(self, folder_path, add_default_date=True, rename_folders=False, specific_files=None):
        # Improved regular expressions for date formats
        existing_date_pattern = re.compile(r'^\d{6}_')
        full_date_pattern = re.compile(r'(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
        short_date_pattern = re.compile(r'(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
        cn_date_pattern = re.compile(r'(\d{1,2})月(\d{1,2})日')
        space_symbol_normalization_pattern = re.compile(r'[\s\-]+')
        multiple_underscores = re.compile(r'_{2,}')
        trailing_symbols_pattern = re.compile(r'[\s_\-]+(?=\.\w+$)|[\s_\-]+$')

        for filename in os.listdir(folder_path):
            original_path = os.path.join(folder_path, filename)

            if specific_files and filename not in specific_files:
                continue

            # Check if the current entry is a directory
            if os.path.isdir(original_path):
                if not rename_folders:
                    continue  # Skip renaming directories if not specified

            # Split the filename and its extension
            name, ext = os.path.splitext(filename)

            # Normalize spaces and symbols in filenames
            normalized_name = space_symbol_normalization_pattern.sub('_', name)
            normalized_name = multiple_underscores.sub('_', normalized_name)

            # Remove trailing spaces, underscores, and hyphens
            normalized_name = trailing_symbols_pattern.sub('', normalized_name)

            # Ensure no existing date
            if existing_date_pattern.match(normalized_name):
                continue  # Skip renaming if date is already properly prefixed

            last_modified_timestamp = os.path.getmtime(original_path)
            last_modified_date = datetime.fromtimestamp(last_modified_timestamp)

            # Date extraction and formatting
            date_match = full_date_pattern.search(normalized_name)
            if date_match:
                date_str = datetime.strptime(date_match.group(), '%Y%m%d').strftime('%y%m%d')
                normalized_name = full_date_pattern.sub('', normalized_name)
            else:
                date_match = short_date_pattern.search(normalized_name)
                if date_match:
                    date_str = last_modified_date.strftime('%y') + date_match.group()
                    normalized_name = short_date_pattern.sub('', normalized_name)
                else:
                    cn_date_match = cn_date_pattern.search(normalized_name)
                    if cn_date_match:
                        month_day_str = datetime.strptime(f'{cn_date_match.group(1)}{cn_date_match.group(2)}',
                                                          '%m%d').strftime('%m%d')
                        date_str = last_modified_date.strftime('%y') + month_day_str
                        normalized_name = cn_date_pattern.sub('', normalized_name)
                    elif add_default_date:
                        date_str = last_modified_date.strftime('%y%m%d')
                    else:
                        continue  # Skip renaming if no date found and default date not allowed

            # Build new filename and rename
            new_filename = f"{date_str}_{normalized_name.strip('_')}"
            new_filename = trailing_symbols_pattern.sub('', new_filename)  # Remove trailing symbols again
            new_path = os.path.join(folder_path, f"{new_filename}{ext}")
            if os.path.isdir(original_path):
                os.rename(original_path, new_path)  # Apply renaming for directories
            else:
                shutil.move(original_path, new_path)  # Apply renaming for files

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

    def generate_standardized_name(self, file_path, add_date, rename_folders):
        folder, filename = os.path.split(file_path)
        temp_folder = tempfile.mkdtemp()
        try:
            if os.path.isdir(file_path):
                temp_path = os.path.join(temp_folder, filename)
                os.mkdir(temp_path)
                self.standardize_filename(temp_folder, add_date, rename_folders, specific_files=[filename])
                std_name = os.listdir(temp_folder)[0]
            else:
                temp_path = os.path.join(temp_folder, filename)
                shutil.copy(file_path, temp_path)
                self.standardize_filename(temp_folder, add_date, rename_folders, specific_files=[filename])
                std_name = os.listdir(temp_folder)[0]
        finally:
            def remove_readonly(func, path, excinfo):
                os.chmod(path, 0o777)
                func(path)

            try:
                shutil.rmtree(temp_folder, onerror=remove_readonly)
            except Exception as e:
                print(f"无法删除临时文件夹: {e}")
        return std_name

    def rename_files(self, ui, selected_files, add_date, rename_folders):
        rename_operations = []
        for var, orig_name, std_name_var in ui.file_vars:
            if var.get() == 1:
                original_path = os.path.join(ui.folder_path.get(), orig_name)
                new_name = std_name_var.get()

                if add_date and orig_name in selected_files:
                    new_name = self.generate_standardized_name(original_path, add_date, rename_folders)

                new_path = os.path.join(ui.folder_path.get(), new_name)

                if not os.path.exists(original_path):
                    print(f"文件未找到：{original_path}")
                    continue

                if os.path.exists(new_path):
                    print(f"目标文件已存在：{new_path}")
                    continue

                try:
                    os.rename(original_path, new_path)
                    print(f"Renaming {original_path} to {new_path}")
                    rename_operations.append((new_path, original_path))
                except Exception as e:
                    print(f"重命名失败：{original_path} -> {new_path}，错误：{e}")

        if rename_operations:
            ui.rename_history.append(rename_operations)
            ui.redo_history.clear()

        ui.refresh_file_list()

    def undo_rename(self, ui):
        if ui.rename_history:
            last_operations = ui.rename_history.pop()
            for new_path, original_path in reversed(last_operations):
                try:
                    os.rename(new_path, original_path)
                    print(f"Undo renaming {new_path} to {original_path}")
                except Exception as e:
                    print(f"撤销重命名失败：{new_path} -> {original_path}，错误：{e}")

            ui.redo_history.append(last_operations)
            ui.refresh_file_list()

    def redo_rename(self, ui):
        if ui.redo_history:
            last_operations = ui.redo_history.pop()
            for original_path, new_path in last_operations:
                try:
                    os.rename(original_path, new_path)
                    print(f"Redo renaming {original_path} to {new_path}")
                except Exception as e:
                    print(f"重做重命名失败：{original_path} -> {new_path}，错误：{e}")

            ui.rename_history.append(last_operations)
            ui.refresh_file_list()
