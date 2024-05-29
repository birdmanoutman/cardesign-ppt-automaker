# ui_components.py
import os
import sys
import tkinter as tk
from tkinter import Button, Checkbutton, IntVar, filedialog, messagebox
from tkinter.ttk import Combobox

import pythoncom
import win32com.client
from PIL import Image, ImageTk

from modules.file_operations import FileOperations
from utils.thumbnail_generator import ThumbnailGenerator


class PPTXToolUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX 工具包")

        self.file_operations = FileOperations(self)
        self.thumbnail_generator = ThumbnailGenerator()

        self.text_output = tk.Text(root, height=15, width=50)
        self.add_date = IntVar()
        self.rename_folders = IntVar()
        self.folder_path = tk.StringVar()
        self.rename_history = []
        self.redo_history = []
        self.file_vars = []
        self.navigation_history = []
        self.navigation_index = -1
        self.button_frame = None  # 初始化 button_frame 属性

        # Load and resize the folder icon
        self.folder_icon_path = self.resource_path('modules/icon_folder.png')
        self.folder_icon = Image.open(self.folder_icon_path)
        self.folder_icon = self.folder_icon.resize((50, 50), Image.LANCZOS)
        self.folder_icon = ImageTk.PhotoImage(self.folder_icon)

        self.setup_ui()

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def setup_ui(self):
        button_frame = self.create_frame(self.root)
        text_frame = self.create_frame(self.root)

        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.create_button(button_frame, "从PPTX提取文本",
                           lambda: self.file_operations.open_ppt_selection_window('extract', self))
        self.create_button(button_frame, "调整文本框",
                           lambda: self.file_operations.open_ppt_selection_window('adjust', self))
        self.create_button(button_frame, "清理PPTX母版",
                           lambda: self.file_operations.open_ppt_selection_window('clean_master', self))
        self.create_button(button_frame, "重命名目录下的文件", self.open_rename_files_window)

    def create_frame(self, parent, **kwargs):
        frame = tk.Frame(parent, padx=20, pady=20, **kwargs)
        frame.pack()
        return frame

    def create_button(self, parent, text, command, **kwargs):
        button = Button(parent, text=text, command=command, **kwargs)
        button.pack(pady=10, anchor='w')
        return button

    def open_rename_files_window(self):
        top = tk.Toplevel(self.root)
        top.title("重命名文件")

        nav_frame = self.create_frame(top)
        nav_frame.pack(fill=tk.X)
        self.create_nav_buttons(nav_frame)

        file_frame = self.create_frame(top)
        file_frame.pack(fill=tk.BOTH, expand=True)

        Checkbutton(file_frame, text="文件名添加默认日期", variable=self.add_date, command=self.refresh_file_list).pack(
            anchor='w')
        Checkbutton(file_frame, text="重命名文件夹", variable=self.rename_folders, command=self.refresh_file_list).pack(
            anchor='w')

        self.create_file_list_canvas(file_frame)
        self.button_frame = tk.Frame(top)  # 定义 button_frame
        self.button_frame.pack(pady=10)
        self.refresh_file_list()

    def create_nav_buttons(self, parent):
        self.btn_back = self.create_nav_button(parent, "←", self.navigate_back, tk.DISABLED)
        self.btn_forward = self.create_nav_button(parent, "→", self.navigate_forward, tk.DISABLED)
        self.btn_up = self.create_nav_button(parent, "↑", self.navigate_up)
        self.btn_refresh = self.create_nav_button(parent, "↻", self.refresh_file_list)

        self.path_entry = Combobox(parent, textvariable=self.folder_path, state='readonly')
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.create_nav_button(parent, "选择文件夹", self.open_folder_dialog)
        self.path_entry.bind("<Return>", self.on_path_entry_return)

    def on_path_entry_return(self, event):
        path = self.path_entry.get()
        if os.path.isdir(path):
            self.navigate_to_path(path)
        else:
            messagebox.showerror("错误", f"路径无效: {path}")

    def create_nav_button(self, parent, text, command, state=tk.NORMAL):
        button = tk.Button(parent, text=text, command=command, state=state)
        button.pack(side=tk.LEFT, padx=5)
        return button

    def create_file_list_canvas(self, parent):
        self.canvas = tk.Canvas(parent)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(parent, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.canvas_frame, anchor="nw")

        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

    def on_mouse_wheel(self, event):
        try:
            if self.canvas.winfo_exists():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception as e:
            print(f"Error during mouse wheel scroll: {e}")

    def open_folder_dialog(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.update_navigation_history(folder_selected)
            self.navigate_to_path(folder_selected)

    def navigate_to_path(self, path):
        self.update_navigation_history(path)
        self.folder_path.set(path)
        self.list_files_and_folders(path)

    def list_files_and_folders(self, folder_selected):
        self.clear_canvas_frame()

        header_frame = self.create_header_frame(self.canvas_frame)
        files_and_folders = self.get_sorted_files_and_folders(folder_selected)

        for index, (original_path, orig_name, std_name) in enumerate(files_and_folders):
            if os.path.isdir(original_path) and not self.rename_folders.get():
                continue
            self.create_file_row(self.canvas_frame, index, original_path, orig_name, std_name)

        self.canvas_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.update_buttons()

    def clear_canvas_frame(self):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

    def create_header_frame(self, parent):
        header_frame = tk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky='nsew', pady=2)
        tk.Label(header_frame, text="选择", width=10).grid(row=0, column=0, padx=5)
        tk.Label(header_frame, text="缩略图", width=10).grid(row=0, column=1, padx=5)
        tk.Label(header_frame, text="原始文件名", width=30).grid(row=0, column=2, padx=5)
        tk.Label(header_frame, text="目标文件名", width=50).grid(row=0, column=3, padx=5)
        return header_frame

    def get_sorted_files_and_folders(self, folder_selected):
        files_and_folders = [(os.path.join(folder_selected, name), name,
                              self.file_operations.generate_standardized_name(os.path.join(folder_selected, name),
                                                                              self.add_date.get(),
                                                                              self.rename_folders.get())) for name in
                             os.listdir(folder_selected) if not self.file_operations.is_hidden(
                os.path.join(folder_selected, name)) and not name.startswith('.')]
        return sorted(files_and_folders, key=self.sort_key)

    def sort_key(self, item):
        original_path, orig_name, std_name = item
        target_path = self.resolve_shortcut(original_path) if orig_name.lower().endswith('.lnk') else original_path
        is_different = orig_name != std_name
        is_folder = os.path.isdir(target_path)
        is_text = self.is_text_file(orig_name)
        is_image = self.is_image_file(orig_name)
        mod_time = os.path.getmtime(target_path)
        return (not is_different, not is_folder, not is_text, is_image, -mod_time)

    def is_text_file(self, filename):
        return filename.lower().endswith(('.txt', '.doc', '.docx', '.ppt', '.pptx', '.pdf'))

    def is_image_file(self, filename):
        return filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))

    def create_file_row(self, parent, index, original_path, orig_name, std_name):
        var = IntVar(value=0)
        orig_name_var = tk.StringVar(value=orig_name)
        std_name_var = tk.StringVar(value=std_name)
        self.file_vars.append((var, orig_name, std_name_var))

        row_frame = tk.Frame(parent)
        row_frame.grid(row=index + 1, column=0, sticky='nsew', pady=5)

        checkbutton = Checkbutton(row_frame, variable=var, width=10,
                                  command=lambda: self.on_file_var_check(var, orig_name, std_name_var))
        checkbutton.grid(row=0, column=0, padx=2)
        thumbnail_label = tk.Label(row_frame, width=50, height=50)
        self.set_thumbnail(original_path, thumbnail_label)
        thumbnail_label.grid(row=0, column=1, padx=2)
        orig_entry = self.create_entry(row_frame, 30, orig_name_var, 'readonly')
        orig_entry.grid(row=0, column=2, padx=5)
        orig_entry.bind("<Button-1>", lambda e, p=original_path: self.handle_file_or_folder_click(p))
        target_entry = self.create_entry(row_frame, 50, std_name_var)
        target_entry.grid(row=0, column=3, padx=5)
        target_entry.bind("<KeyRelease>", lambda e, v=var: v.set(1))

    def create_entry(self, parent, width, textvariable, state='normal'):
        entry = tk.Entry(parent, width=width, textvariable=textvariable, state=state)
        return entry

    def set_thumbnail(self, path, label):
        if os.path.isdir(path) or (path.lower().endswith('.lnk') and os.path.isdir(self.resolve_shortcut(path))):
            self.set_folder_thumbnail(label)
        else:
            self.thumbnail_generator.create_thumbnail(path, label)

    def set_folder_thumbnail(self, label):
        label.config(image=self.folder_icon)
        label.image = self.folder_icon

    def handle_file_or_folder_click(self, path):
        if os.path.isdir(path):
            self.update_navigation_history(path)
            self.folder_path.set(path)
            self.list_files_and_folders(path)
        elif path.lower().endswith('.lnk'):
            target = self.resolve_shortcut(path)
            if target and os.path.exists(target):
                self.handle_file_or_folder_click(target)
            else:
                print(f"Invalid shortcut target: {target}")
                self.file_operations.open_file(path)
        else:
            self.file_operations.open_file(path)

    def resolve_shortcut(self, path):
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path)
            target = shortcut.Targetpath
        except Exception as e:
            print(f"Failed to resolve shortcut for {path}: {e}")
            target = None
        finally:
            pythoncom.CoUninitialize()
        return target

    def navigate_back(self):
        if self.navigation_index > 0:
            self.navigation_index -= 1
            self.update_navigation_state()
            self.navigate_to_path(self.navigation_history[self.navigation_index])

    def navigate_forward(self):
        if self.navigation_index < len(self.navigation_history) - 1:
            self.navigation_index += 1
            self.update_navigation_state()
            self.navigate_to_path(self.navigation_history[self.navigation_index])

    def navigate_up(self):
        if self.folder_path.get():
            parent_folder = os.path.dirname(self.folder_path.get())
            self.update_navigation_history(parent_folder)
            self.navigate_to_path(parent_folder)

    def refresh_file_list(self):
        if self.folder_path.get():
            self.list_files_and_folders(self.folder_path.get())
            self.root.update_idletasks()

    def refresh_standardized_names(self):
        for var, orig_name, std_name_var in self.file_vars:
            if var.get() == 1 and self.add_date.get():
                std_name_var.set(self.file_operations.generate_standardized_name(
                    os.path.join(self.folder_path.get(), orig_name), self.add_date.get(), self.rename_folders.get()))

    def on_file_var_check(self, var, orig_name, std_name_var):
        if var.get() == 1 and self.add_date.get():
            std_name_var.set(self.file_operations.generate_standardized_name(
                os.path.join(self.folder_path.get(), orig_name), self.add_date.get(), self.rename_folders.get()))
        else:
            std_name_var.set(orig_name)

    def update_navigation_history(self, path):
        if self.navigation_index < len(self.navigation_history) - 1:
            self.navigation_history = self.navigation_history[:self.navigation_index + 1]
        self.navigation_history.append(path)
        self.navigation_index += 1
        self.update_navigation_state()

    def update_navigation_state(self):
        self.btn_back.config(state=tk.NORMAL if self.navigation_index > 0 else tk.DISABLED)
        self.btn_forward.config(
            state=tk.NORMAL if self.navigation_index < len(self.navigation_history) - 1 else tk.DISABLED)
        self.path_entry.set(self.navigation_history[self.navigation_index])

    def rename_files(self):
        selected_files = [orig_name for var, orig_name, _ in self.file_vars if var.get() == 1]
        self.file_operations.rename_files(self, selected_files, self.add_date.get(), self.rename_folders.get())
        self.refresh_file_list()

    def undo_rename(self):
        self.file_operations.undo_rename(self)

    def redo_rename(self):
        self.file_operations.redo_rename(self)

    def update_buttons(self):
        if not self.button_frame:
            return
        for widget in self.button_frame.winfo_children():
            widget.destroy()
        Button(self.button_frame, text="确定", command=self.rename_files).pack(side=tk.LEFT, pady=10, padx=5)
        Button(self.button_frame, text="撤销", command=self.undo_rename).pack(side=tk.LEFT, pady=10, padx=5)
        if self.redo_history:
            Button(self.button_frame, text="重做", command=self.redo_rename).pack(side=tk.LEFT, pady=10, padx=5)
