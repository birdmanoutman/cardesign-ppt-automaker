import os
import tkinter as tk
from tkinter import Button, Checkbutton, IntVar, filedialog
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

        # Load and resize the folder icon
        self.folder_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon_folder.png')
        self.folder_icon = Image.open(self.folder_icon_path)
        self.folder_icon = self.folder_icon.resize((50, 50), Image.LANCZOS)  # Adjust the size as needed
        self.folder_icon = ImageTk.PhotoImage(self.folder_icon)

        self.setup_ui()

    def setup_ui(self):
        button_frame = tk.Frame(self.root, padx=20, pady=20)
        text_frame = tk.Frame(self.root, padx=20, pady=20)

        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        Button(button_frame, text="从PPTX提取文本",
               command=lambda: self.file_operations.open_ppt_selection_window('extract', self)).pack(pady=10,
                                                                                                     anchor='w')
        Button(button_frame, text="调整文本框",
               command=lambda: self.file_operations.open_ppt_selection_window('adjust', self)).pack(pady=10, anchor='w')
        Button(button_frame, text="清理PPTX母版",
               command=lambda: self.file_operations.open_ppt_selection_window('clean_master', self)).pack(pady=10,
                                                                                                          anchor='w')
        Button(button_frame, text="重命名目录下的文件", command=self.open_rename_files_window).pack(pady=10, anchor='w')

    def open_rename_files_window(self):
        top = tk.Toplevel(self.root)
        top.title("重命名文件")

        nav_frame = tk.Frame(top, padx=10, pady=10)
        nav_frame.pack(fill=tk.X)

        self.btn_back = tk.Button(nav_frame, text="←", command=self.navigate_back, state=tk.DISABLED)
        self.btn_back.pack(side=tk.LEFT, padx=5)

        self.btn_forward = tk.Button(nav_frame, text="→", command=self.navigate_forward, state=tk.DISABLED)
        self.btn_forward.pack(side=tk.LEFT, padx=5)

        self.btn_up = tk.Button(nav_frame, text="↑", command=self.navigate_up)
        self.btn_up.pack(side=tk.LEFT, padx=5)

        self.btn_refresh = tk.Button(nav_frame, text="↻", command=self.refresh_file_list)
        self.btn_refresh.pack(side=tk.LEFT, padx=5)

        self.combo_path = Combobox(nav_frame, textvariable=self.folder_path, state='readonly')
        self.combo_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        btn_select_folder = tk.Button(nav_frame, text="选择文件夹", command=self.open_folder_dialog)
        btn_select_folder.pack(side=tk.LEFT, padx=5)

        file_frame = tk.Frame(top, padx=20, pady=20)
        file_frame.pack(fill=tk.BOTH, expand=True)

        Checkbutton(file_frame, text="文件名添加默认日期", variable=self.add_date, command=self.refresh_file_list).pack(
            anchor='w')
        Checkbutton(file_frame, text="重命名文件夹", variable=self.rename_folders, command=self.refresh_file_list).pack(
            anchor='w')

        self.canvas = tk.Canvas(file_frame)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.canvas_frame, anchor="nw")

        def _on_mouse_wheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        file_frame.rowconfigure(1, weight=1)
        file_frame.columnconfigure(1, weight=1)

        self.button_frame = tk.Frame(top)
        self.button_frame.pack(pady=10)

        self.refresh_file_list()

    def open_folder_dialog(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.update_navigation_history(folder_selected)
            self.folder_path.set(folder_selected)
            self.list_files_and_folders(folder_selected)

    def list_files_and_folders(self, folder_selected):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

        header_frame = tk.Frame(self.canvas_frame)
        header_frame.grid(row=0, column=0, sticky='nsew', pady=2)
        tk.Label(header_frame, text="选择", width=10).grid(row=0, column=0, padx=5)
        tk.Label(header_frame, text="缩略图", width=10).grid(row=0, column=1, padx=5)
        tk.Label(header_frame, text="原始文件名", width=30).grid(row=0, column=2, padx=5)
        tk.Label(header_frame, text="目标文件名", width=50).grid(row=0, column=3, padx=5)

        files_and_folders = os.listdir(folder_selected)

        self.file_vars = []

        def is_text_file(filename):
            return filename.lower().endswith(('.txt', '.doc', '.docx', '.ppt', '.pptx', '.pdf'))

        def is_image_file(filename):
            return filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))

        def get_file_mod_time(filepath):
            return os.path.getmtime(filepath)

        def resolve_shortcut(path):
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path)
            target = shortcut.Targetpath
            pythoncom.CoUninitialize()
            return target

        def sort_key(item):
            original_path, orig_name, std_name = item
            target_path = resolve_shortcut(original_path) if orig_name.lower().endswith('.lnk') else original_path
            is_different = orig_name != std_name
            is_folder = os.path.isdir(target_path)
            is_text = is_text_file(orig_name)
            is_image = is_image_file(orig_name)
            mod_time = get_file_mod_time(target_path)
            return (not is_different, not is_folder, not is_text, is_image, -mod_time)

        files_and_folders = [(os.path.join(folder_selected, name), name,
                              self.file_operations.generate_standardized_name(os.path.join(folder_selected, name),
                                                                              self.add_date.get(),
                                                                              self.rename_folders.get())) for name in
                             files_and_folders if not self.file_operations.is_hidden(
                os.path.join(folder_selected, name)) and not name.startswith('.')]
        files_and_folders.sort(key=sort_key)

        for index, (original_path, orig_name, std_name) in enumerate(files_and_folders):
            if os.path.isdir(original_path) and not self.rename_folders.get():
                continue

            # Check if original name is different from standardized name
            is_different = orig_name != std_name

            var = IntVar(value=1 if is_different else 0)
            orig_name_var = tk.StringVar(value=orig_name)
            std_name_var = tk.StringVar(value=std_name)
            self.file_vars.append((var, orig_name, std_name_var))

            row_frame = tk.Frame(self.canvas_frame)
            row_frame.grid(row=index + 1, column=0, sticky='nsew', pady=5)

            Checkbutton(row_frame, variable=var, width=10).grid(row=0, column=0, padx=2)

            thumbnail_label = tk.Label(row_frame, width=50, height=50)
            if os.path.isdir(original_path):
                self.set_folder_thumbnail(thumbnail_label)
            else:
                self.thumbnail_generator.create_thumbnail(original_path, thumbnail_label)
            thumbnail_label.grid(row=0, column=1, padx=2)

            orig_entry = tk.Entry(row_frame, width=30, textvariable=orig_name_var, state='readonly')
            orig_entry.grid(row=0, column=2, padx=5)
            orig_entry.bind("<Button-1>", lambda e, p=original_path: self.handle_file_or_folder_click(p))

            target_entry = tk.Entry(row_frame, width=50, textvariable=std_name_var)
            target_entry.grid(row=0, column=3, padx=5)
            target_entry.bind("<KeyRelease>", lambda e, v=var: self.auto_select_checkbox(e, v))

        self.canvas_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))  # Update the scroll region
        self.update_buttons()

    def auto_select_checkbox(self, event, var):
        var.set(1)

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
            if target:
                if os.path.isdir(target):
                    self.update_navigation_history(target)
                    self.folder_path.set(target)
                    self.list_files_and_folders(target)
                else:
                    self.file_operations.open_file(target)
            else:
                self.file_operations.open_file(path)
        else:
            self.file_operations.open_file(path)

    def resolve_shortcut(self, path):
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        target = shortcut.Targetpath
        pythoncom.CoUninitialize()
        return target

    def navigate_back(self):
        if self.navigation_index > 0:
            self.navigation_index -= 1
            self.update_navigation_state()
            self.list_files_and_folders(self.navigation_history[self.navigation_index])

    def navigate_forward(self):
        if self.navigation_index < len(self.navigation_history) - 1:
            self.navigation_index += 1
            self.update_navigation_state()
            self.list_files_and_folders(self.navigation_history[self.navigation_index])

    def navigate_up(self):
        if self.folder_path.get():
            parent_folder = os.path.dirname(self.folder_path.get())
            self.update_navigation_history(parent_folder)
            self.folder_path.set(parent_folder)
            self.list_files_and_folders(parent_folder)

    def refresh_file_list(self):
        if self.folder_path.get():
            self.list_files_and_folders(self.folder_path.get())
            self.root.update_idletasks()  # Force update the window layout

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
        self.combo_path.set(self.navigation_history[self.navigation_index])

    def rename_files(self):
        self.file_operations.rename_files(self)

    def undo_rename(self):
        self.file_operations.undo_rename(self)

    def redo_rename(self):
        self.file_operations.redo_rename(self)

    def update_buttons(self):
        for widget in self.button_frame.winfo_children():
            widget.destroy()
        Button(self.button_frame, text="确定", command=self.rename_files).pack(side=tk.LEFT, pady=10, padx=5)
        Button(self.button_frame, text="撤销", command=self.undo_rename).pack(side=tk.LEFT, pady=10, padx=5)
        if self.redo_history:
            Button(self.button_frame, text="重做", command=self.redo_rename).pack(side=tk.LEFT, pady=10, padx=5)
