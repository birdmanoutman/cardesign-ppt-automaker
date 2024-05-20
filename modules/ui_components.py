import os
import subprocess
import tkinter as tk
from tkinter import Button, Checkbutton, IntVar, filedialog, messagebox
from modules.file_operations import FileOperations
from utils.thumbnail_generator import ThumbnailGenerator

class PPTXToolUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX 工具包")

        self.file_operations = FileOperations()
        self.thumbnail_generator = ThumbnailGenerator()

        self.text_output = tk.Text(root, height=15, width=50)
        self.add_date = IntVar()
        self.rename_folders = IntVar()
        self.folder_path = tk.StringVar()
        self.rename_history = []
        self.redo_history = []
        self.file_vars = []

        self.setup_ui()

    def setup_ui(self):
        button_frame = tk.Frame(self.root, padx=20, pady=20)
        text_frame = tk.Frame(self.root, padx=20, pady=20)

        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        Button(button_frame, text="从PPTX提取文本", command=lambda: self.file_operations.open_ppt_selection_window('extract', self)).pack(pady=10, anchor='w')
        Button(button_frame, text="调整文本框", command=lambda: self.file_operations.open_ppt_selection_window('adjust', self)).pack(pady=10, anchor='w')
        Button(button_frame, text="清理PPTX母版", command=lambda: self.file_operations.open_ppt_selection_window('clean_master', self)).pack(pady=10, anchor='w')
        Button(button_frame, text="重命名目录下的文件", command=self.open_rename_files_window).pack(pady=10, anchor='w')

    def open_rename_files_window(self):
        top = tk.Toplevel(self.root)
        top.title("重命名文件")

        file_frame = tk.Frame(top, padx=20, pady=20)
        file_frame.pack(fill=tk.BOTH, expand=True)

        self.label_path = tk.Label(file_frame, text="当前目录: 未选择")
        btn_folder = tk.Button(file_frame, text="选择文件夹", command=self.open_folder_dialog)
        btn_folder.grid(row=0, column=0, pady=10, sticky='w')
        self.label_path.grid(row=0, column=1, pady=10, sticky='w')

        Checkbutton(file_frame, text="文件名添加默认日期", variable=self.add_date, command=self.refresh_file_list).grid(row=0, column=2, padx=10, sticky='e')
        Checkbutton(file_frame, text="重命名文件夹", variable=self.rename_folders, command=self.refresh_file_list).grid(row=0, column=3, padx=10, sticky='e')

        canvas = tk.Canvas(file_frame)
        canvas.grid(row=1, column=0, columnspan=4, sticky='nsew')

        scrollbar = tk.Scrollbar(file_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.grid(row=1, column=4, sticky='ns')

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=canvas_frame, anchor="nw")

        def _on_mouse_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        file_frame.rowconfigure(1, weight=1)
        file_frame.columnconfigure(1, weight=1)

        self.canvas_frame = canvas_frame
        self.button_frame = tk.Frame(top)
        self.button_frame.pack(pady=10)

        self.refresh_file_list()

    def open_folder_dialog(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.label_path.config(text=f"当前目录: {self.folder_path.get()}")
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

        for index, name in enumerate(files_and_folders):
            original_path = os.path.join(folder_selected, name)
            if self.file_operations.is_hidden(original_path) or name.startswith('.'):
                continue
            if os.path.isdir(original_path) and not self.rename_folders.get():
                continue
            orig_name = name
            std_name = self.file_operations.generate_standardized_name(original_path, self.add_date.get(),
                                                                       self.rename_folders.get())

            var = IntVar(value=1)
            orig_name_var = tk.StringVar(value=orig_name)
            std_name_var = tk.StringVar(value=std_name)
            self.file_vars.append((var, orig_name, std_name_var))

            row_frame = tk.Frame(self.canvas_frame)
            row_frame.grid(row=index + 1, column=0, sticky='nsew', pady=5)

            Checkbutton(row_frame, variable=var, width=10).grid(row=0, column=0, padx=2)

            thumbnail_label = tk.Label(row_frame, width=50, height=50)
            if os.path.isdir(original_path):
                self.thumbnail_generator.set_no_preview(thumbnail_label)  # 对文件夹使用“无预览”图像
            else:
                self.thumbnail_generator.create_thumbnail(original_path, thumbnail_label)
            thumbnail_label.grid(row=0, column=1, padx=2)

            orig_entry = tk.Entry(row_frame, width=30, textvariable=orig_name_var, state='readonly')
            orig_entry.grid(row=0, column=2, padx=5)
            orig_entry.bind("<Button-1>", lambda e, p=original_path: self.handle_file_or_folder_click(p))
            tk.Entry(row_frame, width=50, textvariable=std_name_var).grid(row=0, column=3, padx=5)

        self.update_buttons()

    def handle_file_or_folder_click(self, path):
        if os.path.isdir(path):
            self.folder_path.set(path)
            self.label_path.config(text=f"当前目录: {self.folder_path.get()}")
            self.list_files_and_folders(path)
        else:
            self.file_operations.open_file(path)

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

    def refresh_file_list(self):
        if self.folder_path.get():
            self.list_files_and_folders(self.folder_path.get())
