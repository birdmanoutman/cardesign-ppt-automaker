# ui_components.py
import os
import sys
import threading
import tkinter as tk
from tkinter import Button, Checkbutton, IntVar, filedialog, messagebox
from tkinter.ttk import Combobox

from file_operations import FileOperations
from navigation import Navigation
from thumbnail_generator import ThumbnailGenerator
from resource_manager import ResourceManager

class PPTXToolUI:
    def __init__(self, root):
        """
        初始化用户界面。

        参数：
        root -- Tkinter 主窗口
        """
        self.root = root
        self.root.title("PPTX 工具包")

        self.file_operations = FileOperations()
        self.navigation = Navigation()
        self.thumbnail_generator = ThumbnailGenerator()
        self.resource_manager = ResourceManager()

        self.text_output = tk.Text(root, height=15, width=50)
        self.add_date = IntVar()
        self.rename_folders = IntVar()
        self.folder_path = tk.StringVar()
        self.file_vars = []
        self.button_frame = None

        # 加载文件夹图标
        self.folder_icon = self.resource_manager.get_folder_icon()

        self.setup_ui()

    def setup_ui(self):
        """
        设置主界面的布局和组件。
        """
        button_frame = self.create_frame(self.root)
        text_frame = self.create_frame(self.root)

        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.create_button(button_frame, "从 PPTX 提取文本",
                           lambda: self.file_operations.open_ppt_selection_window('extract', self))
        self.create_button(button_frame, "调整文本框",
                           lambda: self.file_operations.open_ppt_selection_window('adjust', self))
        self.create_button(button_frame, "清理 PPTX 母版",
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
        """
        打开重命名文件的窗口。
        """
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
        self.button_frame = tk.Frame(top)
        self.button_frame.pack(pady=10)
        self.refresh_file_list()

    def create_nav_buttons(self, parent):
        """
        创建导航按钮，包括前进、后退、上一级和刷新。

        参数：
        parent -- 父容器
        """
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
        """
        创建文件列表的滚动区域。

        参数：
        parent -- 父容器
        """
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
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def open_folder_dialog(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.navigation.navigate_to(folder_selected)
            self.update_navigation_state()
            self.navigate_to_path(folder_selected)

    def navigate_to_path(self, path):
        self.folder_path.set(path)
        self.list_files_and_folders(path)

    def list_files_and_folders(self, folder_selected):
        """
        列出指定文件夹中的文件和文件夹。

        参数：
        folder_selected -- 要列出的文件夹路径
        """
        def load_files():
            self.clear_canvas_frame()
            self.file_vars.clear()

            header_frame = self.create_header_frame(self.canvas_frame)
            files_and_folders = self.get_sorted_files_and_folders(folder_selected)

            for index, (original_path, orig_name, std_name) in enumerate(files_and_folders):
                if os.path.isdir(original_path) and not self.rename_folders.get():
                    continue
                self.create_file_row(self.canvas_frame, index, original_path, orig_name, std_name)

            self.canvas_frame.update_idletasks()
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            self.update_buttons()

        threading.Thread(target=load_files).start()

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
        files_and_folders = []
        try:
            for name in os.listdir(folder_selected):
                original_path = os.path.join(folder_selected, name)
                if self.file_operations.is_hidden(original_path) or name.startswith('.'):
                    continue
                std_name = self.file_operations.generate_standardized_name(
                    original_path, self.add_date.get(), self.rename_folders.get())
                files_and_folders.append((original_path, name, std_name))
            files_and_folders.sort(key=self.sort_key)
        except Exception as e:
            messagebox.showerror("错误", f"无法列出文件夹内容：{e}")
        return files_and_folders

    def sort_key(self, item):
        original_path, orig_name, std_name = item
        target_path = self.file_operations.resolve_shortcut(original_path) if orig_name.lower().endswith('.lnk') else original_path
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
        if os.path.isdir(path) or (path.lower().endswith('.lnk') and os.path.isdir(self.file_operations.resolve_shortcut(path))):
            self.set_folder_thumbnail(label)
        else:
            self.thumbnail_generator.create_thumbnail(path, label)

    def set_folder_thumbnail(self, label):
        label.config(image=self.folder_icon)
        label.image = self.folder_icon

    def handle_file_or_folder_click(self, path):
        if os.path.isdir(path):
            self.navigation.navigate_to(path)
            self.update_navigation_state()
            self.folder_path.set(path)
            self.list_files_and_folders(path)
        elif path.lower().endswith('.lnk'):
            target = self.file_operations.resolve_shortcut(path)
            if target and os.path.exists(target):
                self.handle_file_or_folder_click(target)
            else:
                messagebox.showerror("错误", f"快捷方式目标无效：{target}")
        else:
            self.file_operations.open_file(path)

    def navigate_back(self):
        path = self.navigation.back()
        if path:
            self.update_navigation_state()
            self.navigate_to_path(path)

    def navigate_forward(self):
        path = self.navigation.forward()
        if path:
            self.update_navigation_state()
            self.navigate_to_path(path)

    def navigate_up(self):
        current_path = self.folder_path.get()
        if current_path:
            parent_folder = os.path.dirname(current_path)
            self.navigation.navigate_to(parent_folder)
            self.update_navigation_state()
            self.navigate_to_path(parent_folder)

    def refresh_file_list(self):
        current_path = self.folder_path.get()
        if current_path:
            self.list_files_and_folders(current_path)
            self.root.update_idletasks()

    def on_file_var_check(self, var, orig_name, std_name_var):
        if var.get() == 1 and self.add_date.get():
            std_name_var.set(self.file_operations.generate_standardized_name(
                os.path.join(self.folder_path.get(), orig_name), self.add_date.get(), self.rename_folders.get()))
        else:
            std_name_var.set(orig_name)

    def update_navigation_state(self):
        self.btn_back.config(state=tk.NORMAL if self.navigation.can_go_back() else tk.DISABLED)
        self.btn_forward.config(state=tk.NORMAL if self.navigation.can_go_forward() else tk.DISABLED)
        self.path_entry.set(self.navigation.current_path())

    def rename_files(self):
        selected_files = []
        for var, orig_name, std_name_var in self.file_vars:
            if var.get() == 1:
                selected_files.append((orig_name, std_name_var.get()))
        if selected_files:
            if messagebox.askyesno("确认", "确定要重命名选中的文件吗？"):
                self.file_operations.rename_files(self.folder_path.get(), selected_files)
                self.refresh_file_list()
        else:
            messagebox.showwarning("提示", "没有选择任何文件进行重命名。")

    def undo_rename(self):
        if self.file_operations.undo_rename():
            self.refresh_file_list()
        else:
            messagebox.showinfo("提示", "没有可以撤销的操作。")

    def redo_rename(self):
        if self.file_operations.redo_rename():
            self.refresh_file_list()
        else:
            messagebox.showinfo("提示", "没有可以重做的操作。")

    def update_buttons(self):
        if not self.button_frame:
            return
        for widget in self.button_frame.winfo_children():
            widget.destroy()
        Button(self.button_frame, text="确定", command=self.rename_files).pack(side=tk.LEFT, pady=10, padx=5)
        Button(self.button_frame, text="撤销", command=self.undo_rename).pack(side=tk.LEFT, pady=10, padx=5)
        Button(self.button_frame, text="重做", command=self.redo_rename).pack(side=tk.LEFT, pady=10, padx=5)
