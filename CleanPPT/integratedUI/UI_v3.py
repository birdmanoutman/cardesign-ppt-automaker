import ctypes
import os
import subprocess
import tkinter as tk
from tkinter import Button, Checkbutton, IntVar, filedialog, messagebox, ttk

import win32com.client

# Import custom modules
from clean_graph import adjust_text_boxes
from clean_master import clean_master_in_open_presentation
from extract_slides import extract_text_from_presentation
from standardize_filename_v7 import generate_standardized_name

class PPTXToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX 工具包")

        self.text_output = None
        self.add_date = IntVar()
        self.rename_folders = IntVar()
        self.folder_path = tk.StringVar()
        self.rename_history = []
        self.redo_history = []
        self.file_vars = []

        self.setup_ui()

    def setup_ui(self):
        # 主界面布局配置
        button_frame = tk.Frame(self.root, padx=20, pady=20)
        text_frame = tk.Frame(self.root, padx=20, pady=20)

        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 文本输出框
        self.text_output = tk.Text(text_frame, height=15, width=50)
        self.text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # 功能按钮
        Button(button_frame, text="从PPTX提取文本", command=lambda: self.open_ppt_selection_window('extract')).pack(pady=10, anchor='w')
        Button(button_frame, text="调整文本框", command=lambda: self.open_ppt_selection_window('adjust')).pack(pady=10, anchor='w')
        Button(button_frame, text="清理PPTX母版", command=lambda: self.open_ppt_selection_window('clean_master')).pack(pady=10, anchor='w')
        Button(button_frame, text="重命名目录下的文件", command=self.open_rename_files_window).pack(pady=10, anchor='w')

    def list_open_presentations(self):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        return [presentation.Name for presentation in powerpoint.Presentations]

    def open_ppt_selection_window(self, action):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentations_info = self.list_open_presentations()

        top = tk.Toplevel(self.root)
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
                    self.text_output.delete('1.0', tk.END)
                    self.text_output.insert(tk.END, "\n".join(result))
                top.destroy()

        btn_select = tk.Button(top, text="选择", command=on_select)
        btn_select.pack(pady=5)

    @staticmethod
    def is_hidden(filepath):
        """判断文件是否为隐藏文件"""
        attribute = ctypes.windll.kernel32.GetFileAttributesW(filepath)
        return attribute & 2  # FILE_ATTRIBUTE_HIDDEN = 0x2

    @staticmethod
    def open_file(filepath):
        """使用系统默认程序打开文件"""
        if os.path.isdir(filepath):
            subprocess.run(['explorer', filepath], check=True)
        else:
            os.startfile(filepath)

    def open_rename_files_window(self):
        top = tk.Toplevel(self.root)
        top.title("重命名文件")

        file_frame = tk.Frame(top, padx=20, pady=20)
        file_frame.pack(fill=tk.BOTH, expand=True)

        label_path = tk.Label(file_frame, text="当前目录: 未选择")
        btn_folder = tk.Button(file_frame, text="选择文件夹", command=self.open_folder_dialog)
        btn_folder.grid(row=0, column=0, pady=10, sticky='w')
        label_path.grid(row=0, column=1, pady=10, sticky='w')

        Checkbutton(file_frame, text="文件名添加默认日期", variable=self.add_date, command=self.refresh_file_list).grid(row=0, column=2, padx=10, sticky='e')
        Checkbutton(file_frame, text="重命名文件夹", variable=self.rename_folders, command=self.refresh_file_list).grid(row=0, column=3, padx=10, sticky='e')

        # 添加滚动条
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
        self.label_path = label_path
        self.button_frame = tk.Frame(top)
        self.button_frame.pack(pady=10)

    def open_folder_dialog(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.label_path.config(text=f"当前目录: {self.folder_path.get()}")
            self.list_files_and_folders(folder_selected)

    def list_files_and_folders(self, folder_selected):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()

        # 添加标签
        header_frame = tk.Frame(self.canvas_frame)
        header_frame.pack(fill='x', pady=2)
        tk.Label(header_frame, text="选择", width=10).pack(side=tk.LEFT, padx=5)
        tk.Label(header_frame, text="原始文件名", width=40).pack(side=tk.LEFT, padx=5)
        tk.Label(header_frame, text="目标文件名", width=40).pack(side=tk.LEFT, padx=5)

        files_and_folders = os.listdir(folder_selected)

        self.file_vars = []

        for name in files_and_folders:
            # 跳过隐藏文件和文件夹
            original_path = os.path.join(folder_selected, name)
            if self.is_hidden(original_path) or name.startswith('.'):
                continue
            if os.path.isdir(original_path) and not self.rename_folders.get():
                continue  # 跳过文件夹
            orig_name = name
            std_name = generate_standardized_name(original_path, self.add_date.get(), self.rename_folders.get())

            var = IntVar(value=1)
            orig_name_var = tk.StringVar(value=orig_name)
            std_name_var = tk.StringVar(value=std_name)
            self.file_vars.append((var, orig_name, std_name_var))

            row_frame = tk.Frame(self.canvas_frame)
            row_frame.pack(fill='x', pady=2)

            Checkbutton(row_frame, variable=var, width=10).pack(side=tk.LEFT)
            orig_entry = tk.Entry(row_frame, width=40, textvariable=orig_name_var, state='readonly')
            orig_entry.pack(side=tk.LEFT, padx=5)
            orig_entry.bind("<Button-1>", lambda e, p=original_path: self.open_file(p))
            tk.Entry(row_frame, width=40, textvariable=std_name_var).pack(side=tk.LEFT, padx=5)

        self.update_buttons()

    def rename_files(self):
        batch_history = []
        for var, orig_name, std_name_var in self.file_vars:
            if var.get():
                try:
                    orig_path = os.path.join(self.folder_path.get(), orig_name)
                    new_path = os.path.join(self.folder_path.get(), std_name_var.get())
                    os.rename(orig_path, new_path)
                    batch_history.append((orig_path, new_path))  # 记录重命名操作
                except Exception as e:
                    messagebox.showerror("错误", f"重命名 {orig_name} 到 {std_name_var.get()} 失败: {e}")
        if batch_history:
            self.rename_history.append(batch_history)
            if len(self.rename_history) > 10:
                self.rename_history.pop(0)  # 保持最多10个批次的撤销操作
            self.redo_history.clear()  # 清空重做历史
        messagebox.showinfo("完成", "文件重命名完成。")
        self.refresh_file_list()

    def undo_rename(self):
        if not self.rename_history:
            messagebox.showinfo("提示", "没有可以撤销的重命名操作。")
            return
        last_batch = self.rename_history.pop()
        for orig_path, new_path in reversed(last_batch):
            try:
                os.rename(new_path, orig_path)
            except Exception as e:
                messagebox.showerror("错误", f"撤销重命名 {new_path} 到 {orig_path} 失败: {e}")
        self.redo_history.append(last_batch)
        self.refresh_file_list()
        self.update_buttons()

    def redo_rename(self):
        if not self.redo_history:
            messagebox.showinfo("提示", "没有可以重做的重命名操作。")
            return
        last_batch = self.redo_history.pop()
        for orig_path, new_path in last_batch:
            try:
                os.rename(orig_path, new_path)
            except Exception as e:
                messagebox.showerror("错误", f"重做重命名 {orig_path} 到 {new_path} 失败: {e}")
        self.rename_history.append(last_batch)
        self.refresh_file_list()
        self.update_buttons()

    def update_buttons(self):
        # 清除之前的按钮
        for widget in self.button_frame.winfo_children():
            widget.destroy()
        Button(self.button_frame, text="确定", command=self.rename_files).pack(side=tk.LEFT, pady=10, padx=5)
        Button(self.button_frame, text="撤销", command=self.undo_rename).pack(side=tk.LEFT, pady=10, padx=5)
        if self.redo_history:
            Button(self.button_frame, text="重做", command=self.redo_rename).pack(side=tk.LEFT, pady=10, padx=5)

    def refresh_file_list(self):
        if self.folder_path.get():
            self.list_files_and_folders(self.folder_path.get())


def main():
    root = tk.Tk()
    app = PPTXToolApp(root)
    root.mainloop()


# Uncomment the following line when the code is ready to be run
main()
