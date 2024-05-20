import ctypes
import os
import subprocess
import tkinter as tk
from tkinter import Button
from tkinter import filedialog, Checkbutton, IntVar, messagebox

import win32com.client

# Import custom modules
from modules.clean_graph import adjust_text_boxes
from modules import clean_master_in_open_presentation
from modules.extract_slides import extract_text_from_presentation
from modules.standardize_filename_v7 import generate_standardized_name


def main():
    app = tk.Tk()
    app.title("PPTX 工具包")

    # 主界面布局配置
    button_frame = tk.Frame(app, padx=20, pady=20)
    text_frame = tk.Frame(app, padx=20, pady=20)

    button_frame.pack(side=tk.LEFT, fill=tk.Y)
    text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    # 文本输出框
    text_output = tk.Text(text_frame, height=15, width=50)
    text_output.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # 功能函数
    def list_open_presentations():
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        return [presentation.Name for presentation in powerpoint.Presentations]

    def open_ppt_selection_window(action):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentations_info = list_open_presentations()

        top = tk.Toplevel(app)
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
                    text_output.delete('1.0', tk.END)
                    text_output.insert(tk.END, "\n".join(result))
                top.destroy()

        btn_select = tk.Button(top, text="选择", command=on_select)
        btn_select.pack(pady=5)

    def is_hidden(filepath):
        """判断文件是否为隐藏文件"""
        attribute = ctypes.windll.kernel32.GetFileAttributesW(filepath)
        return attribute & 2  # FILE_ATTRIBUTE_HIDDEN = 0x2

    def open_file(filepath):
        """使用系统默认程序打开文件"""
        if os.path.isdir(filepath):
            subprocess.run(['explorer', filepath], check=True)
        else:
            os.startfile(filepath)

    def open_rename_files_window():
        top = tk.Toplevel(app)
        top.title("重命名文件")

        file_frame = tk.Frame(top, padx=20, pady=20)
        file_frame.pack(fill=tk.BOTH, expand=True)

        folder_path = tk.StringVar()
        add_date = IntVar()
        rename_folders = IntVar()
        rename_history = []
        redo_history = []

        def open_folder_dialog():
            folder_selected = filedialog.askdirectory()
            if folder_selected:
                folder_path.set(folder_selected)
                label_path.config(text=f"当前目录: {folder_path.get()}")
                list_files_and_folders(folder_selected)

        def list_files_and_folders(folder_selected):
            for widget in canvas_frame.winfo_children():
                widget.destroy()

            # 添加标签
            header_frame = tk.Frame(canvas_frame)
            header_frame.pack(fill='x', pady=2)
            tk.Label(header_frame, text="选择", width=10).pack(side=tk.LEFT, padx=5)
            tk.Label(header_frame, text="原始文件名", width=40).pack(side=tk.LEFT, padx=5)
            tk.Label(header_frame, text="目标文件名", width=40).pack(side=tk.LEFT, padx=5)

            files_and_folders = os.listdir(folder_selected)

            file_vars = []

            for name in files_and_folders:
                # 跳过隐藏文件和文件夹
                original_path = os.path.join(folder_selected, name)
                if is_hidden(original_path) or name.startswith('.'):
                    continue
                if os.path.isdir(original_path) and not rename_folders.get():
                    continue  # 跳过文件夹
                orig_name = name
                std_name = generate_standardized_name(original_path, add_date.get(), rename_folders.get())

                var = IntVar(value=1)
                orig_name_var = tk.StringVar(value=orig_name)
                std_name_var = tk.StringVar(value=std_name)
                file_vars.append((var, orig_name, std_name_var))

                row_frame = tk.Frame(canvas_frame)
                row_frame.pack(fill='x', pady=2)

                Checkbutton(row_frame, variable=var, width=10).pack(side=tk.LEFT)
                orig_entry = tk.Entry(row_frame, width=40, textvariable=orig_name_var, state='readonly')
                orig_entry.pack(side=tk.LEFT, padx=5)
                orig_entry.bind("<Button-1>", lambda e, p=original_path: open_file(p))
                tk.Entry(row_frame, width=40, textvariable=std_name_var).pack(side=tk.LEFT, padx=5)

            def rename_files():
                batch_history = []
                for var, orig_name, std_name_var in file_vars:
                    if var.get():
                        try:
                            orig_path = os.path.join(folder_path.get(), orig_name)
                            new_path = os.path.join(folder_path.get(), std_name_var.get())
                            os.rename(orig_path, new_path)
                            batch_history.append((orig_path, new_path))  # 记录重命名操作
                        except Exception as e:
                            messagebox.showerror("错误", f"重命名 {orig_name} 到 {std_name_var.get()} 失败: {e}")
                if batch_history:
                    rename_history.append(batch_history)
                    if len(rename_history) > 10:
                        rename_history.pop(0)  # 保持最多10个批次的撤销操作
                    redo_history.clear()  # 清空重做历史
                messagebox.showinfo("完成", "文件重命名完成。")
                refresh_file_list()

            def undo_rename():
                if not rename_history:
                    messagebox.showinfo("提示", "没有可以撤销的重命名操作。")
                    return
                last_batch = rename_history.pop()
                for orig_path, new_path in reversed(last_batch):
                    try:
                        os.rename(new_path, orig_path)
                    except Exception as e:
                        messagebox.showerror("错误", f"撤销重命名 {new_path} 到 {orig_path} 失败: {e}")
                redo_history.append(last_batch)
                refresh_file_list()
                update_buttons()

            def redo_rename():
                if not redo_history:
                    messagebox.showinfo("提示", "没有可以重做的重命名操作。")
                    return
                last_batch = redo_history.pop()
                for orig_path, new_path in last_batch:
                    try:
                        os.rename(orig_path, new_path)
                    except Exception as e:
                        messagebox.showerror("错误", f"重做重命名 {orig_path} 到 {new_path} 失败: {e}")
                rename_history.append(last_batch)
                refresh_file_list()
                update_buttons()

            def update_buttons():
                # 清除之前的按钮
                for widget in button_frame.winfo_children():
                    widget.destroy()
                Button(button_frame, text="确定", command=rename_files).pack(side=tk.LEFT, pady=10, padx=5)
                Button(button_frame, text="撤销", command=undo_rename).pack(side=tk.LEFT, pady=10, padx=5)
                if redo_history:
                    Button(button_frame, text="重做", command=redo_rename).pack(side=tk.LEFT, pady=10, padx=5)

            update_buttons()

        def refresh_file_list():
            if folder_path.get():
                list_files_and_folders(folder_path.get())

        label_path = tk.Label(file_frame, text="当前目录: 未选择")
        btn_folder = tk.Button(file_frame, text="选择文件夹", command=open_folder_dialog)
        btn_folder.grid(row=0, column=0, pady=10, sticky='w')
        label_path.grid(row=0, column=1, pady=10, sticky='w')

        Checkbutton(file_frame, text="文件名添加默认日期", variable=add_date, command=refresh_file_list).grid(row=0,
                                                                                                              column=2,
                                                                                                              padx=10,
                                                                                                              sticky='e')
        Checkbutton(file_frame, text="重命名文件夹", variable=rename_folders, command=refresh_file_list).grid(row=0,
                                                                                                              column=3,
                                                                                                              padx=10,
                                                                                                              sticky='e')

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

        button_frame = tk.Frame(top)
        button_frame.pack(pady=10)

    # 文本提取和调整按钮
    Button(button_frame, text="从PPTX提取文本", command=lambda: open_ppt_selection_window('extract')).pack(pady=10,
                                                                                                           anchor='w')
    Button(button_frame, text="调整文本框", command=lambda: open_ppt_selection_window('adjust')).pack(pady=10,
                                                                                                      anchor='w')
    Button(button_frame, text="清理PPTX母版", command=lambda: open_ppt_selection_window('clean_master')).pack(pady=10,
                                                                                                              anchor='w')
    Button(button_frame, text="重命名目录下的文件", command=open_rename_files_window).pack(pady=10, anchor='w')

    app.mainloop()


# Uncomment the following line when the code is ready to be run
main()
