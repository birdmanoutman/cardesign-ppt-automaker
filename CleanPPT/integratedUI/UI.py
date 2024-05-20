import tkinter as tk
from tkinter import messagebox, Button, filedialog, Checkbutton, IntVar
import win32com.client

# Import custom modules
from modules.clean_graph import adjust_text_boxes
from modules import clean_master_in_open_presentation
from modules.extract_slides import extract_text_from_presentation
from modules.standardize_filename_v7 import standardize_filename

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

    def open_rename_files_window():
        top = tk.Toplevel(app)
        top.title("重命名文件")

        file_frame = tk.Frame(top, padx=20, pady=20)
        file_frame.pack()

        folder_path = tk.StringVar()
        add_date = IntVar()
        rename_folders = IntVar()

        def open_folder_dialog():
            folder_selected = filedialog.askdirectory()
            if folder_selected:
                folder_path.set(folder_selected)
                label_path.config(text=f"当前目录: {folder_path.get()}")

        label_path = tk.Label(file_frame, text="未选择目录")
        btn_folder = Button(file_frame, text="选择文件夹", command=open_folder_dialog)
        btn_folder.pack(side=tk.LEFT, pady=10)
        label_path.pack(side=tk.LEFT, padx=10)

        Checkbutton(file_frame, text="文件名添加默认日期", variable=add_date).pack()
        Checkbutton(file_frame, text="重命名文件夹", variable=rename_folders).pack()
        Button(file_frame, text="标准化文件名", command=lambda: standardize_filename(folder_path.get(), add_date.get(), rename_folders.get())).pack(pady=10)

    # 文本提取和调整按钮
    Button(button_frame, text="从PPTX提取文本", command=lambda: open_ppt_selection_window('extract')).pack(pady=10, anchor='w')
    Button(button_frame, text="调整文本框", command=lambda: open_ppt_selection_window('adjust')).pack(pady=10, anchor='w')
    Button(button_frame, text="清理PPTX母版", command=lambda: open_ppt_selection_window('clean_master')).pack(pady=10, anchor='w')
    Button(button_frame, text="重命名目录下的文件", command=open_rename_files_window).pack(pady=10, anchor='w')

    app.mainloop()

# Uncomment the following line when the code is ready to be run
main()


