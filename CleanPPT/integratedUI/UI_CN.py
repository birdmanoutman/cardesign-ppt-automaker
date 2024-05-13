import tkinter as tk
from tkinter import messagebox, Button

import win32com.client
import win32com.client

from clean_graph import adjust_text_boxes
from extract_slides import extract_text_from_presentation
from clean_master import clean_master_in_open_presentation

app = tk.Tk()
app.title("PPTX Toolkit")

# 创建一个框架用于放置按钮，这个框架将放在左侧
button_frame = tk.Frame(app)
button_frame.pack(side=tk.LEFT, padx=20, pady=20)

# 创建另一个框架用于放置文本输出框，这个框架将放在右侧
text_frame = tk.Frame(app)
text_frame.pack(side=tk.RIGHT, padx=20, pady=20)


def list_open_presentations():
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentations_info = []
    for presentation in powerpoint.Presentations:
        presentations_info.append(presentation.Name)
    return presentations_info


def open_ppt_selection_window(action):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    presentations_info = list_open_presentations()

    top = tk.Toplevel(app)
    top.title("Select Presentation")

    if not presentations_info:
        messagebox.showinfo("Info", "No open presentations found.")
        top.destroy()
        return

    lbl = tk.Label(top, text="Select a presentation:")
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
            if action == 'extract':
                extracted_texts = extract_text_from_presentation(presentation)
                text_output.delete('1.0', tk.END)
                text_output.insert(tk.END, "\n".join(extracted_texts))
            elif action == 'adjust':
                adjust_text_boxes(presentation)
            elif action == 'clean_master':
                clean_master_in_open_presentation(presentation)
            top.destroy()

    btn_select = tk.Button(top, text="Select", command=on_select)
    btn_select.pack(pady=5)



def extract_text():
    open_ppt_selection_window('extract')


def adjust_boxes():
    open_ppt_selection_window('adjust')


# 创建并配置按钮，增加间距并左对齐
btn_extract = Button(button_frame, text="提 取 文 字 ~~", command=extract_text)
btn_extract.pack(pady=20, anchor='w')  # 增加pady值以增加间距，并使用anchor='w'实现左对齐

btn_adjust = Button(button_frame, text="不！要！拉！框！", command=adjust_boxes)
btn_adjust.pack(pady=20, anchor='w')  # 同上，保持风格一致

# 添加新按钮以清理母版
btn_clean_master = tk.Button(button_frame, text="干净的母版！！！", command=lambda: open_ppt_selection_window('clean_master'))
btn_clean_master.pack(pady=20, anchor='w')


# 创建并将文本输出框添加到文本框架
text_output = tk.Text(text_frame, height=15, width=50)
text_output.pack(padx=10)

app.mainloop()
