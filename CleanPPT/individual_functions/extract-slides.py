# extract-slides.py
import tkinter as tk
from tkinter import filedialog
import os
import xml.etree.ElementTree as ET


def extract_text_from_slides(folder_path):
    slides_path = os.path.join(folder_path, 'slides')
    text_content = []

    for slide in os.listdir(slides_path):
        if slide.endswith('.xml'):
            file_path = os.path.join(slides_path, slide)
            tree = ET.parse(file_path)
            root = tree.getroot()
            for elem in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                text_content.append(elem.text)
    return text_content

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        extracted_texts = extract_text_from_slides(folder_path)
        text_output.delete('1.0', tk.END)
        text_output.insert(tk.END, "\n".join(extracted_texts))

app = tk.Tk()
app.title("PPTX Text Extractor")

frame = tk.Frame(app)
frame.pack(pady=20)

btn_browse = tk.Button(frame, text="Select PPTX Folder", command=browse_folder)
btn_browse.pack(side=tk.LEFT, padx=(0, 10))

text_output = tk.Text(frame, height=15, width=50)
text_output.pack(side=tk.LEFT)

app.mainloop()
