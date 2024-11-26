# app.py
import tkinter as tk
from modules.ui_components import PPTXToolUI

class PPTXToolApp:
    def __init__(self, root):
        self.ui = PPTXToolUI(root)

def main():
    root = tk.Tk()
    app = PPTXToolApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
