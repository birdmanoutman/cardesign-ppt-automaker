import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import findBackgroundIMG_V3  # Ensure this is the modified version as shown previously

class App:
    def __init__(self, master):
        self.master = master
        self.master.title("CSV Importer and Image Organizer")

        # Database Configuration
        tk.Label(master, text="Database Host:").grid(row=0)
        tk.Label(master, text="Database User:").grid(row=1)
        tk.Label(master, text="Database Password:").grid(row=2)
        tk.Label(master, text="Database Name:").grid(row=3)

        self.db_host = tk.Entry(master)
        self.db_user = tk.Entry(master)
        self.db_password = tk.Entry(master)
        self.db_name = tk.Entry(master)

        self.db_host.grid(row=0, column=1)
        self.db_user.grid(row=1, column=1)
        self.db_password.grid(row=2, column=1)
        self.db_name.grid(row=3, column=1)

        # CSV File Selection
        tk.Label(master, text="CSV File Path:").grid(row=4)
        self.csv_file_path = tk.Entry(master)
        self.csv_file_path.grid(row=4, column=1)
        tk.Button(master, text="Browse", command=self.browse_csv).grid(row=4, column=2)

        # Source and Destination Folders
        tk.Label(master, text="Source Folder:").grid(row=5)
        tk.Label(master, text="Destination Folder:").grid(row=6)

        self.src_folder_path = tk.Entry(master)
        self.dest_folder_path = tk.Entry(master)

        self.src_folder_path.grid(row=5, column=1)
        self.dest_folder_path.grid(row=6, column=1)

        tk.Button(master, text="Browse", command=lambda: self.browse_folder(self.src_folder_path)).grid(row=5, column=2)
        tk.Button(master, text="Browse", command=lambda: self.browse_folder(self.dest_folder_path)).grid(row=6, column=2)

        # Start Button
        tk.Button(master, text="Start Processing", command=self.start_processing).grid(row=7, column=1)

        # Console/Log Area
        self.log_area = scrolledtext.ScrolledText(master, height=10)
        self.log_area.grid(row=8, column=0, columnspan=3)

    def browse_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.csv_file_path.delete(0, tk.END)
        self.csv_file_path.insert(0, file_path)

    def browse_folder(self, entry_field):
        folder_path = filedialog.askdirectory()
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_path)

    def start_processing(self):
        # Extract information from the GUI
        db_config = {
            'host': self.db_host.get(),
            'user': self.db_user.get(),
            'passwd': self.db_password.get(),
            'database': self.db_name.get()
        }
        src_folder = self.src_folder_path.get()
        dest_folder = self.dest_folder_path.get()
        csv_file_path = self.csv_file_path.get()

        # Run the processing in a separate thread
        thread = threading.Thread(target=self.process_data, args=(db_config, src_folder, dest_folder, csv_file_path))
        thread.daemon = True
        thread.start()

    def process_data(self, db_config, src_folder, dest_folder, csv_file_path):
        # Create an instance of the ImageProcessor and run it
        try:
            processor = findBackgroundIMG_V3.ImageProcessor(db_config, src_folder, dest_folder, csv_file_path)
            processor.run()
            self.log("Process completed successfully.")
        except Exception as e:
            self.log(f"An error occurred: {str(e)}")

    def log(self, message):
        if self.log_area.winfo_exists():
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.see(tk.END)

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
