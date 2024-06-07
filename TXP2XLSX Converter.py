import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class TXP2XLSXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("TXP2XLSX Converter")
        self.create_widgets()

    def create_widgets(self):
        self.select_button = tk.Button(self.root, text="Select Directory", command=self.select_directory)
        self.select_button.pack(pady=20)

    def txp_to_xlsx(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            data = [line.strip().split('\t') for line in lines]
            df = pd.DataFrame(data[1:], columns=data[0])
            output_file = file_path.replace('.txp', '.xlsx')
            df.to_excel(output_file, index=False)
            print(f"Converted {file_path} to {output_file}")
        except Exception as e:
            print(f"Failed to convert {file_path}: {e}")
            messagebox.showerror("Conversion Error", f"Failed to convert {file_path}: {e}")

    def convert_all_txp_in_directory(self, directory):
        try:
            for filename in os.listdir(directory):
                if filename.endswith('.txp'):
                    file_path = os.path.join(directory, filename)
                    self.txp_to_xlsx(file_path)
            messagebox.showinfo("Success", "All files have been successfully converted.")
        except Exception as e:
            print(f"Error in converting files in directory {directory}: {e}")
            messagebox.showerror("Directory Error", f"Error in converting files in directory {directory}: {e}")

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.convert_all_txp_in_directory(directory)

if __name__ == "__main__":
    root = tk.Tk()
    app = TXP2XLSXConverter(root)
    root.mainloop()
