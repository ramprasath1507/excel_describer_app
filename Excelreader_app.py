import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            global data
            data = pd.read_excel(file_path)  
            messagebox.showinfo("File Upload", "File uploaded successfully.")
            describe_data()  
        except Exception as e:
            messagebox.showerror("Error", f"Failed to upload the file: {e}")
    else:
        messagebox.showwarning("File Upload", "No file selected.")

def describe_data():
    if data is not None:
        try:
            global described_data
            described_data = data.describe(include='all')
            messagebox.showinfo("Data Description", "Data described successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to describe the data: {e}")
    else:
        messagebox.showwarning("Data Description", "No data available to describe.")

def download_file():
    if described_data is not None:
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                described_data.to_excel(save_path, index=False)  
                messagebox.showinfo("Download", "Described data saved successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save the file: {e}")
        else:
            messagebox.showwarning("Download", "Save location not specified.")
    else:
        messagebox.showwarning("Download", "No described data to download.")

window = tk.Tk()
window.title("Excel File Describer")


upload_button = tk.Button(window, text="Upload Excel File", command=upload_file, width=40)
upload_button.grid(row=0, column=0, pady=10)

download_button = tk.Button(window, text="Download Described Data", command=download_file, width=40)
download_button.grid(row=1, column=0, pady=10)


data = None
described_data = None

window.mainloop()
