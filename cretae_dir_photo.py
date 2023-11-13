import os
import shutil
import tempfile
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
def load_excel():
    global excel_data, columns_listbox
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            excel_data = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", "Could not load Excel file: {}".format(e))
            return
        columns_listbox.delete(0, tk.END)
        for col in excel_data.columns:
            columns_listbox.insert(tk.END, col)
        messagebox.showinfo("Success", "Excel file loaded successfully")

def add_field():
    selected_fields_listbox.insert(tk.END, columns_listbox.get(tk.ANCHOR))

# Global variable to store the path of the temporary directory
# Create a temporary directory at the module level
temporary_directory = tempfile.TemporaryDirectory()
# Store the path of the temporary directory
temporary_directory_path = temporary_directory.name

def create_directories():
    global temporary_directory_path
    selected_fields = list(selected_fields_listbox.get(0, tk.END))
    result_listbox.delete(0, tk.END)  # Clear previous results
    for _, row in excel_data.iterrows():
        path_components = [temporary_directory_path] + [str(row[field]) for field in selected_fields]
        dir_path = os.path.join(*path_components)
        try:
            os.makedirs(dir_path, exist_ok=True)
            result_listbox.insert(tk.END, dir_path)  # Display each created directory path
        except Exception as e:
            messagebox.showerror("Error", "Could not create directory: {}".format(e))
            return
    messagebox.showinfo("Success", f"Directories created successfully in temporary location: {temporary_directory_path}")

def export_structure():
    global temporary_directory_path
    # Use the path of the temporary directory for exporting
    if os.path.isdir(temporary_directory_path):
        archive_path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
        if archive_path:
            shutil.make_archive(base_name=archive_path.rstrip('.zip'), format='zip', root_dir=temporary_directory_path)
            messagebox.showinfo("Success", "Directory structure exported as zip")
    else:
        messagebox.showerror("Error", "No directory structure available to export.")

# Create the main window
root = tk.Tk()
root.title("Directory Structure Creator")

# Use a single frame for organization and improved layout control
frame = tk.Frame(root)
frame.pack(padx=10, pady=10, fill='both', expand=True)

# Load Excel file button
button_load_excel = tk.Button(frame, text="Load Excel File", command=load_excel)
button_load_excel.pack(fill='x')

# Listbox for Excel columns with a scrollbar
columns_scrollbar = tk.Scrollbar(frame, orient='vertical')
columns_listbox = Listbox(frame, yscrollcommand=columns_scrollbar.set)
columns_scrollbar.config(command=columns_listbox.yview)
columns_scrollbar.pack(side='right', fill='y')
columns_listbox.pack(side='left', fill='both', expand=True)

# Button to add field for directory creation
button_add_field = tk.Button(frame, text="Add Field", command=add_field)
button_add_field.pack(fill='x')

# Listbox for selected fields with a scrollbar
selected_fields_scrollbar = tk.Scrollbar(frame, orient='vertical')
selected_fields_listbox = Listbox(frame, yscrollcommand=selected_fields_scrollbar.set)
selected_fields_scrollbar.config(command=selected_fields_listbox.yview)
selected_fields_scrollbar.pack(side='right', fill='y')
selected_fields_listbox.pack(side='left', fill='both', expand=True)

# Button to remove a field from directory creation
button_remove_field = tk.Button(frame, text="Remove Field", command=lambda: selected_fields_listbox.delete(tk.ANCHOR))
button_remove_field.pack(fill='x')

# Entry for root path
entry_root_path = tk.Entry(frame)
entry_root_path.pack(fill='x')
entry_root_path.insert(0, "heritage_s_2_new")

# Create directories button
button_create_dirs = tk.Button(frame, text="Create Directories", command=create_directories)
button_create_dirs.pack(fill='x')

# Export as ZIP button
button_export_zip = tk.Button(frame, text="Export as ZIP", command=export_structure)
button_export_zip.pack(fill='x')

# Listbox for displaying the result with a scrollbar
result_scrollbar = tk.Scrollbar(frame, orient='vertical')
result_listbox = Listbox(frame, yscrollcommand=result_scrollbar.set)
result_scrollbar.config(command=result_listbox.yview)
result_scrollbar.pack(side='right', fill='y')
result_listbox.pack(side='left', fill='both', expand=True)

# Configure the grid layout weights (removed unnecessary configurations)
frame.grid_rowconfigure(0, weight=1)
frame.grid_columnconfigure(0, weight=1)

# Run the application
root.mainloop()
# Clean up the temporary directory when the program exits
temporary_directory.cleanup()