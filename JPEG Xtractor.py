import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

# Function to get the desktop path
def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), "Desktop")

# Function to read Excel file and extract file names from columns J, K, and L
def read_excel(file_path):
    try:
        # Load Excel file into a DataFrame
        df = pd.read_excel(file_path, sheet_name=0, dtype=str)  # Read everything as strings

        # Flatten all cell values into a list
        all_values = df.values.flatten()

        # Normalize filenames: strip spaces, convert to lowercase
        file_names = [str(val).strip().lower() for val in all_values if isinstance(val, str) and val.strip().lower().endswith('.jpeg')]

        return file_names
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {e}")

# Function to search and copy photos based on file names
def copy_files(file_names, source_dir, destination_dir, log_widget, count_label, progress_bar):
    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)

    # Normalize available file names in source directory
    available_files = {f.lower().strip(): f for f in os.listdir(source_dir) if f.lower().endswith('.jpeg')}
    total_files = len(file_names)
    
    log_widget.insert(tk.END, f"Available JPEGs in source directory: {len(available_files)}\n")

    copied_count = 0
    progress_bar["maximum"] = total_files  # Set progress bar limit

    for index, file_name in enumerate(file_names):
        normalized_name = file_name.strip().lower()

        if normalized_name in available_files:
            source_path = os.path.join(source_dir, available_files[normalized_name])
            destination_path = os.path.join(destination_dir, available_files[normalized_name])

            shutil.copy2(source_path, destination_path)
            copied_count += 1
            log_widget.insert(tk.END, f"Copied: {available_files[normalized_name]}\n")
        else:
            log_widget.insert(tk.END, f"Not found: {file_name}\n")

        # Update progress bar
        progress_bar["value"] = index + 1
        progress_bar.update_idletasks()  # Forces UI refresh

        log_widget.yview(tk.END)  # Auto-scroll log

    count_label.config(text=f"Total JPEGs copied: {copied_count}")

# Function to select a file using a file dialog
def select_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")], initialdir=get_desktop_path())
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

# Function to select a directory using a directory dialog
def select_source(entry_widget):
    dir_path = filedialog.askdirectory(initialdir=get_desktop_path())
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, dir_path)

# Function to select a directory using a directory dialog
def select_directory(entry_widget):
    dir_path = filedialog.askdirectory(initialdir=get_desktop_path())
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, dir_path)

# Function to start the file copying process
def start_process(excel_entry, source_entry, dest_entry, log_widget, count_label, progress_bar):
    excel_path = excel_entry.get()
    source_dir = source_entry.get()
    dest_dir = dest_entry.get()
    
    if not all([excel_path, source_dir, dest_dir]):
        messagebox.showerror("Error", "Please select all required paths.")
        return
    
    try:
        file_names = read_excel(excel_path)
        log_widget.delete(1.0, tk.END)  # Clear log
        progress_bar["value"] = 0  # Reset progress bar

        # Get the Excel file name
        excel_file_name = os.path.basename(excel_path)

        # Replace "EXCEL" with "PHOTOS" in the file name
        folder_name = excel_file_name.replace("EXCEL", "PHOTOS")
        folder_name = os.path.splitext(folder_name)[0]  # Remove the file extension

        # Create the destination folder
        dest_folder_path = os.path.join(dest_dir, folder_name)
        if not os.path.exists(dest_folder_path):
            os.makedirs(dest_folder_path)

        copy_files(file_names, source_dir, dest_folder_path, log_widget, count_label, progress_bar)
        messagebox.showinfo("Success", "File copying process completed.")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# Function to create the GUI
def create_gui():
    root = tk.Tk()
    root.title("JPEG Xtractor")
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width / 2) - (550 / 2)
    y = (screen_height / 2) - (320 / 2)
    root.geometry(f"550x320+{int(x)}+{int(y)}")
    root.resizable(False, False)

    # Set the custom icon (you can choose either the .ico or .png method)
    root.iconbitmap("C:\\Users\\john\\Desktop\\Projects\\jpeg.ico")  # Use a .ico file
    # OR
    #logo = tk.PhotoImage(file="C:\\Users\\john\\Desktop\\Projects\\logo.png")  # Use a .png file
    #root.tk.call('wm', 'iconphoto', root._w, logo)
    
    # Excel file selection
    frame1 = tk.Frame(root)
    frame1.pack(pady=5)
    tk.Label(frame1, text="Excel File:", width=20, anchor="w").pack(side=tk.LEFT)
    excel_entry = tk.Entry(frame1, width=40)
    excel_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(frame1, text="Browse", command=lambda: select_file(excel_entry)).pack(side=tk.LEFT)

    # Source directory selection
    frame2 = tk.Frame(root)
    frame2.pack(pady=5)
    tk.Label(frame2, text="Source Folder:", width=20, anchor="w").pack(side=tk.LEFT)
    source_entry = tk.Entry(frame2, width=40)
    source_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(frame2, text="Browse", command=lambda: select_source(source_entry)).pack(side=tk.LEFT)
    
    # Destination directory selection
    frame3 = tk.Frame(root)
    frame3.pack(pady=5)
    tk.Label(frame3, text="Destination Folder:", width=20, anchor="w").pack(side=tk.LEFT)
    dest_entry = tk.Entry(frame3, width=40)
    dest_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(frame3, text="Browse", command=lambda: select_directory(dest_entry)).pack(side=tk.LEFT)

    # Log widget to display progress
    log_widget = scrolledtext.ScrolledText(root, height=6, width=60)  # Reduced height to make it smaller
    log_widget.pack(pady=5)

    # Label to display copied file count
    count_label = tk.Label(root, text="Total JPEGs copied: 0")
    count_label.pack()

    # Progress Bar
    progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
    progress_bar.pack(pady=5)

    # Start button
    tk.Button(root, text="Start", command=lambda: start_process(excel_entry, source_entry, dest_entry, log_widget, count_label, progress_bar)).pack(pady=10)
    
    root.mainloop()

    


if __name__ == "__main__":
    create_gui()
