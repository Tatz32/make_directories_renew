import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pickle
from tkinter import ttk



# Constants
DEFAULT_PADY = 10
ENTRY_WIDTH = 60
BUTTON_WIDTH = 20
GENERATE_BG = '#00ff7f'
DELETE_BG = '#ff7f7f'
PREVIEW_BG = '#7fffd4'
HISTORY_FILE = "history.pkl"


# Initialize history list
history = []


def directory_operation():
    directory_path = directory_path_entry.get()
    file_path = file_path_entry.get()
    start_row = start_row_entry.get()
    end_row = end_row_entry.get()
    column_number = file_path_get_entry.get()

    if not validate_file_path(file_path) or not validate_directory_path(directory_path):
        return

    if not validate_rows(start_row, end_row) or not validate_column_number(file_path, column_number):
        return

    start_row = int(start_row) - 2
    end_row = int(end_row)
    column_number = int(column_number)

    rows_to_read = list(range(start_row, end_row))

    if file_path.lower().endswith('.csv'):
        df = pd.read_csv(file_path, skiprows=lambda x: x not in rows_to_read)
    else:
        df = pd.read_excel(file_path, skiprows=lambda x: x not in rows_to_read)

    column_data = df.iloc[:, column_number - 1]

    create_directories(directory_path, column_data)

    add_to_history(file_path, directory_path)

    messagebox.showinfo("Success", "Directories created successfully!")

def operation_handler():
    operation = operation_var.get()
    if operation == 'Directory':
        directory_operation()
    elif operation == 'Documentation':
        print("Documentation operation not implemented yet.")
    elif operation == 'Excel':
        print("Excel operation not implemented yet.")
    elif operation == 'PDF':
        print("PDF operation not implemented yet.")

def load_history():
    global history
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'rb') as f:
            history = pickle.load(f)
    else:
        history = []

def save_history():
    with open(HISTORY_FILE, 'wb') as f:
        pickle.dump(history, f)

def add_to_history(file_path, directory_path):
    global history
    history.append((file_path, directory_path))
    save_history()

def clear_fields():
    entries = [directory_path_entry, file_path_entry, file_path_get_entry, start_row_entry, end_row_entry]
    for entry in entries:
        entry.delete(0, tk.END)

def handle_drop(event, entry):
    entry.delete(0, tk.END)
    path = event.data.strip('{}')  
    entry.insert(tk.END, path)

def create_directories(directory_path, column_data):
    if not directory_path.endswith(os.sep):  # Check if the path ends with a slash
        directory_path += os.sep  # If not, add a slash

    for item in column_data:
        directory = directory_path + str(item)
        if not os.path.exists(directory):
            os.makedirs(directory)


def validate_file_path(file_path):
    if not os.path.isfile(file_path):
        messagebox.showerror("Error", "The provided file does not exist.")
        return False
    if not os.access(file_path, os.R_OK):
        messagebox.showerror("Error", "The provided file is not readable.")
        return False
    return True

def validate_directory_path(directory_path):
    if not os.path.isdir(directory_path):
        messagebox.showerror("Error", "The provided directory does not exist.")
        return False
    if not os.access(directory_path, os.W_OK):
        messagebox.showerror("Error", "The provided directory is not writable.")
        return False
    return True

def validate_rows(start_row, end_row):
    try:
        start_row = int(start_row)
        end_row = int(end_row)
    except ValueError:
        messagebox.showerror("Error", "Start and End rows must be integers.")
        return False

    if start_row < 1 or end_row < 1 or end_row <= start_row:
        messagebox.showerror("Error", "Start row must be less than End row and both must be greater than 0.")
        return False

    return True

def validate_column_number(file_path, column_number):
    try:
        column_number = int(column_number)
    except ValueError:
        messagebox.showerror("Error", "Column number must be an integer.")
        return False

    if file_path.lower().endswith('.csv'):
        df = pd.read_csv(file_path, nrows=1)
    else:
        df = pd.read_excel(file_path, nrows=1)

    if column_number < 1 or column_number > len(df.columns):
        messagebox.showerror("Error", f"Column number must be between 1 and {len(df.columns)}.")
        return False

    return True

def create_dirs():
    directory_path = directory_path_entry.get()
    file_path = file_path_entry.get()
    start_row = start_row_entry.get()
    end_row = end_row_entry.get()
    column_number = file_path_get_entry.get()

    if not validate_file_path(file_path) or not validate_directory_path(directory_path):
        return

    if not validate_rows(start_row, end_row) or not validate_column_number(file_path, column_number):
        return

    start_row = int(start_row) - 2
    end_row = int(end_row)
    column_number = int(column_number)

    rows_to_read = list(range(start_row, end_row))

    if file_path.lower().endswith('.csv'):
        df = pd.read_csv(file_path, skiprows=lambda x: x not in rows_to_read)
    else:
        df = pd.read_excel(file_path, skiprows=lambda x: x not in rows_to_read)

    column_data = df.iloc[:, column_number - 1]

    create_directories(directory_path, column_data)

    add_to_history(file_path, directory_path)

    messagebox.showinfo("Success", "Directories created successfully!")

def preview_data():
    file_path = file_path_entry.get()
    if not validate_file_path(file_path):
        return

    if file_path.lower().endswith('.csv'):
        df = pd.read_csv(file_path, nrows=5)
    else:
        df = pd.read_excel(file_path, nrows=5)

    simpledialog.messagebox.showinfo("Data Preview", df.to_string())

def open_directory_dialog():
    directory_path = filedialog.askdirectory()
    directory_path_entry.delete(0, tk.END)
    directory_path_entry.insert(tk.END, directory_path)

def open_file_dialog():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'), ('CSV Files', '*.csv')])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(tk.END, file_path)

def create_dnd_entry(parent, row, label_text, drop_func):
    label = tk.Label(parent, text = label_text)
    label.grid(row=row, column=0, pady=DEFAULT_PADY)
    
    entry = tk.Entry(parent, width = ENTRY_WIDTH)
    entry.grid(row=row, column=1, pady=DEFAULT_PADY)
    entry.drop_target_register(DND_FILES)
    entry.dnd_bind('<<Drop>>', lambda event: drop_func(event, entry))

    return entry

#def operation_handler():
 #   operation = operation_var.get()
  #  if operation == 'Directory':
   #     create_dirs()
    #elif operation == 'Documentation':
     #   print("Documentation operation not implemented yet.")
    #elif operation == 'Excel':
     #   generate_excel()
    #elif operation == 'PDF':
     #   print("PDF operation not implemented yet.")

load_history()

root = TkinterDnD.Tk()

# Dropdown menu for operations
operation_var = tk.StringVar(root)
operation_var.set("Directory")  # default value
operation_combobox = ttk.Combobox(root, textvariable=operation_var, values=['Directory', 'Documentation', 'Excel', 'PDF'])
operation_combobox.grid(row=0, column=0, pady=DEFAULT_PADY)

directory_path_entry = create_dnd_entry(root, 1, "Directory Path：", handle_drop)
dir_open_button = tk.Button(root, text="Open", command=open_directory_dialog)
dir_open_button.grid(row=1, column=2, pady=DEFAULT_PADY)

file_path_entry = create_dnd_entry(root, 2, "File Path：", handle_drop)
file_open_button = tk.Button(root, text="Open", command=open_file_dialog)
file_open_button.grid(row=2, column=2, pady=DEFAULT_PADY)


start_row_entry = create_dnd_entry(root, 3, "Start：", handle_drop)

end_row_entry = create_dnd_entry(root, 4, "End：", handle_drop)

file_path_get_entry = create_dnd_entry(root, 5, "Column No.：", handle_drop)

preview_button = tk.Button(root, text="Preview", width=BUTTON_WIDTH, command=preview_data, bg=PREVIEW_BG)
preview_button.grid(row=6, column=0, columnspan=2, pady=DEFAULT_PADY)

generate_button = tk.Button(root, text="Generate", width=BUTTON_WIDTH, command=operation_handler, bg=GENERATE_BG)
generate_button.grid(row=7, column=0, columnspan=2, pady=DEFAULT_PADY)

delete_button = tk.Button(root, text="Delete", width=BUTTON_WIDTH, command=clear_fields, bg=DELETE_BG)
delete_button.grid(row=8, column=0, columnspan=2, pady=DEFAULT_PADY)

root.mainloop()