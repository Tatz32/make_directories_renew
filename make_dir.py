import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD


def clear_fields():
    directory_path_entry.delete(0, tk.END)
    file_path_entry.delete(0, tk.END)
    file_path_get_entry.delete(0, tk.END)
    start_row_entry.delete(0, tk.END)
    end_row_entry.delete(0, tk.END)

def drop_file(event):
    file_path_entry.delete(0, tk.END)
    file_path = event.data.strip('{}')  
    file_path_entry.insert(tk.END, file_path)

def drop_directory(event):
    directory_path_entry.delete(0, tk.END)
    directory_path = event.data.strip('{}')  
    directory_path_entry.insert(tk.END, directory_path)

def create_dirs():
    # Get the directory path from the entry field
    directory_path = directory_path_entry.get()
    if not directory_path.endswith(os.sep):  # Check if the path ends with a slash
        directory_path += os.sep  # If not, add a slash

    # Get the file path from the entry field
    file_path = file_path_entry.get()

    # Get the start and end row numbers from the entry fields
    start_row = int(start_row_entry.get()) -2  # Subtract 1 because Python uses 0-indexing
    end_row = int(end_row_entry.get())
    column_number = int(file_path_get_entry.get())

    # Define the row range you want to read
    rows_to_read = list(range(start_row, end_row))

    # Read the Excel file
    df = pd.read_excel(file_path, skiprows=lambda x: x not in rows_to_read)

    # Get the data from the first column
    column_a_data = df.iloc[:, column_number -1]

    # Create new directories based on column A data
    for item in column_a_data:
        directory = directory_path + str(item)
        if not os.path.exists(directory):
            os.makedirs(directory)

    # Show a message box when done
    messagebox.showinfo("Success", "Directories created successfully!")


root = TkinterDnD.Tk()

# Create a label and an entry field to display the directory path
directory_path_label = tk.Label(root, text = "Directory Path：")
directory_path_label.grid(row=0, column=0, pady=10)
directory_path_entry = tk.Entry(root, width = 60)
directory_path_entry.grid(row=0, column=1, pady=10)
directory_path_entry.drop_target_register(DND_FILES)
directory_path_entry.dnd_bind('<<Drop>>', drop_directory)

# Create a label and an entry field to display the file path
file_path_label = tk.Label(root, text = "File Path：")
file_path_label.grid(row=1, column=0, pady=10)
file_path_entry = tk.Entry(root, width = 60)
file_path_entry.grid(row=1, column=1, pady=10)
file_path_entry.drop_target_register(DND_FILES)
file_path_entry.dnd_bind('<<Drop>>', drop_file)

# Create a frame, label, and entry field for column number
file_path_get = tk.Frame(root)
file_path_get.grid(row=2, column=0, columnspan=2, pady=10)
file_path_get_label = tk.Label(file_path_get, text = "Column No.：")
file_path_get_label.pack(side=tk.LEFT)
file_path_get_entry = tk.Entry(file_path_get, width = 15)
file_path_get_entry.pack(side=tk.LEFT)

# Create a frame, label, and entry field for the start row number
start_row_frame = tk.Frame(root)
start_row_frame.grid(row=3, column=0, columnspan=2, pady=10)
start_row_label = tk.Label(start_row_frame, text = "Start：")
start_row_label.pack(side=tk.LEFT)
start_row_entry = tk.Entry(start_row_frame, width = 15)
start_row_entry.pack(side=tk.LEFT)

# Create a frame, label, and entry field for the end row number
end_row_frame = tk.Frame(root)
end_row_frame.grid(row=4, column=0, columnspan=2, pady=10)
end_row_label = tk.Label(end_row_frame, text = "End：")
end_row_label.pack(side=tk.LEFT)
end_row_entry = tk.Entry(end_row_frame, width = 15)
end_row_entry.pack(side=tk.LEFT)

# Create a button to start the directory creation
start_button = tk.Button(root, text = "Generate", width = 20, command = create_dirs, bg='#00ff7f')
start_button.grid(row=5, column=0, columnspan=4, pady=15)  

# Create a button to delete the all data in the fileds
delete_button = tk.Button(root, text = "Delete", width = 20, command = clear_fields, bg='#ff7f7f')
delete_button.grid(row=6, column=0, columnspan=4, pady=15)

root.mainloop()