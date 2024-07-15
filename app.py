import tkinter as tk
from tkinter import filedialog
import openpyxl
import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD

filename_to_path = {}

# Function to handle file selection
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    for file in files:
        filename = os.path.basename(file) 
        listbox.insert(tk.END, filename)
        filename_to_path[filename] = file

def rename_files():
        for filename in listbox.get(0, tk.END):
            file_path = filename_to_path[filename]  # Retrieve the full path
            workbook = openpyxl.load_workbook(file_path)
            if mode_var.get() == 1:
                sheet = workbook.worksheets[0]
                cell_value = sheet['B5'].value.split(" ")
                if cell_value[0].startswith("ม."):
                    fileName = cell_value[0].replace("ม.", "m")
                elif cell_value[0].startswith("ป."):
                    fileName = cell_value[0].replace("ป.", "p")
                fileName = fileName.replace("/", "-")
                
            if not fileName.endswith('.xlsx'):
                fileName += '.xlsx'
            directory = os.path.dirname(file_path)
            new_file_path = os.path.join(directory, fileName)

            # Rename the file
            os.rename(file_path, new_file_path)

            print(f"Renamed '{file_path}' to '{new_file_path}'")

# Function to handle drag-and-drop file selection
def handle_drop(event):
    files = root.tk.splitlist(event.data)  # Extract the list of dropped files
    for file in files:
        filename = os.path.basename(file)
        listbox.insert(tk.END, filename)
        filename_to_path[filename] = file  # Update the mapping

def clear_listbox():
    listbox.delete(0, tk.END)  # Clear the listbox
    filename_to_path.clear() 

# Set up the GUI
root = TkinterDnD.Tk()
root.title("Bulk Rename Utility for Office")

frame = tk.Frame(root)
frame.pack(pady=20, fill=tk.BOTH, expand=True)  # Ensure the frame expands

listbox = tk.Listbox(frame)  # Associate the Listbox with the frame
listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)  # Make sure it fills the frame
listbox.drop_target_register('DND_Files')
listbox.dnd_bind('<<Drop>>', handle_drop)


scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

listbox.config(yscrollcommand=scrollbar.set)

mode_var = tk.BooleanVar()
mode_var.set(False)  # Default mode

mode_frame = tk.Frame(root)
mode_frame.pack(pady=10)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# Mode 1 selection Radiobutton in mode_frame
mode_radio1 = tk.Radiobutton(mode_frame, text="รายชื่อนักเรียน", variable=mode_var, value=1)
mode_radio1.pack(side=tk.LEFT, padx=5)

# Mode 2 selection Radiobutton in mode_frame
mode_radio2 = tk.Radiobutton(mode_frame, text="เกรดเฉลี่ย", variable=mode_var, value=2)
mode_radio2.pack(side=tk.LEFT, padx=5)

mode_var.set(1)

add_button = tk.Button(button_frame, text="Add Files", command=select_files)
add_button.pack(side=tk.LEFT, padx=5)

clear_button = tk.Button(button_frame, text="Clear", command=clear_listbox)
clear_button.pack(side=tk.LEFT, padx=5)

rename_button = tk.Button(root, text="Rename Files", command=rename_files)
rename_button.pack(pady=5)

root.mainloop()