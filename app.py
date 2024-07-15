import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import openpyxl
import os
from tkinterdnd2 import TkinterDnD

filename_to_path = {}

# Function to handle file selection
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    for file in files:
        filename = os.path.basename(file) 
        treeView.insert("", tk.END, text=filename, values=(filename,))
        filename_to_path[filename] = file

def rename_files():
        for filename in treeView.get(0, tk.END):
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
        treeView.insert("", tk.END, text=filename, values=(filename,))
        filename_to_path[filename] = file  # Update the mapping

def clear_listbox():
    for item in treeView.get_children():
        treeView.delete(item)
    filename_to_path.clear() 

def apply_dark_theme(widget):
    widget_type = widget.winfo_class()
    if widget_type in ["Frame", "Tk", "Toplevel"]:
        widget.config(bg=dark_bg)
    elif widget_type == "Listbox":
        widget.config(bg=dark_bg, fg=dark_fg, highlightcolor=accent_color, selectbackground=accent_color)
    elif widget_type == "Button":
        widget.config(bg=dark_bg, fg=dark_fg, activebackground=accent_color)
    # Add more conditions for other widget types as needed

#Colour palette
dark_bg = "#333333"  # Dark background color
dark_fg = "#ffffff"  # Light text color
accent_color = "#0078D7"  # Accent color for buttons or highlights


# Set up the GUI
root = TkinterDnD.Tk()
root.title("Bulk Rename Utility for Office")

style = ttk.Style(root)
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")
 
frame = ttk.Frame(root)
frame.pack(fill='both', expand=True)

widgets_frame = ttk.LabelFrame(frame, text="Configuration", height=10)
widgets_frame.grid(row=0, column=0, padx=10, pady= 10, sticky="nsew")
frame.rowconfigure(0, weight=1)  # Allow widgets_frame to expand vertically


# Place the status_combobox at the top
status_combobox = ttk.Combobox(widgets_frame, values=["รายชื่อนักเรียน", "เกรดเฉลี่ย"], state="readonly")
status_combobox.current(0)
status_combobox.grid(row=0, column=0, padx=5, pady = (0,5) ,sticky="ew")
widgets_frame.columnconfigure(0, weight=1)  # Make the combobox expand horizontally

# Move the button_frame below the status_combobox
button_frame = tk.Frame(widgets_frame)
button_frame.grid(row=2, pady= 5, column=0, sticky="ew")
button_frame.columnconfigure(0, weight=1)
button_frame.columnconfigure(1, weight=1)


separator = ttk.Separator(widgets_frame)
separator.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

# Adjust the row numbers for add_button and clear_button to be sequential within button_frame
add_button = ttk.Button(button_frame, text="เพิ่มไฟล์", command=select_files)
add_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

clear_button = ttk.Button(button_frame, text="เคลียร์", command=clear_listbox)
clear_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

rename_button = ttk.Button(button_frame, text="เปลี่ยนชื่อไฟล์", command=rename_files)
rename_button.grid(row=1, column=0, padx =5, columnspan=2, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

treeView = ttk.Treeview(treeFrame, columns=(1,2,3), show="", height="13")
treeView.pack()
treeView.drop_target_register('DND_Files')
treeView.dnd_bind('<<Drop>>', handle_drop)
treeFrame.columnconfigure(0, weight=1)

# boxScroll = ttk.Scrollbar(listbox)
# boxScroll.pack(side="right", fill="y")

# # # Apply dark theme to the main window
# apply_dark_theme(root)

# frame = tk.Frame(root)
# frame.pack(pady=20, fill=tk.BOTH, expand=True)
# apply_dark_theme(frame)  # Apply dark theme to the frame




# scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
# scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# listbox.config(yscrollcommand=scrollbar.set)

# mode_var = tk.BooleanVar()
# mode_var.set(False)  # Default mode

# mode_frame = tk.Frame(root)
# mode_frame.pack(pady=10)


# # Mode 1 selection Radiobutton in mode_frame
# mode_radio1 = tk.Radiobutton(mode_frame, text="รายชื่อนักเรียน", variable=mode_var, value=1)
# mode_radio1.pack(side=tk.LEFT, padx=5)

# # Mode 2 selection Radiobutton in mode_frame
# mode_radio2 = tk.Radiobutton(mode_frame, text="เกรดเฉลี่ย", variable=mode_var, value=2)
# mode_radio2.pack(side=tk.LEFT, padx=5)

# mode_var.set(1)





root.mainloop()