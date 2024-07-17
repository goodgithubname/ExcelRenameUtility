import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import openpyxl
import os
from tkinterdnd2 import TkinterDnD
from tkinter import messagebox

filename_to_path = {}

# Function to handle file selection
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    for file in files:
        filename = os.path.basename(file) 
        treeView.insert("", tk.END, text=filename, values=(filename,))
        filename_to_path[filename] = file

def rename_files():
        if treeView.get_children() == ():
            return
        for filename in treeView.get_children():
            original_text = treeView.item(filename, 'text') 
            file_path = filename_to_path[treeView.item(filename, 'text')]  # Retrieve the full path
            workbook = openpyxl.load_workbook(file_path)
            if status_combobox.get() == "รายชื่อนักเรียน":
                sheet = workbook.worksheets[0]
                class_cell = sheet['B5'].value.split(" ")
                if class_cell[0].startswith("ม."):
                    fileName = class_cell[0].replace("ม.", "m")
                elif class_cell[0].startswith("ป."):
                    fileName = class_cell[0].replace("ป.", "p")
                fileName = fileName.replace("/", "-")
            elif  status_combobox.get() == "ผลการเรียน":
                sheet = workbook.worksheets[0]
                class_cell = sheet['E2'].value.split(" ")
                year_cell = sheet['A2'].value.split(" ")
                if class_cell[1].startswith("ป."):
                    fileName = f"p-{class_cell[3]} {year_cell[4]}"
                elif class_cell[1].startswith("ม."):
                    fileName = f"m-{class_cell[3]} term{year_cell[1]}-{year_cell[3]}"
 
            if not fileName.endswith('.xlsx'):
                fileName += '.xlsx'
            directory = os.path.dirname(file_path)
            new_file_path = os.path.join(directory, fileName)

            if os.path.exists(new_file_path):
                messagebox.showwarning("Operation Aborted", f"File {fileName} already exists. Operation aborted to prevent overwriting.")
                continue  # Skip the renaming for this file

            # Rename the file
            os.rename(file_path, new_file_path)
            # Remove the renamed file from filename_to_path dictionary
            del filename_to_path[original_text]

            # Remove the item from treeView
            treeView.delete(filename)

            print(f"Renamed '{file_path}' to '{new_file_path}'")
        messagebox.showinfo("เสร็จสิ้น", "เปลี่ยนชื่อไฟล์เสร็จสิ้น")


# Function to handle drag-and-drop file selection
def handle_drop(event):
    files = root.tk.splitlist(event.data)  # Extract the list of dropped files
    for file in files:
        filename = os.path.basename(file)
        _, ext = os.path.splitext(filename)  # Extract the file extension
        if ext.lower() == '.xlsx':  # Check if the file is an Excel file
            treeView.insert("", tk.END, text=filename, values=(filename,))
            filename_to_path[filename] = file  # Update the mapping

def clear_listbox():
    for item in treeView.get_children():
        treeView.delete(item)
    filename_to_path.clear() 

# Set up the GUI
root = TkinterDnD.Tk()
root.title("Bulk Rename Utility for Office")

style = ttk.Style(root)
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")
 
frame = ttk.Frame(root)
frame.pack(fill='both', expand=True)

widgets_frame = ttk.LabelFrame(frame, text="การตั้งค่า", height=10)
widgets_frame.grid(row=0, column=0, padx=10, pady= 10, sticky="nsew")
frame.rowconfigure(0, weight=1)  # Allow widgets_frame to expand vertically
frame.columnconfigure(0, weight=1)  # Allow widgets_frame to expand horizontally

# Place the status_combobox at the top
status_combobox = ttk.Combobox(widgets_frame, values=["รายชื่อนักเรียน", "ผลการเรียน"], state="readonly")
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

gif_image = tk.PhotoImage(file="raccoon-dance.gif")

gif_label = tk.Label(widgets_frame, image=gif_image)
gif_label.image = gif_image

gif_label.grid(row=3, column=0, padx=5, pady=5)

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)

treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

treeView = ttk.Treeview(treeFrame, columns=(1,2,3), show="", height="13")
treeView.pack()
treeView.drop_target_register('DND_Files')
treeView.dnd_bind('<<Drop>>', handle_drop)
treeFrame.columnconfigure(0, weight=1)

root.mainloop()