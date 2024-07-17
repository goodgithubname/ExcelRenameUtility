from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["os", "tkinter", "openpyxl", "tkinterdnd2"],
    "include_files": [("path/to/forest-dark.tcl", "forest-dark.tcl")],  # Add this line
}

base="Win32GUI"

setup(
    name="Excel Bulk Renam Utility",
    version="0.1",
    description="Rename excel files based on content inside the files",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=base)]
)