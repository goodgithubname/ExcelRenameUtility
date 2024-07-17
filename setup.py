from cx_Freeze import setup, Executable
import glob

png_files = glob.glob('*.png')

build_exe_options = {
    "packages": ["os", "tkinter", "openpyxl", "tkinterdnd2"],
    "include_files": [("forest-dark.tcl", "forest-dark.tcl")] + png_files,  # Modified line
}

base="Win32GUI"

setup(
    name="Excel Bulk Renam Utility",
    version="0.1",
    description="Rename excel files based on content inside the files",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=base)]
)