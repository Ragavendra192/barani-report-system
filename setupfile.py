from cx_Freeze import setup, Executable
import sys
import os

base = None

# Make sure Flask templates are included
include_files = [
    ("templates", "templates")
]

build_exe_options = {
    "packages": [
        "flask",
        "pyodbc",
        "pandas",
        "openpyxl"
    ],
    "include_files": include_files,
    "excludes": ["tkinter"]
}

setup(
    name="BaraniReportSystem",
    version="1.0",
    description="Flask Reporting System",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=base)]
)
