from cx_Freeze import setup, Executable

exe = Executable(
    script="__init__.pyw",
    base="Win32GUI")

build_options = {"packages": ["xlrd"]}

setup(
        name = "FilePicker",
        version = "0.1",
        description = "Dictionary raw data formatter helper",
        options = {"build_exe": build_options},
        executables = [exe])