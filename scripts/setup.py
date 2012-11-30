from cx_Freeze import setup, Executable

exe = Executable(
    script="filepicker.pyw",
    base="Win32GUI",
    )

setup(
        name = "FilePicker",
        version = "0.1",
        description = "Dictionary raw data formatter helper",
        executables = [exe])
